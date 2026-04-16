<#
.SYNOPSIS
    SharePoint Bridge Server (PowerShell) - local HTTP proxy for uploading files to SharePoint.

.DESCRIPTION
    Runs on port 8081. The browser app POSTs a file here, and this server
    uploads it to SharePoint via Microsoft Graph API using MSAL.PS auth
    (no admin consent required).

    Requires: MSAL.PS module (Install-Module MSAL.PS -Scope CurrentUser)

.USAGE
    .\sp_bridge.ps1
    .\sp_bridge.ps1 -Port 8082
#>
param(
    [int]$Port = 8081
)

$ErrorActionPreference = 'Stop'

# -- Constants ----------------------------------------------------------------

$CLIENT_ID = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'   # Azure CLI public client
$AUTHORITY  = 'https://login.microsoftonline.com/organizations'
$SCOPES     = @('https://graph.microsoft.com/.default')
$GRAPH      = 'https://graph.microsoft.com/v1.0'
$CACHE_FILE = Join-Path $PSScriptRoot '_token_cache.bin'

# -- Auth ---------------------------------------------------------------------

$script:msalToken = $null

function Get-GraphToken {
    # Try silent first if we already have a token
    if ($script:msalToken) {
        try {
            $result = Get-MsalToken -ClientId $CLIENT_ID `
                -Authority $AUTHORITY `
                -Scopes $SCOPES `
                -Silent `
                -ForceRefresh:$false
            if ($result.AccessToken) {
                $script:msalToken = $result
                return $result.AccessToken
            }
        } catch {
            # Fall through to interactive
        }
    }

    # Interactive login
    Write-Host '[bridge] Opening browser for sign-in ...'
    $result = Get-MsalToken -ClientId $CLIENT_ID `
        -Authority $AUTHORITY `
        -Scopes $SCOPES `
        -Interactive
    if (-not $result.AccessToken) {
        throw 'Authentication failed - no access token received'
    }
    $script:msalToken = $result
    return $result.AccessToken
}

# -- SharePoint URL parser ----------------------------------------------------

function Parse-SharePointUrl {
    param([string]$Url)
    $Url = $Url.Trim().TrimEnd('/')
    if ($Url -match '^https?://([^/]+\.sharepoint\.com)(/sites/[^/]+)/([^/?#]+)$') {
        return @{
            Host     = $Matches[1]
            SitePath = $Matches[2]
            LibName  = $Matches[3]
        }
    }
    throw "Invalid SharePoint URL: $Url`nExpected: https://<tenant>.sharepoint.com/sites/<SiteName>/<DocLibName>"
}

# -- Graph helpers ------------------------------------------------------------

function Resolve-Drive {
    param([string]$Token, [string]$SpHost, [string]$SitePath, [string]$LibName)

    $headers = @{ Authorization = "Bearer $Token" }

    # Get site ID
    $siteUrl = "$GRAPH/sites/${SpHost}:${SitePath}"
    $site = Invoke-RestMethod -Uri $siteUrl -Headers $headers -Method Get
    $siteId = $site.id

    # List drives and find matching library
    $drivesUrl = "$GRAPH/sites/$siteId/drives"
    $drives = Invoke-RestMethod -Uri $drivesUrl -Headers $headers -Method Get

    foreach ($d in $drives.value) {
        if ($d.name -eq $LibName) {
            return $d.id
        }
    }

    $names = ($drives.value | ForEach-Object { $_.name }) -join ', '
    throw "Drive '$LibName' not found. Available: $names"
}

function Upload-File {
    param([string]$Token, [string]$DriveId, [string]$Filename, [byte[]]$Data)

    $headers = @{
        Authorization  = "Bearer $Token"
        'Content-Type' = 'application/octet-stream'
    }
    $url = "$GRAPH/drives/$DriveId/root:/${Filename}:/content"
    $result = Invoke-RestMethod -Uri $url -Headers $headers -Method Put -Body $Data
    if ($result.webUrl) { return $result.webUrl }
    if ($result.id)     { return $result.id }
    return 'ok'
}

# -- JSON response helper -----------------------------------------------------

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [int]$StatusCode,
        [hashtable]$Body
    )
    $json = $Body | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    $Response.StatusCode = $StatusCode
    $Response.ContentType = 'application/json'
    $Response.ContentLength64 = $bytes.Length
    $Response.Headers.Add('Access-Control-Allow-Origin', '*')
    $Response.Headers.Add('Access-Control-Allow-Methods', 'POST, OPTIONS, GET')
    $Response.Headers.Add('Access-Control-Allow-Headers', 'Content-Type, X-SP-Url, X-Filename')
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

# -- Main Server --------------------------------------------------------------

# Check MSAL.PS is available
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Write-Host '[bridge] ERROR: MSAL.PS module not installed.'
    Write-Host '[bridge] Run: Install-Module MSAL.PS -Scope CurrentUser'
    exit 1
}
Import-Module MSAL.PS -ErrorAction Stop

# Pre-authenticate
Write-Host '[bridge] Authenticating ...'
try {
    $null = Get-GraphToken
    Write-Host '[bridge] Authenticated OK'
} catch {
    Write-Host "[bridge] Auth failed: $_"
    exit 1
}

# Start HTTP listener
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://127.0.0.1:${Port}/")
$listener.Start()

Write-Host "[bridge] Listening on http://127.0.0.1:${Port}"
Write-Host '[bridge] POST /upload  - upload file to SharePoint'
Write-Host '[bridge] GET  /health  - health check'
Write-Host '[bridge] Press Ctrl+C to stop'

try {
    while ($listener.IsListening) {
        $context  = $listener.GetContext()
        $request  = $context.Request
        $response = $context.Response

        $path   = $request.Url.AbsolutePath
        $method = $request.HttpMethod

        try {
            # CORS preflight
            if ($method -eq 'OPTIONS') {
                Send-JsonResponse -Response $response -StatusCode 200 -Body @{ ok = $true }
                continue
            }

            # Health check
            if ($method -eq 'GET' -and $path -eq '/health') {
                Send-JsonResponse -Response $response -StatusCode 200 -Body @{
                    ok      = $true
                    service = 'sp-bridge-ps'
                }
                continue
            }

            # Upload
            if ($method -eq 'POST' -and $path -eq '/upload') {
                $spUrl    = $request.Headers['X-SP-Url']
                $filename = $request.Headers['X-Filename']
                if (-not $filename) { $filename = 'upload.xlsx' }

                if (-not $spUrl) {
                    Send-JsonResponse -Response $response -StatusCode 400 -Body @{
                        ok    = $false
                        error = 'Missing X-SP-Url header'
                    }
                    continue
                }

                # Read body
                $ms = [System.IO.MemoryStream]::new()
                $request.InputStream.CopyTo($ms)
                $bodyBytes = $ms.ToArray()
                $ms.Dispose()

                if ($bodyBytes.Length -eq 0) {
                    Send-JsonResponse -Response $response -StatusCode 400 -Body @{
                        ok    = $false
                        error = 'Empty body'
                    }
                    continue
                }

                $parsed = Parse-SharePointUrl -Url $spUrl
                Write-Host "[bridge] Uploading '$filename' to $($parsed.Host)$($parsed.SitePath)/$($parsed.LibName) ..."

                $token   = Get-GraphToken
                $driveId = Resolve-Drive -Token $token -SpHost $parsed.Host -SitePath $parsed.SitePath -LibName $parsed.LibName
                $webUrl  = Upload-File -Token $token -DriveId $driveId -Filename $filename -Data $bodyBytes

                Write-Host "[bridge] OK Uploaded: $webUrl"
                Send-JsonResponse -Response $response -StatusCode 200 -Body @{
                    ok       = $true
                    webUrl   = $webUrl
                    filename = $filename
                }
                continue
            }

            # 404
            Send-JsonResponse -Response $response -StatusCode 404 -Body @{
                ok    = $false
                error = 'Not found'
            }
        } catch {
            Write-Host "[bridge] Error: $_"
            try {
                Send-JsonResponse -Response $response -StatusCode 500 -Body @{
                    ok    = $false
                    error = "$_"
                }
            } catch {
                # Response may already be closed
            }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
    Write-Host '[bridge] Stopped'
}
