# sp-submit

Drop-in SharePoint file upload for any web app.

Your browser app sends file bytes to a small local bridge server, which authenticates with Microsoft 365 and uploads to SharePoint via the Graph API. No admin consent, no app registration, no server deployment.

---

## What's In the Box

| File | What it does |
|---|---|
| `sp-submit.js` | Browser ES6 module — your app imports this |
| `sp_bridge.py` | Bridge server (Python) |
| `sp_bridge.ps1` | Bridge server (PowerShell) — pick whichever you prefer |

---

## Requirements

| Requirement | Details |
|---|---|
| A SharePoint Online site | Any site you have write access to |
| Python **or** PowerShell | You only need one |
| Python route | `pip install msal requests` |
| PowerShell route | `Install-Module MSAL.PS -Scope CurrentUser` |
| A modern browser | Chrome, Edge, Firefox, Safari — anything with ES6 module support |

### What you do NOT need

- No Azure app registration
- No client secret or certificate
- No admin consent
- No server deployment — the bridge runs on `localhost`
- No frameworks — works with vanilla JS, React, Vue, whatever

---

## Quick Start

### 1. Install dependencies (pick one)

**Python:**
```bash
pip install msal requests
```

**PowerShell:**
```powershell
Install-Module MSAL.PS -Scope CurrentUser
```

### 2. Start the bridge server

**Python:**
```bash
python sp_bridge.py
```

**PowerShell:**
```powershell
.\sp_bridge.ps1
```

On first run a browser window opens for Microsoft sign-in. After that, tokens are cached and you won't be prompted again until they expire.

### 3. Add to your app

```html
<button id="upload-btn">Upload to SharePoint</button>

<script type="module">
import SPSubmit from './sp-submit.js';

const sp = new SPSubmit({
  // Required — return the SharePoint doc library URL
  getSpUrl: () => 'https://contoso.sharepoint.com/sites/Finance/SharedDocuments',

  // Required — return the file to upload
  getFile: async () => {
    const blob = await generateMyReport();
    return {
      bytes: new Uint8Array(await blob.arrayBuffer()),
      filename: 'Q4-Report.xlsx'
    };
  },

  // Optional — get notified about progress
  onStatus: (message, type) => {
    // type is 'info', 'success', or 'error'
    console.log(`[${type}] ${message}`);
  }
});

// Wire it to a button
sp.attachButton(document.getElementById('upload-btn'));
</script>
```

That's it.

---

## API Reference

### `new SPSubmit(options)`

| Option | Type | Required | Default | Description |
|---|---|---|---|---|
| `getSpUrl` | `() => string` | **Yes** | — | Returns the SharePoint document library URL |
| `getFile` | `async () => { bytes, filename }` | **Yes** | — | Returns the file bytes and filename |
| `onStatus` | `(msg, type) => void` | No | no-op | Status callback. `type` is `'info'`, `'success'`, or `'error'` |
| `bridgeUrl` | `string` | No | `http://127.0.0.1:8081` | Bridge server base URL |
| `healthTimeout` | `number` | No | `3000` | Bridge health-check timeout in ms |

### Methods

| Method | Returns | Description |
|---|---|---|
| `submit()` | `Promise<boolean>` | Upload the file. Returns `true` on success. |
| `attachButton(element)` | `this` | Add a click handler to an existing button |
| `createButton(container, opts?)` | `this` | Create a `<button>` and append it to a container |

`createButton` options: `{ text: 'Submit to SharePoint', className: '' }`

---

## SharePoint URL Format

The URL must point to a **document library** inside a **site**, using this exact format:

```
https://<tenant>.sharepoint.com/sites/<SiteName>/<DocLibraryName>
```

### Examples

| URL | Valid? | Why |
|---|---|---|
| `https://contoso.sharepoint.com/sites/Finance/SharedDocuments` | Yes | Standard site + default doc library |
| `https://contoso.sharepoint.com/sites/HR/PolicyDocs` | Yes | Custom doc library name |
| `https://contoso.sharepoint.com/sites/Finance` | **No** | Missing the document library name |
| `https://contoso.sharepoint.com/Finance/SharedDocuments` | **No** | Missing `/sites/` |
| `https://contoso.sharepoint.com/:f:/sites/Finance/SharedDocuments` | **No** | This is a sharing link, not a direct URL |

### How to find the correct URL

1. Open your SharePoint site in a browser
2. Navigate into the document library where files should go
3. Look at the address bar — it will be something like:
   ```
   https://contoso.sharepoint.com/sites/Finance/SharedDocuments/Forms/AllItems.aspx
   ```
4. Take everything **before** `/Forms/...` — that's your URL:
   ```
   https://contoso.sharepoint.com/sites/Finance/SharedDocuments
   ```

### Common mistakes

- **Trailing slash** — `https://.../SharedDocuments/` — this is fine, the tool strips it
- **Extra path segments** — `https://.../SharedDocuments/Subfolder` — not supported; omit subfolder
- **Sharing links** — URLs with `/:f:/` or `/:x:/` are sharing links, not direct URLs
- **Personal OneDrive** — `https://contoso-my.sharepoint.com/...` — this tool is for SharePoint sites, not OneDrive

---

## Bridge Server Details

The bridge listens on `http://127.0.0.1:8081` (localhost only — not exposed to the network).

| Endpoint | Method | Purpose |
|---|---|---|
| `/health` | GET | Check if bridge is running. Returns `{ "ok": true }` |
| `/upload` | POST | Upload a file to SharePoint |

### POST /upload

**Headers:**
- `X-SP-Url` — the SharePoint doc library URL (required)
- `X-Filename` — destination filename (defaults to `upload.xlsx`)
- `Content-Type: application/octet-stream`

**Body:** raw file bytes

**Response:**
```json
{ "ok": true, "webUrl": "https://...", "filename": "MyFile.xlsx" }
```

### Authentication

The bridge uses the **Azure CLI public client ID** (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) — a Microsoft first-party app that is pre-consented in every tenant. This means:

- No app registration needed
- No admin consent needed
- Works with any Microsoft 365 org account
- You sign in through your normal browser
- Tokens are cached locally in `_token_cache.bin`

### Changing the port

**Python:**
```python
# Edit PORT = 8081 at the top of sp_bridge.py
```

**PowerShell:**
```powershell
.\sp_bridge.ps1 -Port 8082
```

**Browser side:**
```js
new SPSubmit({ bridgeUrl: 'http://127.0.0.1:8082', ... });
```

---

## .gitignore

Add this to your `.gitignore` — the token cache should never be committed:

```
_token_cache.bin
```

---

## License

MIT
