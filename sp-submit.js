/**
 * SPSubmit — Standalone SharePoint file submission module.
 *
 * Drop this file into any web app. Provide a config object and call submit().
 * Requires a running bridge server (sp_bridge.py or sp_bridge.ps1) on localhost.
 *
 * Usage:
 *   import SPSubmit from './sp-submit.js';
 *
 *   const sp = new SPSubmit({
 *     bridgeUrl:  'http://127.0.0.1:8081',          // bridge server base URL
 *     getSpUrl:   () => myConfig.sharepointUrl,      // returns SP doc-lib URL
 *     getFile:    async () => ({                      // returns file to upload
 *       bytes: arrayBufferOrUint8Array,
 *       filename: 'Report.xlsx'
 *     }),
 *     onStatus:   (msg, type) => showToast(msg, type) // 'info' | 'success' | 'error'
 *   });
 *
 *   // Option A — attach to an existing button
 *   sp.attachButton(document.getElementById('my-btn'));
 *
 *   // Option B — create a button inside a container
 *   sp.createButton(document.getElementById('toolbar'), {
 *     text: 'Submit to SharePoint',
 *     className: 'sp-submit-btn'
 *   });
 *
 *   // Option C — call programmatically
 *   await sp.submit();
 */

const DEFAULTS = {
  bridgeUrl: 'http://127.0.0.1:8081',
  healthTimeout: 3000,
};

export default class SPSubmit {

  #bridgeUrl;
  #healthTimeout;
  #getSpUrl;
  #getFile;
  #onStatus;
  #busy = false;

  /**
   * @param {Object} opts
   * @param {string}   [opts.bridgeUrl='http://127.0.0.1:8081']  Bridge server URL
   * @param {number}   [opts.healthTimeout=3000]                  Health-check timeout (ms)
   * @param {Function} opts.getSpUrl   — () => string            SharePoint doc-lib URL
   * @param {Function} opts.getFile    — async () => { bytes, filename }
   * @param {Function} [opts.onStatus] — (message, type) => void  type: 'info'|'success'|'error'
   */
  constructor(opts = {}) {
    if (typeof opts.getSpUrl !== 'function')
      throw new Error('SPSubmit: getSpUrl callback is required');
    if (typeof opts.getFile !== 'function')
      throw new Error('SPSubmit: getFile callback is required');

    this.#bridgeUrl     = (opts.bridgeUrl || DEFAULTS.bridgeUrl).replace(/\/+$/, '');
    this.#healthTimeout = opts.healthTimeout ?? DEFAULTS.healthTimeout;
    this.#getSpUrl      = opts.getSpUrl;
    this.#getFile       = opts.getFile;
    this.#onStatus      = typeof opts.onStatus === 'function' ? opts.onStatus : () => {};
  }

  /* ── Public API ────────────────────────────────────────────── */

  /** Attach click handler to an existing button element. */
  attachButton(btnEl) {
    if (!btnEl) return;
    btnEl.addEventListener('click', () => this.submit());
    return this;
  }

  /** Create a new <button> inside `container` and wire it up. */
  createButton(container, { text = 'Submit to SharePoint', className = '' } = {}) {
    if (!container) return this;
    const btn = document.createElement('button');
    btn.textContent = text;
    if (className) btn.className = className;
    btn.addEventListener('click', () => this.submit());
    container.appendChild(btn);
    return this;
  }

  /**
   * Submit file to SharePoint via the bridge server.
   * Returns true on success, false on failure.
   */
  async submit() {
    if (this.#busy) {
      this.#onStatus('A submission is already in progress.', 'info');
      return false;
    }
    this.#busy = true;

    try {
      // 1. Validate SP URL
      const spUrl = (this.#getSpUrl() || '').trim();
      if (!spUrl) {
        this.#onStatus('SharePoint Document Library URL is not configured.', 'error');
        return false;
      }

      // 2. Health-check bridge
      try {
        const res = await fetch(`${this.#bridgeUrl}/health`, {
          signal: AbortSignal.timeout(this.#healthTimeout),
        });
        if (!res.ok) throw new Error('Bridge returned ' + res.status);
      } catch {
        this.#onStatus(
          'SharePoint bridge server is not running. Start sp_bridge.py or sp_bridge.ps1 first.',
          'error'
        );
        return false;
      }

      // 3. Get file from caller
      this.#onStatus('Preparing file for upload...', 'info');
      const file = await this.#getFile();
      if (!file || !file.bytes) {
        this.#onStatus('getFile() did not return a file.', 'error');
        return false;
      }
      const bytes    = file.bytes instanceof ArrayBuffer ? new Uint8Array(file.bytes) : file.bytes;
      const filename = file.filename || 'upload.xlsx';

      // 4. Upload
      this.#onStatus('Submitting to SharePoint...', 'info');
      const resp = await fetch(`${this.#bridgeUrl}/upload`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/octet-stream',
          'X-SP-Url':     spUrl,
          'X-Filename':   filename,
        },
        body: bytes,
      });

      const result = await resp.json();
      if (result.ok) {
        this.#onStatus(`Submitted to SharePoint \u2014 ${result.filename}`, 'success');
        return true;
      } else {
        this.#onStatus('SharePoint upload failed: ' + result.error, 'error');
        return false;
      }
    } catch (e) {
      console.error('SPSubmit error:', e);
      this.#onStatus('Submit failed: ' + e.message, 'error');
      return false;
    } finally {
      this.#busy = false;
    }
  }
}
