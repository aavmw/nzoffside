function apiFetch_(path, opts) {
  const BASE = CFG.apiBase;   // e.g. "https://<ngrok>.ngrok.io"
  const KEY  = CFG.apiKey;

  // normalize base (remove trailing /)
  const base = BASE.replace(/\/+$/, '');

  // normalize path (ensure leading /, no duplicate slashes)
  const rawPath = path.startsWith('/') ? path : '/' + path;

  // always prefix with /v1
  const pth = '/v1' + rawPath.replace(/^\/+/, '/');  

  const url = base + pth;

  const headers = {
    'X-Api-Key': KEY,
    'ngrok-skip-browser-warning': '1'
  };

  const payload = opts && opts.body ? JSON.stringify(opts.body) : undefined;
  const method  = (opts && opts.method) || 'get';

  let lastErr, res;
  for (let i = 0; i < 3; i++) {
    res = UrlFetchApp.fetch(url, {
      method,
      contentType: 'application/json',
      headers,
      payload,
      muteHttpExceptions: true,
      followRedirects: true
    });

    const code = res.getResponseCode();
    const txt  = res.getContentText();
    console.log({ url, code, preview: txt.slice(0, 160) });

    if (code >= 200 && code < 300) {
      try {
        const json = JSON.parse(txt);
        if (json && typeof json === 'object') return json;
        throw new Error('Non-JSON success payload');
      } catch (e) {
        throw new Error('Expected JSON but got: ' + txt.slice(0, 500));
      }
    }

    if (code === 429 || (code >= 500 && code < 600)) {
      Utilities.sleep(500 * (i + 1));
      lastErr = 'API ' + code + ': ' + txt.slice(0, 500);
      continue;
    }

    throw new Error('API ' + code + ': ' + txt.slice(0, 500));
  }

  throw new Error(lastErr || 'Unknown API error');
}
