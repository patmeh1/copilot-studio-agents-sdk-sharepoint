"""SharePoint REST helpers executed inside a signed-in SharePoint page (uses the browser's auth cookies)."""
import json

JS_REST = """async ({method, url, body, site}) => {
  const base = site || location.origin;
  let digest = null;
  if (method !== 'GET') {
    const ci = await fetch(base + '/_api/contextinfo', {method: 'POST', headers: {Accept: 'application/json;odata=nometadata'}});
    digest = (await ci.json()).FormDigestValue;
  }
  const headers = {Accept: 'application/json;odata=nometadata'};
  if (digest) headers['X-RequestDigest'] = digest;
  if (body !== null && body !== undefined) headers['Content-Type'] = 'application/json;odata=nometadata';
  const res = await fetch(url.startsWith('http') ? url : base + url, {method, headers, body: body === null || body === undefined ? undefined : JSON.stringify(body)});
  const text = await res.text();
  let data = null; try { data = text ? JSON.parse(text) : null; } catch (e) { data = text; }
  return {status: res.status, data};
}"""


def rest(page, method, url, body=None, site=None):
    r = page.evaluate(JS_REST, {"method": method, "url": url, "body": body, "site": site})
    return r["status"], r["data"]


def must(page, method, url, body=None, site=None, ok=(200, 201, 204)):
    status, data = rest(page, method, url, body, site)
    if status not in ok:
        raise RuntimeError(f"{method} {url} -> {status}: {json.dumps(data)[:800]}")
    return data
