// Microsoft Graph — app-only, for the scheduled sync.
//
// The rest of the ART suite uses delegated auth: a person signs in and Graph acts as
// them. A scheduled job has no person, so this uses client credentials instead. That
// needs an application permission (Sites.Selected preferred over Sites.ReadWrite.All,
// since it grants this one site rather than the whole tenant) and admin consent.
//
// The wall board itself still uses ordinary delegated MSAL sign-in, exactly like the
// other tools. Only this unattended half needs app-only.

export class Graph {
  constructor({ tenantId, clientId, clientSecret }) {
    for (const [k, v] of Object.entries({ tenantId, clientId, clientSecret })) {
      if (!v) throw new Error(`Graph: ${k} is missing.`);
    }
    Object.assign(this, { tenantId, clientId, clientSecret });
    this.token = null;
    this.tokenExpires = 0;
    this.ids = {};
  }

  async #accessToken() {
    if (this.token && Date.now() < this.tokenExpires - 60_000) return this.token;
    const res = await fetch(`https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    });
    if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
    const d = await res.json();
    this.token = d.access_token;
    this.tokenExpires = Date.now() + (d.expires_in ?? 3600) * 1000;
    return this.token;
  }

  async req(method, path, body) {
    // SharePoint throttles bursts -- a first run writes 125 rows back to back. 429 and 503
    // come with Retry-After; honour it a few times before giving up.
    for (let attempt = 1; ; attempt++) {
      const res = await fetch('https://graph.microsoft.com/v1.0' + path, {
        method,
        headers: {
          Authorization: 'Bearer ' + (await this.#accessToken()),
          'Content-Type': 'application/json',
          // SharePoint refuses $filter on unindexed columns without this.
          Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly',
        },
        ...(body ? { body: JSON.stringify(body) } : {}),
      });
      if ((res.status === 429 || res.status === 503) && attempt < 5) {
        const wait = Math.min(60, Number(res.headers.get('retry-after')) || attempt * 5);
        console.log(`  (SharePoint busy, waiting ${wait}s)`);
        await new Promise(r => setTimeout(r, wait * 1000));
        continue;
      }
      if (!res.ok) throw new Error(`${method} ${path} -> ${res.status} ${(await res.text()).slice(0, 300)}`);
      return res.status === 204 ? null : res.json();
    }
  }

  // A file in the site's document library, as text. Used for the RepairQ stock export,
  // which is far too big for a list field -- the Orders page writes it, this reads it.
  async fileText(path) {
    const site = this.ids.site;
    if (!site) throw new Error('Graph: site id not resolved yet');
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${site}/drive/root:/${path.split('/').map(encodeURIComponent).join('/')}:/content`,
      { headers: { Authorization: 'Bearer ' + (await this.#accessToken()) } });
    if (res.status === 404) return null;                  // nothing uploaded yet
    if (!res.ok) throw new Error(`GET ${path} -> ${res.status} ${(await res.text()).slice(0, 200)}`);
    return res.text();
  }

  async siteId(hostname, sitePath) {
    if (this.ids.site) return this.ids.site;
    const d = await this.req('GET', `/sites/${hostname}:${sitePath}`);
    this.ids.site = d.id;
    return d.id;
  }

  async listId(name) {
    if (this.ids[name]) return this.ids[name];
    const site = this.ids.site;
    const d = await this.req('GET', `/sites/${site}/lists/${encodeURIComponent(name)}`);
    this.ids[name] = d.id;
    return d.id;
  }

  async items(listName) {
    const lid = await this.listId(listName);
    const out = [];
    let url = `/sites/${this.ids.site}/lists/${lid}/items?expand=fields&$top=999`;
    for (;;) {
      const d = await this.req('GET', url);
      out.push(...(d.value ?? []));
      const next = d['@odata.nextLink'];
      if (!next) break;
      url = next.replace('https://graph.microsoft.com/v1.0', '');
    }
    return out;
  }

  async create(listName, fields) {
    const lid = await this.listId(listName);
    return this.req('POST', `/sites/${this.ids.site}/lists/${lid}/items`, { fields });
  }

  async update(listName, itemId, fields) {
    const lid = await this.listId(listName);
    return this.req('PATCH', `/sites/${this.ids.site}/lists/${lid}/items/${itemId}/fields`, fields);
  }

  async remove(listName, itemId) {
    const lid = await this.listId(listName);
    return this.req('DELETE', `/sites/${this.ids.site}/lists/${lid}/items/${itemId}`);
  }
}
