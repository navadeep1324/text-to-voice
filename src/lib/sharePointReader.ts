export interface SharePointConfig {
  tenant_id: string;
  client_id: string;
  client_secret: string;
  host: string;        // "klezaio.sharepoint.com"
  site_path: string;   // "/sites/Kleza"
  file_path: string;   // path inside the document library, e.g. "Kleza/Digital Marketing/Clients Info/Kleza/Daily Updates to Vik/Daliy Updates.xlsx"
}

async function getGraphToken(config: SharePointConfig): Promise<string> {
  const url = `https://login.microsoftonline.com/${config.tenant_id}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: config.client_id,
    client_secret: config.client_secret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const res = await fetch(url, { method: 'POST', body });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`SharePoint auth failed (${res.status}): ${text}`);
  }

  const json = await res.json();
  if (!json.access_token) throw new Error('No access_token in auth response');
  return json.access_token;
}

async function getSiteId(token: string, host: string, sitePath: string): Promise<string> {
  const url = `https://graph.microsoft.com/v1.0/sites/${host}:${sitePath}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `Failed to get SharePoint site (${res.status}). ` +
      `Make sure the Azure AD app has "Sites.Read.All" permission with admin consent granted.\n${text}`
    );
  }
  const json = await res.json();
  return json.id;
}

export async function downloadSharePointFile(config: SharePointConfig): Promise<Buffer> {
  const token = await getGraphToken(config);
  const siteId = await getSiteId(token, config.host, config.site_path);

  // Access file by its known path inside the document library
  const encodedPath = config.file_path.split('/').map(encodeURIComponent).join('/');
  const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`;

  const res = await fetch(contentUrl, {
    headers: { Authorization: `Bearer ${token}` },
    redirect: 'follow',
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to download file from SharePoint (${res.status}): ${text}`);
  }

  return Buffer.from(await res.arrayBuffer());
}
