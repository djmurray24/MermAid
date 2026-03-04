import { loginRequest, sharePointConfig } from "./authConfig";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Get an access token silently (or prompt login if needed)
async function getToken(msalInstance) {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error("No logged in user");

  try {
    const result = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
    return result.accessToken;
  } catch {
    // Silent failed, fall back to popup
    const result = await msalInstance.acquireTokenPopup(loginRequest);
    return result.accessToken;
  }
}

// Build the base Graph URL for our SharePoint folder
function folderBaseUrl() {
  const { hostname, siteName, folderPath } = sharePointConfig;
  const encodedPath = folderPath.split("/").map(encodeURIComponent).join("/");
  return `${GRAPH_BASE}/sites/${hostname}:/sites/${siteName}:/drive/root:${encodedPath}`;
}

// List all .mmd files in the folder
export async function listDiagrams(msalInstance) {
  const token = await getToken(msalInstance);
  const url = `${folderBaseUrl()}:/children?$filter=endswith(name,'.mmd')&$select=id,name,lastModifiedDateTime,size`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) throw new Error(`Failed to list diagrams: ${res.statusText}`);
  const data = await res.json();

  return (data.value || []).map(f => ({
    id: f.id,
    name: f.name.replace(/\.mmd$/, ""),
    lastModified: new Date(f.lastModifiedDateTime).toLocaleString(),
    size: f.size,
  }));
}

// Load a diagram's content by file ID
export async function loadDiagram(msalInstance, fileId) {
  const token = await getToken(msalInstance);
  const url = `${GRAPH_BASE}/sites/${sharePointConfig.hostname}:/sites/${sharePointConfig.siteName}:/drive/items/${fileId}/content`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) throw new Error(`Failed to load diagram: ${res.statusText}`);
  return await res.text();
}

// Save (create or update) a diagram as a .mmd file
export async function saveDiagram(msalInstance, title, content) {
  const token = await getToken(msalInstance);
  const fileName = `${title.replace(/[^a-z0-9\-_\s]/gi, "")}.mmd`;
  const url = `${folderBaseUrl()}/${encodeURIComponent(fileName)}:/content`;

  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "text/plain",
    },
    body: content,
  });

  if (!res.ok) throw new Error(`Failed to save diagram: ${res.statusText}`);
  const data = await res.json();

  return {
    id: data.id,
    name: title,
    lastModified: new Date(data.lastModifiedDateTime).toLocaleString(),
  };
}

// Delete a diagram by file ID
export async function deleteDiagram(msalInstance, fileId) {
  const token = await getToken(msalInstance);
  const url = `${GRAPH_BASE}/sites/${sharePointConfig.hostname}:/sites/${sharePointConfig.siteName}:/drive/items/${fileId}`;

  const res = await fetch(url, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) throw new Error(`Failed to delete diagram: ${res.statusText}`);
}