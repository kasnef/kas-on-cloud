import axios from "axios";
import type {
  FileUploadItem,
  GetFolderChildrenFromSharepoint,
  GetItemByIdFromSharepoint,
  GetItemByPathFromSharepoint,
  GetListFileFromSharepoint,
  SearchItemInSharepoint,
  SharepointDownloadResult,
  DownloadItemFromSharepoint,
  getDriveId,
} from "../types/config";
import { helper } from "../utils/helper";

const siteIdCache = new Map<string, string>();
const libraryIdCache = new Map<string, string>();

export async function showLog(show: boolean): Promise<boolean> {
  return show;
}

export async function getSiteId(
  tenantName: string,
  siteName: string,
  accessToken: string,
  isShowLog: boolean = false,
) {
  if (siteIdCache.has(`${tenantName}-${siteName}`)) {
    const cachedSiteId = siteIdCache.get(`${tenantName}-${siteName}`);
    if (isShowLog) {
      console.log(
        `[kas-on-cloud]: Using cached site ID for "${siteName}": ${cachedSiteId}`,
      );
    }
    return cachedSiteId;
  }

  if (!tenantName) {
    throw new Error("[kas-on-cloud]: Tenant name is required to get site ID");
  }

  if (!siteName) {
    throw new Error("[kas-on-cloud]: Site name is required to get site ID");
  }

  if (!accessToken) {
    throw new Error("[kas-on-cloud]: Access token is required to get site ID");
  }

  const url = `https://graph.microsoft.com/v1.0/sites/${tenantName}.sharepoint.com:/sites/${siteName}`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to get site ID: ${response.statusText}`,
    );
  }

  if (!response.data || !response.data.id) {
    throw new Error("[kas-on-cloud]: Site ID not found in the response");
  }

  const siteId = response.data.id.split(",")[1];

  if (!siteId) {
    throw new Error("[kas-on-cloud]: Site ID not found in the response");
  }

  siteIdCache.set(`${tenantName}-${siteName}`, siteId);

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Site id for "${siteName}": ${siteId}`);
  }

  return siteId;
}

export async function getDriveIdFromSite({
  siteId,
  accessToken,
}: getDriveId): Promise<string> {
  if (!siteId)
    throw new Error("[kas-on-cloud]: Site ID is required to get drive ID");
  if (!accessToken)
    throw new Error("[kas-on-cloud]: Access token is required to get drive ID");

  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      },
    );

    const driveId = response.data?.id;

    if (!driveId)
      throw new Error("[kas-on-cloud]: Could not find drive ID for this site");

    return driveId;
  } catch (error: any) {
    const msg =
      error.response?.data?.error?.message || error.message || "Unknown error";
    throw new Error(`[kas-on-cloud]: Failed to get drive ID - ${msg}`);
  }
}

export async function getAllDrivesFromSite(
  siteId: string,
  accessToken: string,
): Promise<
  {
    id: string;
    name: string;
    webUrl: string;
  }[]
> {
  if (!siteId)
    throw new Error("[kas-on-cloud]: Site ID is required to get drives");
  if (!accessToken)
    throw new Error("[kas-on-cloud]: Access token is required to get drives");

  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      },
    );

    const drives = response.data?.value || [];

    return drives.map((d: any) => ({
      id: d.id,
      name: d.name,
      webUrl: d.webUrl,
    }));
  } catch (error: any) {
    const msg =
      error.response?.data?.error?.message || error.message || "Unknown error";
    throw new Error(`[kas-on-cloud]: Failed to get all drives - ${msg}`);
  }
}

export async function getDocumentLibraryId(
  tenantName: string, // for getSiteId
  siteName: string, // for getSiteId
  accessToken: string,
  isShowLog: boolean = false,
) {
  if (libraryIdCache.has(`${tenantName}-${siteName}`)) {
    const cachedLibraryId = libraryIdCache.get(`${tenantName}-${siteName}`);
    if (isShowLog) {
      console.log(
        `[kas-on-cloud]: Using cached document library ID for "${siteName}": ${cachedLibraryId}`,
      );
    }
    return cachedLibraryId;
  }

  if (!siteName) {
    throw new Error(
      "[kas-on-cloud]: Site name is required to get document library ID",
    );
  }

  if (!accessToken) {
    throw new Error(
      "[kas-on-cloud]: Access token is required to get document library ID",
    );
  }

  const siteId = await getSiteId(tenantName, siteName, accessToken, isShowLog);

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to get document library ID: ${response.statusText}`,
    );
  }

  if (
    !response.data ||
    !response.data.value ||
    response.data.value.length === 0
  ) {
    throw new Error(
      "[kas-on-cloud]: No document libraries found in the response",
    );
  }

  const libraries = response.data.value;

  const libraryId = libraries[0]?.id;

  if (!libraryId) {
    throw new Error(
      `[kas-on-cloud]: Document library "${libraryId}" not found`,
    );
  }

  libraryIdCache.set(`${tenantName}-${siteName}`, libraryId);

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Document library ID: ${libraryId}`);
  }

  return libraryId;
}

export async function clearCache() {
  siteIdCache.clear();
  libraryIdCache.clear();
  console.log("[kas-on-cloud]: Microsoft caches cleared");
}

// #region upload

export async function uploadToSharePoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  fileName: string,
  fileContent: Buffer,
  isShowLog = false,
  folderPath = "",
) {
  const missingParams = Object.entries({
    accessToken,
    tenantName,
    siteName,
    fileName,
    fileContent,
  })
    .filter(([_, v]) => v === undefined || v === null || v === "")
    .map(([k]) => k);

  if (missingParams.length > 0) {
    throw new Error(
      `[kas-on-cloud]: Missing required Microsoft config params: ${missingParams.join(", ")}`,
    );
  }

  const librabyId = await getDocumentLibraryId(
    tenantName,
    siteName,
    accessToken,
    isShowLog,
  );

  const normalizeFolderPath = helper.normailzePath(folderPath);

  const encodedPath = normalizeFolderPath?.trim()
    ? `${`root:/${normalizeFolderPath}`}`
    : `${"root:"}`;

  const url = `https://graph.microsoft.com/v1.0/drives/${librabyId}/${encodedPath}/${fileName}:/content`;

  const response = await axios.put(url, fileContent, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/octet-stream",
    },
  });

  if (isShowLog) {
    console.log(
      `[kas-on-cloud]: File "${fileName}" uploaded successfully to SharePoint`,
    );
  }

  if (!response.data) {
    throw new Error("[kas-on-cloud]: No data returned from upload response");
  }

  const res = {
    id: response.data.id,
    name: response.data.name,
    size: response.data.size,
    url: response.data.webUrl,
    downloadUrl: response.data["@microsoft.graph.downloadUrl"],
    createdDateTime: response.data.createdDateTime,
    lastModifiedDateTime: response.data.lastModifiedDateTime,
  };

  console.log(`[kas-on-cloud]: Upload result: ${JSON.stringify(res, null, 2)}`);

  return res;
}

export async function multiUploadToSharepoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  files: FileUploadItem[],
  isShowLog = false,
  folderPath = "",
) {
  const missingParams = Object.entries({
    accessToken,
    tenantName,
    siteName,
    files,
  })
    .filter(([_, v]) => !v)
    .map(([k]) => k);

  if (missingParams.length > 0) {
    throw new Error(
      `[kas-on-cloud]: Missing required Microsoft config params: ${missingParams.join(", ")}`,
    );
  }

  if (!Array.isArray(files) || files.length === 0) {
    throw new Error(`[kas-on-cloud]: 'files' must be a non-empty array`);
  }

  const librabyId = await getDocumentLibraryId(
    tenantName,
    siteName,
    accessToken,
    isShowLog,
  );

  const normalizeFolderPath = helper.normailzePath(folderPath);

  const encodedPath = normalizeFolderPath?.trim()
    ? `${`root:/${normalizeFolderPath}`}`
    : `${"root:"}`;

  const result = [];

  for (const file of files) {
    const { fileName, fileContent } = file;

    const url = `https://graph.microsoft.com/v1.0/drives/${librabyId}/${encodedPath}/${fileName}:/content`;

    if (!fileName || !fileContent) {
      throw new Error(
        `[kas-on-cloud]: Each file must have 'fileName' and 'fileContent' properties`,
      );
    }

    const response = await axios.put(url, fileContent, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/octet-stream",
      },
    });

    if (response.status !== 201) {
      throw new Error(
        `[kas-on-cloud]: Failed to upload file "${fileName}": ${response.statusText}`,
      );
    }

    if (isShowLog) {
      console.log(
        `[kas-on-cloud]: File "${fileName}" uploaded successfully to SharePoint`,
      );
    }

    result.push(response.data);
  }

  const res = result.map((item) => ({
    id: item.data.id,
    name: item.data.name,
    size: item.data.size,
    url: item.data.webUrl,
    downloadUrl: item.data["@microsoft.graph.downloadUrl"],
    createdDateTime: item.data.createdDateTime,
    lastModifiedDateTime: item.data.lastModifiedDateTime,
  }));

  return res;
}

// #endregion

// #region Read / Download

export async function getItemListFromSharepoint({
  siteId,
  accessToken,
  isShowLog = false,
  driveId = "",
  isShorten = false,
}: GetListFileFromSharepoint) {
  if (!siteId) {
    throw new Error("[kas-on-cloud]: Site ID is required to get item list");
  }

  if (!accessToken) {
    throw new Error(
      "[kas-on-cloud]: Access token is required to get item list",
    );
  }

  const url = driveId
    ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`
    : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to get item list: ${response.statusText}`,
    );
  }

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Item list retrieved successfully`);
    console.log(`[kas-on-cloud]: ${response.data.value.length} items found`);
    if (isShorten) {
      console.log(`[kas-on-cloud]: Shortened item list: comming soon`);
    } else {
      console.log(
        `[kas-on-cloud]: ${JSON.stringify(response.data.value, null, 2)}`,
      );
    }
  }

  return response.data.value;
}

export async function getItemByIdFromSharepoint({
  siteId,
  accessToken,
  itemId,
  driveId = "",
  isShowLog = false,
}: GetItemByIdFromSharepoint) {
  if (!siteId) {
    throw new Error("[kas-on-cloud]: Site ID is required to get item");
  }

  if (!accessToken) {
    throw new Error("[kas-on-cloud]: Access token is required to get item");
  }

  if (!itemId) {
    throw new Error("[kas-on-cloud]: Item ID is required to get item");
  }

  const url = driveId
    ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`
    : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (response.status !== 200) {
      throw new Error(
        `[kas-on-cloud]: Failed to get item: ${response.statusText}`,
      );
    }

    const item = response.data;

    if (isShowLog) {
      console.log(`[kas-on-cloud]: Item retrieved successfully`);
      console.log(`[kas-on-cloud]: Item name: ${item.name}`);
      console.log(`[kas-on-cloud]: Item ID: ${item.id}`);
      if (item["@microsoft.graph.downloadUrl"]) {
        console.log(`[kas-on-cloud]: Download URL found`);
      }
    }

    return {
      id: item.id,
      name: item.name,
      size: item.size,
      webUrl: item.webUrl,
      downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
      lastModifiedDateTime: item.lastModifiedDateTime,
      createdDateTime: item.createdDateTime,
      file: item.file || null,
      folder: item.folder || null,
    };
  } catch (error: any) {
    const message =
      error.response?.data?.error?.message || error.message || "Unknown error";
    throw new Error(`[kas-on-cloud]: Failed to get item: ${message}`);
  }
}

export async function getItemByPathFromSharepoint({
  siteId,
  accessToken,
  path,
  driveId = "",
  isShowLog = false,
}: GetItemByPathFromSharepoint) {
  if (!siteId) {
    throw new Error("[kas-on-cloud]: Site ID is required to get item by path");
  }
  if (!accessToken) {
    throw new Error(
      "[kas-on-cloud]: Access token is required to get item by path",
    );
  }
  if (!path) {
    throw new Error("[kas-on-cloud]: Path is required to get item by path");
  }

  const url = driveId
    ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}${path}`
    : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive${path}`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to get item by path: ${response.statusText}`,
    );
  }

  const item = response.data;

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Item retrieved by path successfully`);
    console.log(`[kas-on-cloud]: Item name: ${item.name}`);
    if (item["@microsoft.graph.downloadUrl"]) {
      console.log(`[kas-on-cloud]: Download URL available`);
    }
  }

  return {
    id: item.id,
    name: item.name,
    size: item.size,
    webUrl: item.webUrl,
    downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    createdDateTime: item.createdDateTime,
  };
}

export async function getFolderChildrenFromSharepoint({
  siteId,
  accessToken,
  folderItemId = "",
  folderPath = "",
  driveId = "",
  isShowLog = false,
}: GetFolderChildrenFromSharepoint) {
  if (!siteId) {
    throw new Error(
      "[kas-on-cloud]: Site ID is required to get folder children",
    );
  }
  if (!accessToken) {
    throw new Error(
      "[kas-on-cloud]: Access token is required to get folder children",
    );
  }
  if (!folderItemId && !folderPath) {
    throw new Error(
      "[kas-on-cloud]: Either folderItemId or folderPath is required",
    );
  }

  const url = driveId
    ? folderItemId
      ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${folderItemId}/children`
      : `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}${folderPath}/children`
    : folderItemId
      ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${folderItemId}/children`
      : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive${folderPath}/children`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to get folder children: ${response.statusText}`,
    );
  }

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Folder children retrieved successfully`);
    console.log(`[kas-on-cloud]: ${response.data.value.length} items found`);
  }

  return response.data.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    size: item.size,
    isFolder: !!item.folder,
    webUrl: item.webUrl,
    downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
  }));
}

export async function searchItemInSharepoint({
  siteId,
  accessToken,
  keyword,
  driveId = "",
  isShowLog = false,
}: SearchItemInSharepoint) {
  if (!siteId) {
    throw new Error("[kas-on-cloud]: Site ID is required to search item");
  }
  if (!accessToken) {
    throw new Error("[kas-on-cloud]: Access token is required to search item");
  }
  if (!keyword) {
    throw new Error("[kas-on-cloud]: Keyword is required to search item");
  }

  const url = driveId
    ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/search(q='${encodeURIComponent(keyword)}')`
    : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${encodeURIComponent(keyword)}')`;

  const response = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(
      `[kas-on-cloud]: Failed to search item: ${response.statusText}`,
    );
  }

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Search completed`);
    console.log(`[kas-on-cloud]: Found ${response.data.value.length} results`);
  }

  return response.data.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    size: item.size,
    webUrl: item.webUrl,
    downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
  }));
}

export async function downloadItemFromSharepoint({
  siteId,
  accessToken,
  itemId,
  driveId = "",
}: DownloadItemFromSharepoint): Promise<SharepointDownloadResult> {
  if (!siteId)
    throw new Error("[kas-on-cloud]: Site ID is required to download item");
  if (!accessToken)
    throw new Error(
      "[kas-on-cloud]: Access token is required to download item",
    );
  if (!itemId)
    throw new Error("[kas-on-cloud]: Item ID is required to download item");

  const url = driveId
    ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/content`
    : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/content`;

  const response = await axios.get(url, {
    responseType: "stream",
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const fileName =
    response.headers["content-disposition"]
      ?.split("filename=")[1]
      ?.replace(/"/g, "") || "download.bin";

  const contentType =
    response.headers["content-type"] || "application/octet-stream";

  return {
    stream: response.data,
    fileName,
    contentType,
  };
}

// #endregion
