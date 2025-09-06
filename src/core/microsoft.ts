import axios from "axios";
import type {
  FileUploadItem,
  GetListFileFromSharepoint,
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