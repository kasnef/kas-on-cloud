import axios from "axios";
import { helper } from "../utils/helper";
import type { MicrosoftConfig, MicrosoftUploadConfig } from "../types/config";
import { generateMicrosoftAccessToken } from "../utils/microsoft-connect";

const siteIdCache = new Map<string, string>();
const libraryIdCache = new Map<string, string>();

export async function showLog(show: boolean): Promise<boolean> {
  return show;
}

export async function getSiteId(
  tenentName: string,
  siteName: string,
  accessToken: string,
  isShowLog: boolean = false,
) {
  if (siteIdCache.has(`${tenentName}-${siteName}`)) {
    const cachedSiteId = siteIdCache.get(`${tenentName}-${siteName}`);
    if (isShowLog) {
      console.log(
        `[kas-on-cloud]: Using cached site ID for "${siteName}": ${cachedSiteId}`,
      );
    }
    return cachedSiteId;
  }

  if (!tenentName) {
    throw new Error("[kas-on-cloud]: Tenent name is required to get site ID");
  }

  if (!siteName) {
    throw new Error("[kas-on-cloud]: Site name is required to get site ID");
  }

  if (!accessToken) {
    throw new Error("[kas-on-cloud]: Access token is required to get site ID");
  }

  const url = `https://graph.microsoft.com/v1.0/sites/${tenentName}.sharepoint.com:/sites/${siteName}`;

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

  siteIdCache.set(`${tenentName}-${siteName}`, siteId);

  if (isShowLog) {
    console.log(`[kas-on-cloud]: Site id for "${siteName}": ${siteId}`);
  }

  return siteId;
}

export async function getDocumentLibraryId(
  tenentName: string, // for getSiteId
  siteName: string, // for getSiteId
  accessToken: string,
  isShowLog: boolean = false,
) {
  if (libraryIdCache.has(`${tenentName}-${siteName}`)) {
    const cachedLibraryId = libraryIdCache.get(`${tenentName}-${siteName}`);
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

  const siteId = await getSiteId(tenentName, siteName, accessToken, isShowLog);

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

  libraryIdCache.set(`${tenentName}-${siteName}`, libraryId);

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
  tokenDto: MicrosoftConfig,
  uploadDto: MicrosoftUploadConfig,
) {
  const {
    tenentId,
    clientId,
    clientSecret,
    scope = "https://graph.microsoft.com/.default",
  } = tokenDto;

  const {
    tenentName,
    siteName,
    fileName,
    fileContent,
    isShowLog = false,
    folderPath = "",
  } = uploadDto;

  const missingParams = Object.entries({
    tenentName,
    siteName,
    fileName,
    fileContent,
    isShowLog,
  })
    .filter(([_, v]) => !v)
    .map(([k]) => k);

  if (missingParams.length > 0) {
    throw new Error(
      `[kas-on-cloud]: Missing required Microsoft config params: ${missingParams.join(", ")}`,
    );
  }

  const accessToken = await generateMicrosoftAccessToken(
    {
      tenentId,
      clientId,
      clientSecret,
      scope,
    },
    isShowLog,
  );

  const librabyId = await getDocumentLibraryId(
    tenentName,
    siteName,
    accessToken.accessToken,
    isShowLog,
  );

  const normalizeFolderPath = helper.normailzePath(folderPath);

  const encodedPath = normalizeFolderPath?.trim()
    ? `${`root:/${normalizeFolderPath}`}`
    : `${"root:"}`;

  const url = `https://graph.microsoft.com/v1.0/drives/${librabyId}/${encodedPath}/${fileName}:/content`;

  const response = await axios.put(url, fileContent, {
    headers: {
      Authorization: `Bearer ${accessToken.accessToken}`,
      "Content-Type": "application/octet-stream",
    },
  });

  if (response.status !== 201) {
    throw new Error(
      `[kas-on-cloud]: Failed to upload file: ${response.statusText}`,
    );
  }

  if (isShowLog) {
    console.log(
      `[kas-on-cloud]: File "${fileName}" uploaded successfully to SharePoint`,
    );
  }

  return response.data;
}

// export async function multiUploadToSharePoint(
//   tenentName: string,
//   siteName: string,
//   files: { fileName: string; fileContent: Buffer }[],
//   accessToken: string,
//   isShowLog: boolean = false,
// ) {
//   const siteId = await getSiteId(tenentName, siteName, accessToken, isShowLog);
//   const libraryId = await getDocumentLibraryId(
//     tenentName,
//     siteName,
//     accessToken,
//     isShowLog,
//   );

//   const uploadPromises = files.map(({ fileName, fileContent }) =>
//     uploadToSharePoint(
//       tenentName,
//       siteName,
//       fileName,
//       fileContent,
//       accessToken,
//       isShowLog,
//     ),
//   );

//   return Promise.all(uploadPromises);
// }