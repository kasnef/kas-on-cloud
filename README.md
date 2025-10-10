<!-- Title -->

# 📎 kas-on-cloud – Upload Files to SharePoint & Google Drive with Node.js

> If this was helpful, give it a star!⭐ Thanks so muchhh ❤️

<!-- Badges -->

![OneDrive](https://img.shields.io/badge/OneDrive-white?style=for-the-badge&logo=Microsoft%20OneDrive&logoColor=0078D4)
![Google Drive](https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=googledrive&logoColor=white)
![TypeScript](https://img.shields.io/badge/typescript-%23007ACC.svg?style=for-the-badge&logo=typescript&logoColor=white)
![npm](https://img.shields.io/badge/npm-CB3837?style=for-the-badge&logo=npm&logoColor=white)
![yarn](https://img.shields.io/badge/Yarn-2C8EBB?style=for-the-badge&logo=yarn&logoColor=white)

> ⚡ A lightweight and flexible **Node.js/TypeScript** library for uploading files to **Microsoft SharePoint Document Libraries** and **Google Drive**.
> 📤 Supports **single and multiple file uploads**, automatic caching, and easy integration with **Microsoft Graph API** and **Google Drive API**.
> 🚀 Ideal for developers building backend services, CLI tools, or automation systems that need reliable **cloud file storage upload** features.

- [Microsoft Graph API – Docs](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
- [Google Drive API – Docs](https://developers.google.com/drive)

---

<!-- Table of Contents -->

## 📚 Table of Contents

- [👨‍💻 Author](#-author)
- [🔧 Installation](#-installation)
- [☁️ Upload to SharePoint](#️-upload-to-sharepoint)
  - [🧾 Prerequisites](#-prerequisites)
  - [📄 Upload a Single File](#-upload-a-single-file)
  - [📁 Upload Multiple Files](#-upload-multiple-files)
  - [🧹 Clear Cache](#-clear-cache)
  - [ℹ️ Notes](#️-notes)
- [📂 Upload to Google Drive](#-upload-to-google-drive)
<!--
  - [🧾 Prerequisites](#-prerequisites-for-google-drive)
  - [📄 Upload Usage](#-google-drive-upload-usage) -->

- [📜 License](#-license)

---

### 👨‍💻 Author

**Kasnef** – Developer & Creator of [kas-on-cloud](https://www.npmjs.com/package/kas-on-cloud)

[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/kasnef)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com/in/kasnef)
[![Telegram](https://img.shields.io/badge/Telegram-2CA5E0?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/kasnef)

> Passionate about backend systems, cloud automation, and building developer tools ⚙️  
> _“Code smart, not hard.”_ 💡

---

## 🔧 Installation

```bash
npm install kas-on-cloud
# or
yarn add kas-on-cloud
```

---

## ⚙️ Setup Instructions for SharePoint & Google Drive

### 🟦 SharePoint Integration Guide

#### Coming soon!

---

## ☁️ Upload to SharePoint

### 🧾 Prerequisites

To use SharePoint upload, you need:

- A valid OAuth 2.0 Access Token
- Your tenant name (e.g., `mytenant`)
- Site name (e.g., `mySite`)
- Files to upload as `Buffer`

---

### ⚙️ generateMicrosoftAccessToken()

> Generates a Microsoft access token for authenticating with the Microsoft Graph API. This function also handles token caching and renewal for optimal performance.

```ts
generateMicrosoftAccessToken(
  config: MicrosoftConfig,
  isShowLog = false
): Promise<MicrosoftAccessTokenResponse>
```

#### Parameters

- **config (MicrosoftConfig)**: An object containing the necessary authentication credentials.
  - _tenantId (string)_: The ID of your Azure Active Directory tenant.
  - _clientId (string)_: The client ID of your registered application in Azure AD.
  - _clientSecret (string)_: The client secret for the application.
  - _scope (string, optional)_: The requested permission scope. Defaults to https://graph.microsoft.com/.default.
  - _grantType (string, optional)_: The grant type. Defaults to client_credentials.
- **isShowLog (boolean, optional)**: Set to true to display detailed logs during execution. Defaults to `false`.

#### ✅ Example

```ts
import { generateMicrosoftAccessToken } from "kas-on-cloud";

const config = {
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
  scope: "https://graph.microsoft.com/.default",
};

const tokenResponse = await generateMicrosoftAccessToken(config, true);
const accessToken = tokenResponse.accessToken;
```

#### ↪️ Response

```JSON
{
  "accessToken": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
  "expiresIn": 3599,
  "extExpiresIn": 3599
}
```

---

### ♻️ renewMicrosoftToken()

> Refreshes an expired Microsoft access token automatically using client credentials.

```ts
import { renewMicrosoftToken } from "kas-on-cloud";

const config = {
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
  scope: "https://graph.microsoft.com/.default",
};

renewMicrosoftToken(
  config: MicrosoftConfig,
  isShowLog = false
  ): Promise<MicrosoftAccessTokenResponse>
```

#### Parameters

- `config` (MicrosoftConfig): Contains Azure AD credentials.
- `tenantId` (string): Azure Active Directory tenant ID.
- `clientId` (string): Registered app’s client ID.
- `clientSecret` (string): App’s client secret.
- `isShowLog` (boolean, optional): Enable logs for debugging.

#### ✅ Example

```ts
import { renewMicrosoftToken } from "kas-on-cloud";

const newToken = await renewMicrosoftToken({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
});
console.log(newToken.accessToken);
```

#### ↪️ Response

```ts
{
  "accessToken": "eyJ0eXAiOiJKV1QiLCJhbGciOi...",
  "expiresIn": 3599,
  "extExpiresIn": 3599
}
```

---

### 🧭 getSiteId()

> Fetches the unique identifier for a SharePoint site based on the tenant and site names.

```ts
getSiteId(
  tenantName: string,
  siteName: string,
  accessToken: string,
  isShowLog = false
): Promise<string>
```

#### Parameters

- `tenantName` (string): Your SharePoint tenant name (e.g., `mytenant`).
- `siteName` (string): The name of the SharePoint site (e.g., `mySite`).
- `accessToken` (string): A valid OAuth 2.0 access token.
- `isShowLog` (boolean, optional): Set to `true` to enable logging. Defaults to `false`.

#### ✅ Example

```ts
import { getSiteId } from "kas-on-cloud";

const siteId = await getSiteId("mytenant", "mysite", accessToken, true);
console.log(`Site ID: ${siteId}`);
```

#### ↪️ Response

```JSON
"mytenant.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

---

### 🧭 getDriveIdFromSite()

> Fetches the default Drive ID (document library) for a given SharePoint site using its siteId.

```ts
getDriveIdFromSite(
  siteId: string,
  accessToken: string
): Promise<string>
```

#### Parameters

- `siteId` (string): The unique identifier of your SharePoint site. You can obtain it using getSiteId() beforehand.
- `accessToken` (string): A valid OAuth 2.0 access token with Microsoft Graph permissions (e.g., Files.Read, Sites.Read.All).

#### ✅ Example

```ts
import { getDriveIdFromSite } from "kas-on-cloud";

const driveId = await getDriveIdFromSite(siteId, accessToken);
console.log(`Drive ID: ${driveId}`);
```

#### ↪️ Response

```JSON
"b!4U8mx5a9hEya7v8g3s-BYJfC3s1MbRNNt5jZcVMBaE8RjhUe7X3ylRbXeA8S"
```

---

### 🧭 getAllDrivesFromSite()

> Retrieves a list of all document libraries (drives) associated with a SharePoint site.

```ts
getAllDrivesFromSite(
  siteId: string,
  accessToken: string
): Promise<
  {
    id: string;
    name: string;
    webUrl: string;
  }[]
>
```

#### Parameters

- `siteId` (string): The unique identifier of your SharePoint site. You can obtain it using getSiteId() beforehand.
- `accessToken` (string): A valid OAuth 2.0 access token with Microsoft Graph permissions (e.g., Files.Read, Sites.Read.All).

#### ✅ Example

```ts
import { getAllDrivesFromSite } from "kas-on-cloud";

const drives = await getAllDrivesFromSite(siteId, accessToken);
console.log("Available drives:", drives);
```

#### ↪️ Response

```JSON
[
  {
    "id": "b!9A7VJ9spkUeE1h6aaYZb9aRAxG01GzxJm_3uZsA3Sd3zD5yFvPp2PCbU4u7Hg",
    "name": "Documents",
    "webUrl": "https://mytenant.sharepoint.com/sites/mysite/Shared%20Documents"
  },
  {
    "id": "b!kU4tsP8ghEeE9u8Zb5X2WqUB2b4KxAqLp9uX2PzNqEbZyY7x2Y4rJzK5Zs4RA",
    "name": "Site Assets",
    "webUrl": "https://mytenant.sharepoint.com/sites/mysite/SiteAssets"
  }
]
```

---

### 🗂 getDocumentLibraryId()

> Fetches the ID of the default document library for a given SharePoint site.

```ts
getDocumentLibraryId(
  tenantName: string,
  siteName: string,
  accessToken: string,
  isShowLog = false
): Promise<string>
```

#### Parameters

- `tenantName (string)`: The SharePoint tenant name.
- `siteName (string)`: The name of the SharePoint site.
- `accessToken (string)`: A valid OAuth 2.0 access token.
- `isShowLog (boolean, optional)`: Set to true to enable logging. Defaults to `false`.

#### ✅ Example

```ts
import { getDocumentLibraryId } from "kas-on-cloud";

const libraryId = await getDocumentLibraryId(
  "mytenant",
  "mysite",
  accessToken,
  true,
);
console.log(`Document Library ID: ${libraryId}`);
```

#### ↪️ Response

> The function returns a string containing the document library ID.

```code
"b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

---

### 🔎 getItemListFromSharepoint()

> Lists all items from the SharePoint root or a specific document library.

```Ts
getItemListFromSharepoint({
  siteId: string,
  accessToken: string,
  isShowLog?: boolean,
  driveId?: string,
  isShorten?: boolean,
}): Promise<any[]>
```

#### Parameters

- `siteId (string):` The ID of the SharePoint site.
- `accessToken (string)`: A valid OAuth 2.0 access token.
- `isShowLog (boolean, optional)`: Set to true to enable logging. Defaults to `false`.
- `driveId (string, optional)`: The ID of a specific document library (drive). If not provided, it queries the default library.
- `isShorten (boolean, optional)`: Coming soon - intended to shorten the result list.

#### ✅ Example

```Ts
import { getItemListFromSharepoint } from "kas-on-cloud";

const items = await getItemListFromSharepoint({
  siteId: "mytenant.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  accessToken: accessToken,
  isShowLog: true,
});
```

#### ↪️ Response

> An array of objects representing the items in the library.

```JSON
[
  {
    "createdDateTime": "2025-01-01T12:00:00Z",
    "id": "01...",
    "lastModifiedDateTime": "2025-01-01T12:00:00Z",
    "name": "MyFile.txt",
    "webUrl": "https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/MyFile.txt",
    "size": 1024,
    "file": {
      "mimeType": "text/plain"
    }
  }
]
```

---

### 🧭 getItemByPathFromSharepoint()

> Retrieves a file or folder from SharePoint using its path instead of an ID.

```ts
import { getItemByPathFromSharepoint } from "kas-on-cloud";

getItemByPathFromSharepoint({
  siteId: string,
  path: string,
  accessToken: string,
  isShowLog?: boolean,
}): Promise<any>
```

#### Parameters

- `siteId` (string): The SharePoint site ID.
- `path` (string): The full item path (e.g., /Shared Documents/MyFolder/file.txt).
- `accessToken` (string): A valid OAuth 2.0 access token.
- `isShowLog` (boolean, optional): Enable logs for debugging.

#### ✅ Example

```ts
import { getItemByPathFromSharepoint } from "kas-on-cloud";

const item = await getItemByPathFromSharepoint({
  siteId: "your-site-id",
  path: "/Shared Documents/Reports/Report.xlsx",
  accessToken,
  isShowLog: true,
});
```

#### ↪️ Response

```ts
{
  "id": "01...",
  "name": "Report.xlsx",
  "size": 54231,
  "webUrl": "https://tenant.sharepoint.com/sites/mysite/Shared%20Documents/Reports/Report.xlsx"
}
```

---

### 📂 getFolderChildrenFromSharepoint()

> Fetches all files and subfolders inside a given SharePoint folder.

```ts
import { getFolderChildrenFromSharepoint } from "kas-on-cloud";

getFolderChildrenFromSharepoint({
  siteId: string,
  folderPath: string,
  accessToken: string,
  isShowLog?: boolean,
}): Promise<any[]>
```

#### Parameters

- `siteId` (string): The SharePoint site ID.
- `folderPath` (string): The full item path (e.g., /Shared Documents/MyFolder/file.txt).
- `accessToken` (string): A valid OAuth 2.0 access token.
- `isShowLog` (boolean, optional): Enable logs for debugging.

#### ✅ Example

```ts
import { getFolderChildrenFromSharepoint } from "kas-on-cloud";

const children = await getFolderChildrenFromSharepoint({
  siteId: "your-site-id",
  folderPath: "/Shared Documents/Reports",
  accessToken,
});
```

#### ↪️ Response

```ts
[
  {
    id: "01...",
    name: "Annual.pdf",
    size: 20014,
    webUrl:
      "https://tenant.sharepoint.com/sites/mysite/Shared%20Documents/Reports/Annual.pdf",
  },
  {
    id: "02...",
    name: "Q1",
    folder: {},
  },
];
```

---

### 🧾 getItemByIdFromSharepoint()

> Retrieves detailed information about a specific item (file or folder) in a SharePoint document library by its unique ID.

```ts
getItemByIdFromSharepoint({
  siteId: string,
  accessToken: string,
  itemId: string,
  driveId?: string,
  isShowLog?: boolean,
}): Promise<any>
```

#### Parameters

- **siteId (string)**: The unique identifier of the SharePoint site.
- **accessToken (string)**: A valid Microsoft Graph API OAuth 2.0 access token.
- **itemId (string)**: The unique ID of the item (file or folder) to retrieve.
- **driveId (string, optional)**: The document library (drive) ID containing the item. If not provided, defaults to the root drive.
- **isShowLog (boolean, optional)**: Set to `true` to log detailed information to the console. Defaults to `false`.

#### ✅ Example

```ts
import { getItemByIdFromSharepoint } from "kas-on-cloud";

const item = await getItemByIdFromSharepoint({
  siteId: "mytenant.sharepoint.com,xxxx-xxxx-xxxx,xxxx-xxxx-xxxx",
  accessToken: accessToken,
  itemId: "01ABCDEFG12345",
  isShowLog: true,
});

console.log("Item info:", item);
```

#### ↪️ Response

> Returns an object containing metadata and file/folder details.

```json
{
  "id": "01ABCDEFG12345",
  "name": "ProjectReport.pdf",
  "size": 1048576,
  "webUrl": "https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/ProjectReport.pdf",
  "downloadUrl": "https://public.sharepoint.com/...",
  "lastModifiedDateTime": "2025-10-01T12:00:00Z",
  "createdDateTime": "2025-09-29T15:30:00Z",
  "file": {
    "mimeType": "application/pdf"
  },
  "folder": null
}
```

---

### 🔍 searchItemInSharepoint()

> Searches for items in a SharePoint site based on name or keyword.

```ts
import { searchItemInSharepoint } from "kas-on-cloud";

searchItemInSharepoint({
  siteId: string,
  query: string,
  accessToken: string,
  isShowLog?: boolean,
}): Promise<any[]>
```

#### Parameters

- `siteId` (string): The SharePoint site ID.
- `query` (string): The search query or keyword.
- `accessToken` (string): A valid OAuth 2.0 access token.
- `isShowLog` (boolean, optional): Enable logs for debugging.

#### ✅ Example

```ts
import { searchItemInSharepoint } from "kas-on-cloud";

const results = await searchItemInSharepoint({
  siteId: "your-site-id",
  query: "invoice",
  accessToken,
  isShowLog: true,
});
```

#### ↪️ Response

```ts
[
  {
    id: "01...",
    name: "invoice_Jan.pdf",
    path: "/Shared Documents/Billing",
    webUrl:
      "https://tenant.sharepoint.com/sites/mysite/Shared%20Documents/Billing/invoice_Jan.pdf",
  },
];
```

---

### 📥 downloadItemFromSharepoint()

> Downloads a SharePoint file and returns it as a Buffer. This method supports backend proxying so that clients can download files without requiring direct SharePoint access.

```ts
import { downloadItemFromSharepoint } from "kas-on-cloud";

downloadItemFromSharepoint({
  siteId: string,
  itemId: string,
  accessToken: string,
  isShowLog?: boolean,
}): Promise<{ fileName: string; buffer: Buffer }>
```

#### Parameters

- `siteId` (string): The SharePoint site ID.
- `itemId` (string): The unique ID of the file to download.
- `accessToken` (string): A valid OAuth 2.0 access token.
- `isShowLog` (boolean, optional): Enable logs for debugging.

#### ✅ Example

```ts
import { downloadItemFromSharepoint } from "kas-on-cloud";

const file = await downloadItemFromSharepoint({
  siteId: "your-site-id",
  itemId: "01ABCDEFG",
  accessToken,
});

// Example BE proxy:
res.setHeader('Content-Type', file.contentType);
res.setHeader('Content-Disposition', file.contentDisposition);
file.stream.pipe(res);
```

#### ↪️ Response

```ts
{
  stream: Readable {
    _readableState: {
      objectMode: false,
      highWaterMark: 16384,
      buffer: BufferList { head: null, tail: null, length: 0 },
      length: 0,
      pipes: null,
      flowing: true,
      ended: false,
      endEmitted: false
    },
    readable: true,
    _events: [Object: null prototype] { end: [Function: onend] },
    _eventsCount: 1
  },

  fileName: "Report_Q3_2025.xlsx",
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  contentDisposition: 'attachment; filename="Report_Q3_2025.xlsx"',
}
```

---

### 📄 uploadToSharePoint()

> Uploads a single file to a SharePoint document library, with an option to specify a target folder.

```ts
uploadToSharePoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  fileName: string,
  fileContent: Buffer,
  isShowLog = false,
  folderPath = ""
): Promise<string>
```

#### Parameters

- `accessToken (string)`: A valid OAuth 2.0 access token.
- `tenantName (string)`: The SharePoint tenant name.
- `siteName (string)`: The SharePoint site name.
- `fileName (string)`: The name for the file on SharePoint.
- `fileContent (Buffer)`: The file content as a Buffer.
- `isShowLog (boolean, optional)`: Set to true to enable logging. Defaults to `false`.
- `folderPath (string, optional)`: The destination folder path on SharePoint (e.g., MyFolder/SubFolder)

#### ✅ Example

```ts
import { uploadToSharePoint } from "kas-on-cloud";

const sharepointUrl = await uploadToSharePoint(
  accessToken,
  "mytenant",
  "mySite",
  "myFile.txt",
  Buffer.from("Hello, SharePoint!"),
  true, // show log (default: false)
  "MyFolder", // folder path on sharepoint (optional)
);
console.log(`File uploaded to: ${sharepointUrl}`);
```

#### ↪️ Response

> The function returns a string containing the web URL of the uploaded file.

````STRING
"https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/MyFolder/myFile.txt"```
````

---

### 📤 multiUploadToSharepoint()

> Uploads multiple files to a SharePoint document library simultaneously.

```ts
multiUploadToSharepoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  files: FileUploadItem[],
  isShowLog = false,
  folderPath = ""
): Promise<any[]>
```

#### Parameters

- `accessToken (string)`: A valid OAuth 2.0 access token.
- `tenantName (string)`: The SharePoint tenant name.
- `siteName (string)`: The SharePoint site name.
- `files (FileUploadItem[])`: An array of file objects, where each object contains fileName and fileContent.
- `fileName (string)`: The name of the file.
- `fileContent (Buffer)`: The file content.
- `isShowLog (boolean, optional)`: Set to true to enable logging. Defaults to `false`.
- `folderPath (string, optional)`: The destination folder path for all files.

#### ✅ Example

```ts
import { multiUploadToSharepoint } from "kas-on-cloud";

const files = [
  { fileName: "file1.txt", fileContent: Buffer.from("File 1 content") },
  { fileName: "file2.txt", fileContent: Buffer.from("File 2 content") },
];

const results = await multiUploadToSharepoint(
  accessToken,
  "mytenant",
  "mySite",
  files,
  true, // show log (default: false)
  "MyFolder/SubFolder", // folder path on sharepoint (optional)
);
console.log("Upload results:", results);
```

#### ↪️ Response

> An array of objects containing the details of each uploaded file.

```JSON
[
  {
    "id": "01...",
    "name": "file1.txt",
    "webUrl": "https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/MyFolder/SubFolder/file1.txt",
    "size": 14
  },
  {
    "id": "02...",
    "name": "file2.txt",
    "webUrl": "https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/MyFolder/SubFolder/file2.txt",
    "size": 14
  }
]
```

---

### 🧹 Clear Cache

> Clears the cached site and document library IDs. Useful when you need to fetch fresh data.

```ts
clearCache(): void
```

#### ✅ Example

```Ts
import { clearCache } from "kas-on-cloud";

clearCache(); // Clears cached site and library IDs
```

---

### 📄 Notes

- Requires Microsoft Graph API permissions

- Handles caching automatically for performance

- Logs output with `[kas-on-cloud]` prefix for traceability

## 📂 Upload to Google Drive

> 🚧 _This feature is coming soon. We're actively working on Google Drive support!_

<!-- ---
### 🧾 Prerequisites for Google Drive
Planned requirements:

- OAuth 2.0 access token from Google
- Google Drive API enabled
- Proper scopes for file upload (`https://www.googleapis.com/auth/drive.file`)

---

### 📄 Google Drive Upload Usage

```ts
// Will be available in future release
import { uploadToGoogleDrive } from "kas-on-cloud";

await uploadToGoogleDrive({
  accessToken: "your-google-access-token",
  fileName: "hello.txt",
  fileContent: Buffer.from("Hello Google Drive!"),
  folderId: "optional-folder-id",
});
```

📢 _Stay tuned for upcoming updates and enhancements!_

--- -->

## 📜 License

![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

This project is licensed under the [MIT License](LICENSE).
