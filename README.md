<!-- Title -->

# 📎 kas-on-cloud – Upload Files to SharePoint & Google Drive with Node.js

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

- [🔧 Installation](#-installation)
- [☁️ Upload to SharePoint](#️-upload-to-sharepoint)
  - [🧾 Prerequisites](#-prerequisites)
  - [📄 Upload a Single File](#-upload-a-single-file)
  - [📁 Upload Multiple Files](#-upload-multiple-files)
  - [🧹 Clear Cache](#-clear-cache)
  - [ℹ️ Notes](#️-notes)

- [📂 Upload to Google Drive](#-upload-to-google-drive)
  - [🧾 Prerequisites](#-prerequisites-for-google-drive)
  - [📄 Upload Usage](#-google-drive-upload-usage)

- [📜 License](#-license)

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

const libraryId = await getDocumentLibraryId("mytenant", "mysite", accessToken, true);
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
  "MyFolder" // folder path on sharepoint (optional)
);
console.log(`File uploaded to: ${sharepointUrl}`);
```

#### ↪️ Response

> The function returns a string containing the web URL of the uploaded file.

```STRING
"https://mytenant.sharepoint.com/sites/mySite/Shared%20Documents/MyFolder/myFile.txt"```
```

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
  "MyFolder/SubFolder" // folder path on sharepoint (optional)
);
console.log("Upload results:", results);
```

#### ↪️ Response

>An array of objects containing the details of each uploaded file.

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

- Logs output with ``[kas-on-cloud]`` prefix for traceability

---

## 📂 Upload to Google Drive

> 🚧 _This feature is coming soon. We're actively working on Google Drive support!_

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

---

## 📜 License

![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

This project is licensed under the [MIT License](LICENSE).
