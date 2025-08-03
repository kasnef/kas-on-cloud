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

### ⚙️ Get access token

```ts
import { generateMicrosoftAccessToken } from "kas-on-cloud";
const accessToken = await generateMicrosoftAccessToken({
    "myTenantId",
    "clientId",
    "clientSecret",
    "scope", // default: 'https://graph.microsoft.com/.default'
  },
  true // show log (default: false)
)
```

### 🧭 getSiteId()

```ts
getSiteId(
  tenantName: string,
  siteName: string,
  accessToken: string,
  isShowLog = false
)
```
#### ✅ Example
```ts
const siteId = await getSiteId("mytenant", "mysite", token);
```

---

### 🗂 getDocumentLibraryId()
Fetches the document library ID for a given SharePoint site.
```ts
getDocumentLibraryId(
  tenantName: string,
  siteName: string,
  accessToken: string,
  isShowLog = false
)
```

#### ✅ Example
```ts
const libraryId = await getDocumentLibraryId("mytenant", "mysite", token);
```

---

### 📄 uploadToSharePoint()

```ts
uploadToSharePoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  fileName: string,
  fileContent: Buffer,
  isShowLog = false,
  folderPath = ""
)
```

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
```

---

### 📤 multiUploadToSharepoint()

```ts
multiUploadToSharepoint(
  accessToken: string,
  tenantName: string,
  siteName: string,
  files: FileUploadItem[],
  isShowLog = false,
  folderPath = ""
)
```

#### ✅ Example

```ts
import { multiUploadToSharepoint } from "kas-on-cloud";

const files = [
  { fileName: "file1.txt", fileContent: Buffer.from("File 1") },
  { fileName: "file2.txt", fileContent: Buffer.from("File 2") },
];

const result = await multiUploadToSharepoint(
  accessToken,
  "mytenant",
  "mySite",
  files,
  true, // show log (default: false)
  "MyFolder/SubFolder", // folder path on sharepoint (optional)
);
```

---

### 🔎 getItemListFromSharepoint()

List all files from SharePoint root or a specific document library.

```ts
getItemListFromSharepoint({
  siteId: string,
  accessToken: string,
  isShowLog?: boolean,
  driveId?: string,
  isShorten?: boolean,
})
```

#### ✅ Example

```ts
await getItemListFromSharepoint({
  siteId: "abc123xyz",
  accessToken: token,
  isShowLog: true,
});
```

---

### 🧹 Clear Cache

```ts
import { clearCache } from "kas-on-cloud";

clearCache(); // Clears cached site and library IDs
```

---

### 📄 Notes

- Requires Microsoft Graph API permissions

- Handles caching automatically for performance

- Logs output with [kas-on-cloud] prefix for traceability

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
