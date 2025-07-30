# KAS ON CLOUD
![OneDrive](https://img.shields.io/badge/OneDrive-white?style=for-the-badge&logo=Microsoft%20OneDrive&logoColor=0078D4)
![Google Drive](https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=googledrive&logoColor=white)
![TypeScript](https://img.shields.io/badge/typescript-%23007ACC.svg?style=for-the-badge&logo=typescript&logoColor=white)
![npm](https://img.shields.io/badge/npm-CB3837?style=for-the-badge&logo=npm&logoColor=white)
![yarn](https://img.shields.io/badge/Yarn-2C8EBB?style=for-the-badge&logo=yarn&logoColor=white)
### A lightweight Node.js library for uploading files to **Microsoft SharePoint Document Library || Google drive**.
### Supports both single and multiple file uploads.

## 📜 Licenses
![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

## USAGE
### Upload file to Microsoft Sharepoint Document Library

#### Before usage
##### 1. App Registration in Azure


#### 📥 Installation

```bash
npm install kas-on-cloud
```

#### 🛠️ Usage
#### 🔐 Prerequisites

- A valid OAuth 2.0 Access Token
- Your tenant name (e.g., mytenant)
- The site name (e.g., mySite)
- Files to upload as Buffer

#### 📄 Upload a Single File
```
import { uploadToSharePoint } from "kas-on-cloud";

const sharepointUrl = await uploadToSharePoint(
  accessToken,
  "mytenant",
  "mySite",
  "myFile.txt",
  Buffer.from("Hello, SharePoint!"),
  true,
  "MyFolder"
);
```
#### 📁 Upload Multiple Files
```
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
  true,
  "MyFolder/SubFolder"
);
```
#### 🧹 Clear Cache
```
import { clearCache } from "kas-on-cloud";

clearCache(); // Clears cached site and library IDs
```
### 📌 Notes
