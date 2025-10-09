# 📝 Changelog

All notable changes to this project will be recorded in this file.

---

## 🚀 [1.0.10] - 2025-10-09 - Fix build & republish (09/10/2025)

### 🛠 Fixes & Maintenance
- Rebuilt package due to missing compiled files in version **1.0.9**.
- Re-published to ensure all new SharePoint handler functions are properly included in the distributed build.

### ⚙️ Technical Notes
- No functional code changes.
- No API modifications.
- Same features as version **1.0.9**:
  - `getDriveIdFromSite`
  - `getAllDrivesFromSite`
  - `getItemByIdFromSharepoint`
  - `getItemByPathFromSharepoint`
  - `getFolderChildrenFromSharepoint`
  - `searchItemInSharepoint`
  - `downloadItemFromSharepoint`
  - `renewMicrosoftToken`

> ✅ This release only fixes the missing build artifacts — functionally identical to `v1.0.9`.
---

## 🚀 [1.0.9] - 2025-10-09

### ✨ Added

* **getDriveIdFromSite()** → Fetch the default Drive ID for a SharePoint site using `siteId`.
* **getAllDrivesFromSite()** → Retrieve all document libraries (drives) in a SharePoint site.
* **getItemByIdFromSharepoint()** → Retrieve detailed information about a file or folder by its `itemId`.
* **getItemByPathFromSharepoint()** → Retrieve item metadata using a file or folder path.
* **getFolderChildrenFromSharepoint()** → List all children items within a given folder by ID or path.
* **searchItemInSharepoint()** → Search for files and folders in a SharePoint drive by keyword.
* **downloadItemFromSharepoint()** → Stream a file’s content for download via Microsoft Graph API.
* **renewMicrosoftToken()** → Automatically renew and cache Microsoft OAuth tokens for seamless reauthentication.

### 🧩 Improved

* Unified error messages with `[kas-on-cloud]` prefix for better debugging consistency.
* Enhanced log outputs with `isShowLog` flag across multiple SharePoint helper functions.

### 🧰 Technical

* Added internal caching layer for Microsoft tokens using an in-memory `tokenCache`.
* Improved response mapping for SharePoint API to include file/folder type and metadata fields.

---

## 🧭 [1.0.7] - 2025-09-28

### 💠 Updated

* Change the response of `uploadToSharePoint()` and `multiUploadToSharepoint()`

### 🧩 Improved

* Updated docs and typings for SharePoint helper utilities.

---

## 🧭 [1.0.7] - 2025-09-06

### 💠 Updated

- Updated documentation
- Updated the cache function
- Update response for `uploadToSharePoint` and `multiUploadToSharepoint`

### 🧩 Improved

* Updated docs and typings for SharePoint helper utilities.

### 🐞 Fixed bug
- Fixed an issue where the library was unusable (My bad for not testing thoroughly :<)
- Fixed typos in code parameters

---

## 🪶 [1.0.0] - 2025-09-10

### 🚀 Initial Release

- Initial release of kas-on-cloud – a unified cloud uploader library for Node.js, starting with full support for Microsoft SharePoint.
- Ideal for automation, server-side upload tasks, or building internal tools.

### 📦 Package Info
- `Node.js` 18+
- Bundled with tsup for both ESM & CommonJS support
- Ready for future OneDrive integration (folder structure prepared)

### 📂 File structure
```bash
dist/                      ← Built code
src/
  ├── core/                ← Main upload logic
  ├── utils/               ← Auth + helpers
  ├── types/               ← Shared interfaces
  ├── __test__/            ← Test files (excluded from build)
index.ts                   ← Export API
```
