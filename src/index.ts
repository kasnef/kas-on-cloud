import {
  getSiteId,
  getDocumentLibraryId,
  getItemListFromSharepoint,
  uploadToSharePoint,
  multiUploadToSharepoint,
  clearCache,
} from "./core/microsoft";

import { generateMicrosoftAccessToken } from "./utils/microsoft-connect";