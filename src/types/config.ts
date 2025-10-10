import { Readable } from "stream";
export interface MicrosoftConfig {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  scope?: string | "https://graph.microsoft.com/.default";
  grantType?: string | "client_credentials";
}

export type ShowLog = boolean;

export interface getDriveId {
  siteId: string;
  accessToken: string;
}

export interface MicrosoftUploadConfig {
  accessToken: string;
  tenantName: string;
  siteName: string;
  fileName: string;
  fileContent: Buffer;
  isShowLog?: ShowLog;
  folderPath?: string; // Optional folder path in SharePoint
}

export interface FileUploadItem {
  fileName: string;
  fileContent: Buffer;
}

export interface GetListFileFromSharepoint {
  siteId: string;
  accessToken: string;
  isShowLog?: ShowLog;
  driveId?: string;
  isShorten?: boolean;
}

export interface GetListFileFromSharepointResponse {
  siteId: string;
  folder: {
    id: string;
    name: string;
    createdBy: {
      user: {
        displayName: string;
        email: string;
        id: string;
      };
    };
    eTag: string;
    lastModifiedDateTime: string;
    lastModifiedBy: {
      user: {
        displayName: string;
        email: string;
        id: string;
      };
    };
  }[];
}

export interface GetItemByIdFromSharepoint {
  siteId: string;
  accessToken: string;
  itemId: string;
  driveId?: string;
  isShowLog?: boolean;
}

export interface GetItemByPathFromSharepoint {
  siteId: string;
  accessToken: string;
  path: string;
  driveId?: string;
  isShowLog?: boolean;
}

export interface GetFolderChildrenFromSharepoint {
  siteId: string;
  accessToken: string;
  folderItemId?: string;
  folderPath?: string;
  driveId?: string;
  isShowLog?: boolean;
}

export interface SearchItemInSharepoint {
  siteId: string;
  accessToken: string;
  keyword: string;
  driveId?: string;
  isShowLog?: boolean;
}

export interface DownloadItemFromSharepoint {
  siteId: string;
  accessToken: string;
  itemId: string;
  driveId?: string;
}

export interface SharepointDownloadResult {
  stream: Readable;
  fileName: string;
  contentType: string;
  contentDisposition: string;
}
