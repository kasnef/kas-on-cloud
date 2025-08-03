export interface MicrosoftConfig {
  tenentId?: string;
  clientId?: string;
  clientSecret?: string;
  scope?: string | 'https://graph.microsoft.com/.default';
  grandType?: string | 'client_credentials';
}

export type ShowLog = boolean;

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
    }
  }[];
}
