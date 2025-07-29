export interface MicrosoftConfig {
  tenentId?: string;
  clientId?: string;
  clientSecret?: string;
  scope?: string | 'https://graph.microsoft.com/.default';
  grandType?: string | 'client_credentials';
}

export type ShowLog = boolean;

export interface MicrosoftUploadConfig {
  tenentName: string;
  siteName: string;
  fileName: string;
  fileContent: Buffer;
  isShowLog?: ShowLog;
  folderPath?: string; // Optional folder path in SharePoint
}