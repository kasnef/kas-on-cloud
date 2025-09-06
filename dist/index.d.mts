interface MicrosoftConfig {
    tenentId?: string;
    clientId?: string;
    clientSecret?: string;
    scope?: string | 'https://graph.microsoft.com/.default';
    grandType?: string | 'client_credentials';
}
type ShowLog = boolean;
interface FileUploadItem {
    fileName: string;
    fileContent: Buffer;
}
interface GetListFileFromSharepoint {
    siteId: string;
    accessToken: string;
    isShowLog?: ShowLog;
    driveId?: string;
    isShorten?: boolean;
}

declare function getSiteId(tenantName: string, siteName: string, accessToken: string, isShowLog?: boolean): Promise<any>;
declare function getDocumentLibraryId(tenantName: string, // for getSiteId
siteName: string, // for getSiteId
accessToken: string, isShowLog?: boolean): Promise<any>;
declare function clearCache(): Promise<void>;
declare function uploadToSharePoint(accessToken: string, tenantName: string, siteName: string, fileName: string, fileContent: Buffer, isShowLog?: boolean, folderPath?: string): Promise<any>;
declare function multiUploadToSharepoint(accessToken: string, tenantName: string, siteName: string, files: FileUploadItem[], isShowLog?: boolean, folderPath?: string): Promise<any[]>;
declare function getItemListFromSharepoint({ siteId, accessToken, isShowLog, driveId, isShorten, }: GetListFileFromSharepoint): Promise<any>;

type MicrosoftAccessTokenResponse = {
    accessToken: string;
    expiresIn: number;
    extExpiresIn: number;
};
declare function generateMicrosoftAccessToken(config: MicrosoftConfig, isShowLog?: boolean): Promise<MicrosoftAccessTokenResponse>;

export { clearCache, generateMicrosoftAccessToken, getDocumentLibraryId, getItemListFromSharepoint, getSiteId, multiUploadToSharepoint, uploadToSharePoint };
