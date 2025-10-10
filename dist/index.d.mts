import { Readable } from 'stream';

interface MicrosoftConfig {
    tenantId?: string;
    clientId?: string;
    clientSecret?: string;
    scope?: string | "https://graph.microsoft.com/.default";
    grantType?: string | "client_credentials";
}
type ShowLog = boolean;
interface getDriveId {
    siteId: string;
    accessToken: string;
}
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
interface GetItemByIdFromSharepoint {
    siteId: string;
    accessToken: string;
    itemId: string;
    driveId?: string;
    isShowLog?: boolean;
}
interface GetItemByPathFromSharepoint {
    siteId: string;
    accessToken: string;
    path: string;
    driveId?: string;
    isShowLog?: boolean;
}
interface GetFolderChildrenFromSharepoint {
    siteId: string;
    accessToken: string;
    folderItemId?: string;
    folderPath?: string;
    driveId?: string;
    isShowLog?: boolean;
}
interface SearchItemInSharepoint {
    siteId: string;
    accessToken: string;
    keyword: string;
    driveId?: string;
    isShowLog?: boolean;
}
interface DownloadItemFromSharepoint {
    siteId: string;
    accessToken: string;
    itemId: string;
    driveId?: string;
}
interface SharepointDownloadResult {
    stream: Readable;
    fileName: string;
    contentType: string;
    contentDisposition: string;
}

declare function getSiteId(tenantName: string, siteName: string, accessToken: string, isShowLog?: boolean): Promise<any>;
declare function getDriveIdFromSite({ siteId, accessToken, }: getDriveId): Promise<string>;
declare function getAllDrivesFromSite(siteId: string, accessToken: string): Promise<{
    id: string;
    name: string;
    webUrl: string;
}[]>;
declare function getDocumentLibraryId(tenantName: string, // for getSiteId
siteName: string, // for getSiteId
accessToken: string, isShowLog?: boolean): Promise<any>;
declare function clearCache(): Promise<void>;
declare function uploadToSharePoint(accessToken: string, tenantName: string, siteName: string, fileName: string, fileContent: Buffer, isShowLog?: boolean, folderPath?: string): Promise<{
    id: any;
    name: any;
    size: any;
    url: any;
    downloadUrl: any;
    createdDateTime: any;
    lastModifiedDateTime: any;
}>;
declare function multiUploadToSharepoint(accessToken: string, tenantName: string, siteName: string, files: FileUploadItem[], isShowLog?: boolean, folderPath?: string): Promise<{
    id: any;
    name: any;
    size: any;
    url: any;
    downloadUrl: any;
    createdDateTime: any;
    lastModifiedDateTime: any;
}[]>;
declare function getItemListFromSharepoint({ siteId, accessToken, isShowLog, driveId, isShorten, }: GetListFileFromSharepoint): Promise<any>;
declare function getItemByIdFromSharepoint({ siteId, accessToken, itemId, driveId, isShowLog, }: GetItemByIdFromSharepoint): Promise<{
    id: any;
    name: any;
    size: any;
    webUrl: any;
    downloadUrl: any;
    lastModifiedDateTime: any;
    createdDateTime: any;
    file: any;
    folder: any;
}>;
declare function getItemByPathFromSharepoint({ siteId, accessToken, path, driveId, isShowLog, }: GetItemByPathFromSharepoint): Promise<{
    id: any;
    name: any;
    size: any;
    webUrl: any;
    downloadUrl: any;
    lastModifiedDateTime: any;
    createdDateTime: any;
}>;
declare function getFolderChildrenFromSharepoint({ siteId, accessToken, folderItemId, folderPath, driveId, isShowLog, }: GetFolderChildrenFromSharepoint): Promise<any>;
declare function searchItemInSharepoint({ siteId, accessToken, keyword, driveId, isShowLog, }: SearchItemInSharepoint): Promise<any>;
declare function downloadItemFromSharepoint({ siteId, accessToken, itemId, driveId, }: DownloadItemFromSharepoint): Promise<SharepointDownloadResult>;

type MicrosoftAccessTokenResponse = {
    accessToken: string;
    expiresIn: number;
    extExpiresIn: number;
};
declare function generateMicrosoftAccessToken(config: MicrosoftConfig, isShowLog?: boolean): Promise<MicrosoftAccessTokenResponse>;
declare function renewMicrosoftToken(config: MicrosoftConfig, isShowLog?: boolean): Promise<string>;

export { clearCache, downloadItemFromSharepoint, generateMicrosoftAccessToken, getAllDrivesFromSite, getDocumentLibraryId, getDriveIdFromSite, getFolderChildrenFromSharepoint, getItemByIdFromSharepoint, getItemByPathFromSharepoint, getItemListFromSharepoint, getSiteId, multiUploadToSharepoint, renewMicrosoftToken, searchItemInSharepoint, uploadToSharePoint };
