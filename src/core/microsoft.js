"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.showLog = showLog;
exports.getSiteId = getSiteId;
exports.getDocumentLibraryId = getDocumentLibraryId;
exports.clearCache = clearCache;
exports.uploadToSharePoint = uploadToSharePoint;
exports.multiUploadToSharepoint = multiUploadToSharepoint;
var axios_1 = require("axios");
var helper_1 = require("../utils/helper");
var siteIdCache = new Map();
var libraryIdCache = new Map();
function showLog(show) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, show];
        });
    });
}
function getSiteId(tenantName_1, siteName_1, accessToken_1) {
    return __awaiter(this, arguments, void 0, function (tenantName, siteName, accessToken, isShowLog) {
        var cachedSiteId, url, response, siteId;
        if (isShowLog === void 0) { isShowLog = false; }
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (siteIdCache.has("".concat(tenantName, "-").concat(siteName))) {
                        cachedSiteId = siteIdCache.get("".concat(tenantName, "-").concat(siteName));
                        if (isShowLog) {
                            console.log("[kas-on-cloud]: Using cached site ID for \"".concat(siteName, "\": ").concat(cachedSiteId));
                        }
                        return [2 /*return*/, cachedSiteId];
                    }
                    if (!tenantName) {
                        throw new Error("[kas-on-cloud]: Tenent name is required to get site ID");
                    }
                    if (!siteName) {
                        throw new Error("[kas-on-cloud]: Site name is required to get site ID");
                    }
                    if (!accessToken) {
                        throw new Error("[kas-on-cloud]: Access token is required to get site ID");
                    }
                    url = "https://graph.microsoft.com/v1.0/sites/".concat(tenantName, ".sharepoint.com:/sites/").concat(siteName);
                    return [4 /*yield*/, axios_1.default.get(url, {
                            headers: {
                                Authorization: "Bearer ".concat(accessToken),
                                "Content-Type": "application/json",
                            },
                        })];
                case 1:
                    response = _a.sent();
                    if (response.status !== 200) {
                        throw new Error("[kas-on-cloud]: Failed to get site ID: ".concat(response.statusText));
                    }
                    if (!response.data || !response.data.id) {
                        throw new Error("[kas-on-cloud]: Site ID not found in the response");
                    }
                    siteId = response.data.id.split(",")[1];
                    if (!siteId) {
                        throw new Error("[kas-on-cloud]: Site ID not found in the response");
                    }
                    siteIdCache.set("".concat(tenantName, "-").concat(siteName), siteId);
                    if (isShowLog) {
                        console.log("[kas-on-cloud]: Site id for \"".concat(siteName, "\": ").concat(siteId));
                    }
                    return [2 /*return*/, siteId];
            }
        });
    });
}
function getDocumentLibraryId(tenantName_1, siteName_1, accessToken_1) {
    return __awaiter(this, arguments, void 0, function (tenantName, // for getSiteId
    siteName, // for getSiteId
    accessToken, isShowLog) {
        var cachedLibraryId, siteId, url, response, libraries, libraryId;
        var _a;
        if (isShowLog === void 0) { isShowLog = false; }
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (libraryIdCache.has("".concat(tenantName, "-").concat(siteName))) {
                        cachedLibraryId = libraryIdCache.get("".concat(tenantName, "-").concat(siteName));
                        if (isShowLog) {
                            console.log("[kas-on-cloud]: Using cached document library ID for \"".concat(siteName, "\": ").concat(cachedLibraryId));
                        }
                        return [2 /*return*/, cachedLibraryId];
                    }
                    if (!siteName) {
                        throw new Error("[kas-on-cloud]: Site name is required to get document library ID");
                    }
                    if (!accessToken) {
                        throw new Error("[kas-on-cloud]: Access token is required to get document library ID");
                    }
                    return [4 /*yield*/, getSiteId(tenantName, siteName, accessToken, isShowLog)];
                case 1:
                    siteId = _b.sent();
                    url = "https://graph.microsoft.com/v1.0/sites/".concat(siteId, "/drives");
                    return [4 /*yield*/, axios_1.default.get(url, {
                            headers: {
                                Authorization: "Bearer ".concat(accessToken),
                                "Content-Type": "application/json",
                            },
                        })];
                case 2:
                    response = _b.sent();
                    if (response.status !== 200) {
                        throw new Error("[kas-on-cloud]: Failed to get document library ID: ".concat(response.statusText));
                    }
                    if (!response.data ||
                        !response.data.value ||
                        response.data.value.length === 0) {
                        throw new Error("[kas-on-cloud]: No document libraries found in the response");
                    }
                    libraries = response.data.value;
                    libraryId = (_a = libraries[0]) === null || _a === void 0 ? void 0 : _a.id;
                    if (!libraryId) {
                        throw new Error("[kas-on-cloud]: Document library \"".concat(libraryId, "\" not found"));
                    }
                    libraryIdCache.set("".concat(tenantName, "-").concat(siteName), libraryId);
                    if (isShowLog) {
                        console.log("[kas-on-cloud]: Document library ID: ".concat(libraryId));
                    }
                    return [2 /*return*/, libraryId];
            }
        });
    });
}
function clearCache() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            siteIdCache.clear();
            libraryIdCache.clear();
            console.log("[kas-on-cloud]: Microsoft caches cleared");
            return [2 /*return*/];
        });
    });
}
function uploadToSharePoint(accessToken_1, tenantName_1, siteName_1, fileName_1, fileContent_1) {
    return __awaiter(this, arguments, void 0, function (accessToken, tenantName, siteName, fileName, fileContent, isShowLog, folderPath) {
        var missingParams, librabyId, normalizeFolderPath, encodedPath, url, response;
        if (isShowLog === void 0) { isShowLog = false; }
        if (folderPath === void 0) { folderPath = ""; }
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    missingParams = Object.entries({
                        accessToken: accessToken,
                        tenantName: tenantName,
                        siteName: siteName,
                        fileName: fileName,
                        fileContent: fileContent,
                        isShowLog: isShowLog,
                    })
                        .filter(function (_a) {
                        var _ = _a[0], v = _a[1];
                        return !v;
                    })
                        .map(function (_a) {
                        var k = _a[0];
                        return k;
                    });
                    if (missingParams.length > 0) {
                        throw new Error("[kas-on-cloud]: Missing required Microsoft config params: ".concat(missingParams.join(", ")));
                    }
                    return [4 /*yield*/, getDocumentLibraryId(tenantName, siteName, accessToken, isShowLog)];
                case 1:
                    librabyId = _a.sent();
                    normalizeFolderPath = helper_1.helper.normailzePath(folderPath);
                    encodedPath = (normalizeFolderPath === null || normalizeFolderPath === void 0 ? void 0 : normalizeFolderPath.trim())
                        ? "".concat("root:/".concat(normalizeFolderPath))
                        : "".concat("root:");
                    url = "https://graph.microsoft.com/v1.0/drives/".concat(librabyId, "/").concat(encodedPath, "/").concat(fileName, ":/content");
                    return [4 /*yield*/, axios_1.default.put(url, fileContent, {
                            headers: {
                                Authorization: "Bearer ".concat(accessToken),
                                "Content-Type": "application/octet-stream",
                            },
                        })];
                case 2:
                    response = _a.sent();
                    if (response.status !== 201) {
                        throw new Error("[kas-on-cloud]: Failed to upload file: ".concat(response.statusText));
                    }
                    if (isShowLog) {
                        console.log("[kas-on-cloud]: File \"".concat(fileName, "\" uploaded successfully to SharePoint"));
                    }
                    return [2 /*return*/, response.data];
            }
        });
    });
}
function multiUploadToSharepoint(accessToken_1, tenantName_1, siteName_1, files_1) {
    return __awaiter(this, arguments, void 0, function (accessToken, tenantName, siteName, files, isShowLog, folderPath) {
        var missingParams, librabyId, normalizeFolderPath, encodedPath, result, _i, files_2, file, fileName, fileContent, url, response;
        if (isShowLog === void 0) { isShowLog = false; }
        if (folderPath === void 0) { folderPath = ""; }
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    missingParams = Object.entries({
                        accessToken: accessToken,
                        tenantName: tenantName,
                        siteName: siteName,
                        files: files,
                    })
                        .filter(function (_a) {
                        var _ = _a[0], v = _a[1];
                        return !v;
                    })
                        .map(function (_a) {
                        var k = _a[0];
                        return k;
                    });
                    if (missingParams.length > 0) {
                        throw new Error("[kas-on-cloud]: Missing required Microsoft config params: ".concat(missingParams.join(", ")));
                    }
                    if (!Array.isArray(files) || files.length === 0) {
                        throw new Error("[kas-on-cloud]: 'files' must be a non-empty array");
                    }
                    return [4 /*yield*/, getDocumentLibraryId(tenantName, siteName, accessToken, isShowLog)];
                case 1:
                    librabyId = _a.sent();
                    normalizeFolderPath = helper_1.helper.normailzePath(folderPath);
                    encodedPath = (normalizeFolderPath === null || normalizeFolderPath === void 0 ? void 0 : normalizeFolderPath.trim())
                        ? "".concat("root:/".concat(normalizeFolderPath))
                        : "".concat("root:");
                    result = [];
                    _i = 0, files_2 = files;
                    _a.label = 2;
                case 2:
                    if (!(_i < files_2.length)) return [3 /*break*/, 5];
                    file = files_2[_i];
                    fileName = file.fileName, fileContent = file.fileContent;
                    url = "https://graph.microsoft.com/v1.0/drives/".concat(librabyId, "/").concat(encodedPath, "/").concat(fileName, ":/content");
                    if (!fileName || !fileContent) {
                        throw new Error("[kas-on-cloud]: Each file must have 'fileName' and 'fileContent' properties");
                    }
                    return [4 /*yield*/, axios_1.default.put(url, fileContent, {
                            headers: {
                                Authorization: "Bearer ".concat(accessToken),
                                "Content-Type": "application/octet-stream",
                            },
                        })];
                case 3:
                    response = _a.sent();
                    if (response.status !== 201) {
                        throw new Error("[kas-on-cloud]: Failed to upload file \"".concat(fileName, "\": ").concat(response.statusText));
                    }
                    if (isShowLog) {
                        console.log("[kas-on-cloud]: File \"".concat(fileName, "\" uploaded successfully to SharePoint"));
                    }
                    result.push(response.data);
                    _a.label = 4;
                case 4:
                    _i++;
                    return [3 /*break*/, 2];
                case 5: return [2 /*return*/, result];
            }
        });
    });
}
