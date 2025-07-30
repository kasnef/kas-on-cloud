"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.helper = void 0;
var helper = /** @class */ (function () {
    function helper() {
    }
    helper.normailzePath = function (path) {
        return (path === null || path === void 0 ? void 0 : path.replace(/^\/+|\/+$/g, '')) || '';
    };
    return helper;
}());
exports.helper = helper;
