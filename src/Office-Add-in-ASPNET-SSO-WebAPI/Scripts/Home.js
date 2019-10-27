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
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
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
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides functions to get ask the Office host to get an access token to the add-in
    and to pass that token to the server to get Microsoft Graph data.
*/
Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#getGraphDataButton').click(getGraphData);
    });
};
var retryGetAccessToken = 0;
function getGraphData() {
    return __awaiter(this, void 0, void 0, function () {
        var bootstrapToken, exception_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true })];
                case 1:
                    bootstrapToken = _a.sent();
                    // The /api/values controller will make the token exchange and use the 
                    // access token it gets back to make the call to MS Graph.
                    // Server-side errors are caught in the .fail block of getData.
                    getData("/api/values", bootstrapToken);
                    return [3 /*break*/, 3];
                case 2:
                    exception_1 = _a.sent();
                    // The only exceptions caught here are exceptions in your code in the try block
                    // and errors returned from the call of `getAccessToken` above.
                    if (exception_1.code) {
                        handleClientSideErrors(exception_1);
                    }
                    else {
                        showResult(["EXCEPTION: " + JSON.stringify(exception_1)]);
                    }
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
function getData(relativeUrl, accessToken) {
    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
        writeFileNamesToOfficeDocument(result)
            .then(function () {
            showResult(["Your data has been added to the document."]);
        })
            .catch(function (error) {
            // The error from writeFileNamesToOfficeDocument will begin 
            // "Unable to add filenames to document."
            showResult([JSON.stringify(error)]);
        });
    })
        .fail(function (result) {
        handleServerSideErrors(result);
    });
}
function handleClientSideErrors(error) {
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option.
            showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
            break;
        case 13006:
            // Only seen in Office on the Web.
            showResult(["Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
            break;
        case 13008:
            // Only seen in Office on the Web.
            showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
            break;
        case 13010:
            // Only seen in Office on the Web.
            showResult(["Follow the instructions to change your browser's zone configuration."]);
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            dialogFallback();
            break;
    }
}
function handleServerSideErrors(result) {
    // Our special handling on the server will cause the result that is returned
    // from a AADSTS50076 (a 2FA challenge) to have a Message property but no ExceptionMessage.
    var message = JSON.parse(result.responseText).Message;
    // Results from other errors (other than AADSTS50076) will have an ExceptionMessage property.
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    // Microsoft Graph requires an additional form of authentication. Have the Office host 
    // get a new token using the Claims string, which tells AAD to prompt the user for all 
    // required forms of authentication.
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    if (exceptionMessage) {
        // On rare occasions the bootstrap token is unexpired when Office validates it,
        // but expires by the time it is sent to AAD for exchange. AAD will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Retry the call of getAccessToken (no more than once). This time Office will return a 
        // new unexpired bootstrap token.
        if ((exceptionMessage.indexOf("The provided value for the 'assertion' is not valid. The assertion has expired.") !== -1)
            && (retryGetAccessToken <= 0)) {
            retryGetAccessToken++;
            getGraphData();
        }
        else {
            // For debugging: 
            // showResult(["AAD ERROR: " + JSON.stringify(exchangeResponse)]);  
            // For all other AAD errors, fallback to non-SSO sign-in.                            
            dialogFallback();
        }
    }
}
// Displays the data, assumed to be an array.
function showResult(data) {
    // Note that in this sample, the data parameter is an array of OneDrive file/folder
    // names. Encoding/sanitizing to protect against Cross-site scripting (XSS) attacks
    // is not needed because there are restrictions on what characters can be used in 
    // OneDrive file and folder names. These restrictions do not necessarily apply 
    // to other kinds of data including other kinds of Microsoft Graph data. So, to 
    // make this method safely reusable in other contexts, it uses the jQuery text() 
    // method which automatically encodes values that are passed to it.
    $.each(data, function (i) {
        var li = $('<li/>').addClass('ms-ListItem').appendTo($('#file-list'));
        var outerSpan = $('<span/>').addClass('ms-ListItem-secondaryText').appendTo(li);
        $('<span/>').addClass('ms-fontColor-themePrimary').appendTo(outerSpan).text(data[i]);
    });
}
function logError(result) {
    console.log("Status: " + result.status);
    console.log("Code: " + result.error.code);
    console.log("Name: " + result.error.name);
    console.log("Message: " + result.error.message);
}
// Dialog API
var loginDialog;
var redirectTo = "/files/index";
function dialogFallback() {
    var url = "/azureadauth/login";
    showLoginPopup(url);
}
// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
function processMessage(arg) {
    console.log("Message received in processMessage: " + JSON.stringify(arg));
    var message = JSON.parse(arg.message);
    if (message.status === "success") {
        // We now have a valid access token.
        loginDialog.close();
        getData("/api/files", message.accessToken);
    }
    else {
        // Something went wrong with authentication or the authorization of the web application.
        loginDialog.close();
        showResult(["Unable to successfully authenticate user or authorize application. Error is: " + message.error]);
    }
}
// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
    var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;
    // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
    Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
        console.log("Dialog has initialized. Wiring up events");
        loginDialog = result.value;
        loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    });
}
//# sourceMappingURL=Home.js.map