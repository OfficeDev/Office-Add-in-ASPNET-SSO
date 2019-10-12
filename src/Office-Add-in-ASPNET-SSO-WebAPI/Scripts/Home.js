// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/


Office.initialize = function (reason) {
	// Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
		// After the DOM is loaded, app-specific code can run.
		// Add any initialization logic to this function.

        $("#getGraphAccessTokenButton").click(function () {
			getOneDriveFiles();
		});
	});
}

var timesGetOneDriveFilesHasRun = 0;
var triedWithoutForceConsent = false;

function getOneDriveFiles() {
	timesGetOneDriveFilesHasRun++;
	triedWithoutForceConsent = true;
   // getDataWithToken({ asyncContext: 'foo', forceConsent: false, forceAddAccount: false, authChallenge: '' });

    // To test the non-SSO authorization path, comment out the preceding line and uncomment the following line.

    dialogFallback();

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
                showResult([error.toString()]);
            });
	})
	.fail(function (result) {
		handleServerSideErrors(result);
	});
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

		// If consent was not granted (or was revoked) for one or more permissions,
		// the add-in's web service relays the AADSTS65001 error. Try to get the token
		// again with the forceConsent option.
		if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
			showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);
			dialogFallback();
			return;
		}
		else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
			showResult(['The add-in is asking for a type of permission that is not recognized.']);
			return;
		}
		else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
			showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
			return;
		}
	}

	// If the token sent to MS Graph is expired or invalid, start the whole process over.
	if (result.code === 'InvalidAuthenticationToken') {
		timesGetOneDriveFilesHasRun = 0;
		triedWithoutForceConsent = false;
		getOneDriveFiles();
	}
	else {
		logError("Unspecified server side error " + result);
	}
}

function handleClientSideErrors(result) {

	// TODO: Handle forGraphAccess error once that's implemented and fallback to dialog.

	switch (result.error.code) {

		case 13001:
			// The user is not logged in, or the user cancelled without responding a
			// prompt to provide a 2nd authentication factor. (See comment about two-
			// factor authentication in the fail callback of the getData method.)
			// Either way start over and force a sign-in. 
			getDataWithToken({ forceAddAccount: true });
			break;
		case 13002:
			// The user's sign-in or consent was aborted. Ask the user to try again
			// but no more than once again.
			if (timesGetOneDriveFilesHasRun < 2) {
				showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
			} else {
				logError("13002 consent aborted. " + result);
			}
			break;
		case 13003:
			// The user is logged in with an account that is neither work or school, nor Microsoft Account.
			showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
			break;
		case 13005:
			// The Office host has not been authorized to the add-in's web service
			// or the user has not granted the service permission to their `profile`.
			getDataWithToken({ forceConsent: true });
			break;
		case 13006:
			// Unspecified error in the Office host.
			showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
			break;
		case 13007:
			// The Office host cannot get an access token to the add-ins web service/application.
            showResult(['That operation cannot be done at this time. Please try again later. (13007)']);
			break;
		case 13008:
			// The user triggered an operation that calls getAccessTokenAsync before a previous call of it completed.
			showResult(['Please try that operation again after the current operation has finished.']);
			break;
		case 13009:
			// The add-in does not support forcing consent. Try signing the user in without forcing consent, unless
			// that's already been tried.
			if (triedWithoutForceConsent) {
				showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
			} else {
				getDataWithToken({ forceConsent: false });
			}
			break;
		default:
			logError("Unspecified client side error " + result);
			dialogFallback();
			break;
	}
}

function getDataWithToken(options) {
	Office.context.auth.getAccessTokenAsync(options,
		function (result) {
			if (result.status === "succeeded") {
				var accessToken = result.value;
				getData("/api/values", accessToken);
			}
			else {
				handleClientSideErrors(result);
			}
		});
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
    let message = JSON.parse(arg.message);

    if (message.status === "success") {
        // We now have a valid access token.
        loginDialog.close();
        getData("/api/files", message.accessToken);
    } else {
        // Something went wrong with authentication or the authorization of the web application.
        loginDialog.close();
        showResult(["Unable to successfully authenticate user or authorize application. Error is: " + message.error]);
    }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
	var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

	// height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
	Office.context.ui.displayDialogAsync(fullUrl,
		{ height: 60, width: 30 }, function (result) {
			console.log("Dialog has initialized. Wiring up events");
			loginDialog = result.value;
			loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
		});
}