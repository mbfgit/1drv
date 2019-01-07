function getInfo(filesinfo) {
    var msg = [];
    for(var i=0;i<filesinfo.value.length;i++) {
	msg.push({name: filesinfo.value[i].name, uri: filesinfo.value[i].webUrl,
	    fileId: filesinfo.value[i].id, driveId: filesinfo.value[i].parentReference.driveId,
	    driveType: filesinfo.value[i].parentReference.driveType,
	    mimeType: filesinfo.value[i].file.mimeType, size: filesinfo.value[i].size,
            listId: filesinfo.value[i].sharepointIds.listId,
            listItemId: filesinfo.value[i].sharepointIds.listItemId});
    }
    return msg;
}
var oneDriveFilePickerError = 
    function ( e ) {
	if (window.ecmError) {
	    window.ecmError(JSON.stringify(e, null, 4));
	}
    }
var oneDriveFilePickerSuccess = 
    function ( filesinfo ) {
	// OSX + iOS
	if(window.webkit && window.webkit.messageHandlers.ecmShare) {
	    var msg = { files : getInfo(filesinfo)};
	    window.webkit.messageHandlers.ecmShare.postMessage(msg);
	}
	// CEF
	if (window.ecmShare) {
	    var msg = { files : getInfo(filesinfo)};
	    window.ecmShare(JSON.stringify(msg,null,2));
	}
    }
var oneDriveFilePickerCancel =
    function ( ) {
	if (window.ecmCancel) {
	    window.ecmCancel();
	}
    }
function getOptions(clientId, loginHint, isConsumerAccount, action) {
    var odOptions = {
	clientId: clientId,
	action: action,
	multiSelect: true,
	openInNewWindow: true,
	success: function(files) {
	    oneDriveFilePickerSuccess(files);
	},
	cancel: function() {
	    oneDriveFilePickerCancel();
	},
	error: function(e) {
	    oneDriveFilePickerError(e);
	},
	redirectUri: window.location.href,
	advanced: {
		queryParameters: "select=id,name,size,file,folder,webUrl,parentReference,sharepointIds",
		navigation: {
			sourceTypes: ["OneDrive", "Sites"]
		}
	}
    };
    if(typeof loginHint !== "undefined" && typeof isConsumerAccount !== "undefined") {
        odOptions.advanced.loginHint = loginHint;
        odOptions.advanced.isConsumerAccount = isConsumerAccount;
    }
    return odOptions;
}
function launchOneDrivePicker(clientId, loginHint, isConsumerAccount, action) {
    action = action || "query";
    var odOptions = getOptions(clientId, loginHint, isConsumerAccount, action);
    OneDrive.open( odOptions );
}
