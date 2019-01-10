function getInfo(filesinfo) {
    var msg = [];
    if(filesinfo.value) {
        for(var i=0;i<filesinfo.value.length;i++) {
            var file = {
                name: filesinfo.value[i].name,
                uri: filesinfo.value[i].webUrl,
                fileId: filesinfo.value[i].id,
                size: filesinfo.value[i].size,
                driveId: "",
                driveType: "",
                mimeType: "",
                listId: "",
                listItemId: ""
            };
            if(filesinfo.value[i].sharepointIds) {
                file.listId = filesinfo.value[i].sharepointIds.listId || "";
                file.listItemId = filesinfo.value[i].sharepointIds.listItemId || "";
            }
            if(filesinfo.value[i].file) {
                file.mimeType = filesinfo.value[i].file.mimeType || "";
            }
            if(filesinfo.value[i].parentReference) {
                file.driveId = filesinfo.value[i].parentReference.driveId || "";
                file.driveType = filesinfo.value[i].parentReference.driveType || "";
            }
            msg.push(file);
        }
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
    try {
        var additionalOptions = JSON.parse(document.getElementById('additionalOptions').value || '{}');
        if(additionalOptions) {
            if(additionalOptions.advanced) {
                if(additionalOptions.advanced.loginHint)
                {
                    odOptions.advanced.loginHint = additionalOptions.advanced.loginHint;
                }
                if(typeof additionalOptions.advanced.isConsumerAccount !== "undefined")
                {
                    odOptions.advanced.isConsumerAccount = additionalOptions.advanced.isConsumerAccount;
                }
                if(additionalOptions.advanced.endpointHint)
                {
                    odOptions.advanced.endpointHint = additionalOptions.advanced.endpointHint;
                }
                if(additionalOptions.advanced.accessToken)
                {
                    odOptions.advanced.accessToken = additionalOptions.advanced.accessToken;
                }
            }
            if(additionalOptions.navigation) {
                if(additionalOptions.navigation.sourceTypes)
                {
                    if(odOptions.navigation) {
                        odOptions.navigation.sourceTypes = additionalOptions.navigation.sourceTypes;
                    } else {
                        odOptions.navigation = { sourceTypes : additionalOptions.navigation.sourceTypes };
                    }
                }
            }
        }
    }
    catch(e) {
        console.error(e);
    }
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
function launchOneDrivePickerInDiv(clientId, parentDiv, loginHint, isConsumerAccount, action) {
    action = action || "query";
    var odOptions = getOptions(clientId, loginHint, isConsumerAccount, action);
        odOptions.advanced.iframeParentDiv = true;
    OneDrive.open( odOptions );
}
