        function getInfo(filesinfo) {
            var msg = [];
            for(var i=0;i<filesinfo.value.length;i++) {
                msg.push({name: filesinfo.value[i].name, uri: filesinfo.value[i].webUrl,
                    fileId: filesinfo.value[i].id, driveId: filesinfo.value[i].parentReference.driveId,
                    driveType: filesinfo.value[i].parentReference.driveType,
                    mimeType: filesinfo.value[i].file.mimeType, size: filesinfo.value[i].size});
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
                showResults('OneDrive Picker Cancelled');
                if (window.ecmCancel) {
                    window.ecmCancel();
                }
            }
        function getOptions(action, clientId) {
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
                    queryParameters: "select=id,name,size,file,folder,webUrl,parentReference"
                }
            };
            return odOptions;
        }
        function launchOneDrivePicker(clientId,action) {
            action = action || "query";
            var odOptions = getOptions(action, clientId);
            OneDrive.open( odOptions );
        }
