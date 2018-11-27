        function redeemURL(url, timeout) {
            timeout = timeout || 2000;
            fetch(url,{"credentials":"include","headers":{},"referrerPolicy":"no-referrer-when-downgrade","body":null,"method":"GET","mode":"no-cors"})
                .then(function(reponse) {
                setTimeout(function() {
                    // OSX + iOS
                    var msg = { url : url};
                    if(window.webkit && window.webkit.messageHandlers.URLRedeemed) {
                        window.webkit.messageHandlers.URLRedeemed.postMessage(msg);
                    }
                    // CEF
                    else if (window.URLRedeemed) {
                        window.URLRedeemed(JSON.stringify(msg,null,2));
                    }
                },timeout);
            });
        }
        function launchGraphAPILogin(clientId, guid) {
            scopes = "offline_access+openid+profile+User.Read+Files.ReadWrite.All";
            redirectUrl = window.location.href;
            // v2.0 authorize URL
            authorizeUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
            // Permission scopes
            requestUrl = authorizeUrl + "?scope="+scopes;
            // Code grant, will receive a code that can be redeemed for a token
            requestUrl += "&response_type=code";
            // client ID
            requestUrl += "&client_id="+clientId;
            // Add your app's redirect URL
            requestUrl += "&redirect_uri="+redirectUrl;
            requestUrl += "&response_mode=query"
            requestUrl += "&state="+guid;
            var opened = window.open(requestUrl);
        }
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
                    queryParameters: "select=id,name,size,file,folder,image,webUrl,thumbnails,activities,remoteItem,root,photo,searchResult,shared,createdBy,createdDateTime,lastModifiediBy,lastModifiedDateTime,cTag,eTag,permissions,parentReference,@microsoft.graph.sourceUrl"
                }
            };
            return odOptions;
        }
        function launchOneDrivePicker(clientId,action) {
            action = action || "query";
            var odOptions = getOptions(action, clientId);
            OneDrive.open( odOptions );
        }
