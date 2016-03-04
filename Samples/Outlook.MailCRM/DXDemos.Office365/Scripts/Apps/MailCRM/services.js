(function () {
    "use strict";

    angular.module("MailCRM.services", []).factory("stateSvc", ["$rootScope", "$http", function ($rootScope, $http) {
        var stateSvc = {};

        //private variables
        var hub, token = "";

        //public variables
        stateSvc.contact = {};
        stateSvc.clientId;
        stateSvc.idToken = null;
        stateSvc.activeNavIndex = 0;

        //initialize called when add-in loads to setup web sockets
        stateSvc.initialize = function () {
            //get a handle to the oAuthHub on the server
            hub = $.connection.oAuthHub;

            //create a function that the hub can call to broadcast oauth completion messages
            hub.client.oAuthComplete = function (user) {
                //the server just sent the add-in a token
                stateSvc.idToken.user = user;
                $rootScope.$broadcast("oAuthComplete", "/lookup");
            };

            //start listening on the hub for tokens
            $.connection.hub.start().done(function () {
                hub.server.initialize();

                //get the client identifier the popup will use to talk back
                stateSvc.clientId = $.connection.hub.id;
            });
        };

        //START: private functions
        var ensureTokenExists = function (callback) {
            if (token.length > 0)
                callback(token);
            else {
                //try to get a valid user token from exchange
                Office.context.mailbox.getUserIdentityTokenAsync(function (result) {
                    if (result.status == "succeeded") {
                        token = result.value;
                        callback(token)
                    }
                    else
                        callback(null);
                });
            }
        };
        //END: private functions

        //START: public functions
        stateSvc.validateUser = function (callback) {
            //try to get a valid user token from exchange
            Office.context.mailbox.getUserIdentityTokenAsync(function (asyncResult) {
                if (asyncResult.status == "succeeded") {
                    token = asyncResult.value;
                
                    //call the validate user API to ensure the user has a AAD token in cache
                    var data = { token: asyncResult.value };
                    $http.post("/api/User/Validate", data, { headers: { "Accept": "application/json; odata=verbose" } })
                        .success(function (result) {
                            stateSvc.idToken = result;
                            //check response
                            if (!result.validToken)
                                callback("/error"); //ERROR: the getUserIdentityTokenAsync token was invalid
                            else if (!result.validUser) {
                                callback(null); //user does not have a complete profile...take them to login
                            }
                            else {
                                //user exists and has valid token...goto detail
                                callback("/lookup");
                            }
                        })
                        .error(function (er) {
                            callback("/error"); //ERROR: error calling the validate web api
                        });
                }
                else
                    callback("/error"); //ERROR: error calling getUserIdentityTokenAsync 
            });
        };

        //lookup contact
        stateSvc.get = function (callback) {
            $http({
                method: "GET",
                url: "/api/Contact?id=" + encodeURIComponent(stateSvc.from.emailAddress.toLowerCase())
            }).success(function (data) {
                stateSvc.contact = data;
                $rootScope.$broadcast("contactUpdated");
                callback(data != null);
            }).error(function (er) {
                callback(false);
            });
        };

        //add a contact
        stateSvc.postContact = function (new_contact, callback) {
            $http.post("/api/Contact", new_contact, {
                headers: {
                    "callerName": Office.context.mailbox.userProfile.displayName,
                    "callerEmail": Office.context.mailbox.userProfile.emailAddress.toLowerCase()
                }
            }).success(function (data) {
                callback(true);
            }).error(function (er) {
                callback(false);
            });
        };

        //add a contnoteact
        stateSvc.postNote = function (note, callback) {
            $http.post("/api/Note", note, {
            }).success(function (data) {
                callback(true);
            }).error(function (er) {
                callback(false);
            });
        };

        //waiting indicator
        stateSvc.wait = function (val) {
            $rootScope.$broadcast("wait", val);
        };

        return stateSvc;
    }]);
})();