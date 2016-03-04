(function () {
    "use strict";

    angular.module("MailCRM.controllers", [])
        //waitingCtrl for spinner
        .controller("waitingCtrl", ["$scope", "$location", "stateSvc", function ($scope, $location, stateSvc) {
            $scope.wait = true;
            $scope.$on("wait", function (evt, val) {
                $scope.wait = val;
            });
        }])

        .controller("loginCtrl", ["$scope", "$location", "$window", "stateSvc", function ($scope, $location, $window, stateSvc) {
            //initialize the service
            stateSvc.initialize();
            var o365Redirect = "";

            // The initialize function must be run each time a new page is loaded
            Office.initialize = function (reason) {
                //get details of the item based on the item type
                var item = Office.context.mailbox.item;
                if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                    stateSvc.from = item.from;
                } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
                    stateSvc.from = item.organizer;
                }

                //validate the user and then navigate to the correct view
                stateSvc.validateUser(function (path) {
                    if (path != null)
                        $location.path(path);
                    else {
                        o365Redirect = stateSvc.idToken.loginUrl + encodeURIComponent(stateSvc.idToken.redirectUrl + stateSvc.idToken.user.id + "/" + stateSvc.clientId);
                        stateSvc.wait(false);
                    }
                });
            };

            //log into Office 365...uses popup
            $scope.loginO365 = function () {
                stateSvc.wait(true);
                $window.open(o365Redirect, "_blank", "width=720, height=300, scrollbars=0, toolbar=0, menubar=0, resizable=0, status=0, titlebar=0");
            };

            //wait for oauth response
            $scope.$on("oAuthComplete", function (evt, path) {
                $location.path(path);
                $scope.$apply();
            });
        }])

        //crmCtrl for managing the main logic
        .controller("crmCtrl", ["$scope", "stateSvc", function ($scope, stateSvc) {
            stateSvc.wait(false);
            $scope.contact = stateSvc.contact;
            $scope.newNote = {};

            var getDate = function () {
                var d = new Date();
                var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
                var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
                return days[d.getDay()] + ", " + months[d.getMonth()] + " " + d.getDate() + ", " + d.getFullYear();
            };

            $scope.postNote = function () {
                var d = new Date();
                $scope.newNote.email = stateSvc.from.emailAddress.toLowerCase();
                $scope.newNote.author_name = Office.context.mailbox.userProfile.displayName;
                $scope.newNote.author_email = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
                $scope.newNote.post_date = getDate();
                stateSvc.postNote($scope.newNote, function (success) {
                    if (success) {
                        stateSvc.contact.notes.splice(0, 0, $scope.newNote);
                        $scope.newNote = {};
                    }
                    else {
                        //TODO: notify of error
                    }
                })
            };

            $scope.$on("contactUpdated", function (event) {
                $scope.contact = stateSvc.contact;
            });
        }])

        //geoCtrl for Bing Maps integration
        .controller("geoCtrl", ["$scope", "stateSvc", function ($scope, stateSvc) {
            map = new Microsoft.Maps.Map(document.getElementById("divmap"), { credentials: "Amyr179PM1V3p0HVqiDFiDQERbex13DaD3U4sBQIcdkaKAv0b3xfUStygzs1pjl6", mapTypeId: Microsoft.Maps.MapTypeId.road });
            map.getCredentials(function (cred) {
                var req = 'https://dev.virtualearth.net/REST/v1/Locations?q=' + encodeURIComponent(stateSvc.contact.address) + '&output=json&jsonp=mapCallback&key=Amyr179PM1V3p0HVqiDFiDQERbex13DaD3U4sBQIcdkaKAv0b3xfUStygzs1pjl6';
                var script = document.createElement("script");
                script.setAttribute("type", "text/javascript");
                script.setAttribute("src", req);
                document.body.appendChild(script);
            });
        }])

        .controller("initializeCtrl", ["$scope", "$location", "stateSvc", function ($scope, $location, stateSvc) {
            //TODO
        }])

        //lookupCtrl for managing contact lookups and creates
        .controller("lookupCtrl", ["$scope", "$location", "$timeout", "stateSvc", function ($scope, $location, $timeout, stateSvc) {
            stateSvc.get(function (result) {
                if (result) {
                    //redirect to the detail view
                    $location.path("/detail");
                }
                else {
                    //show the create form
                    stateSvc.wait(false);
                }
            });

            $scope.contact = { name: stateSvc.from.displayName };
            $scope.postContact = function () {
                stateSvc.wait(true);
                $scope.contact.id = stateSvc.from.emailAddress.toLowerCase();
                stateSvc.postContact($scope.contact, function (postSuccess) {
                    if (postSuccess) {
                        //the add worked...now we need to go get the contact details
                        stateSvc.get(function (result) {
                            if (result) {
                                //redirect to the detail view
                                $location.path("/detail");
                            }
                            else {
                                //TODO: show error...should have been created
                                stateSvc.wait(false);
                            }
                        });
                    }
                    else {
                        //TODO: show error
                    }
                });
            };
        }]);
})();


//scripts for map callbacks
var map;
function mapCallback(result) {
    if (result.resourceSets[0].resources.length > 0) {
        var p = result.resourceSets[0].resources[0].point;
        var loc = new Microsoft.Maps.Location(p.coordinates[0], p.coordinates[1]);
        var pin = new Microsoft.Maps.Pushpin(loc);
        map.entities.push(pin);
        map.setView({ center: loc, zoom: 15 });
    }
}