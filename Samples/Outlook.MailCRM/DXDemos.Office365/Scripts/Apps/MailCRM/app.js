(function () {
    "use strict";
    
    angular.module("mailCRMApp", ["ngRoute", "MailCRM.services", "MailCRM.controllers"])
        .config(function ($routeProvider, $locationProvider) {
            $routeProvider.when("/login", {
                controller: "loginCtrl",
                templateUrl: "../../templates/view-MailCRM-Login.html"
            })
            .when("/lookup", {
                controller: "lookupCtrl",
                templateUrl: "../../templates/view-MailCRM-Create.html"
            })
            .when("/detail", {
                controller: "crmCtrl",
                templateUrl: "../../templates/view-MailCRM-Detail.html"
            })
            .when("/contacts", {
                controller: "crmCtrl",
                templateUrl: "../../templates/view-MailCRM-Contacts.html"
            })
            .when("/invoices", {
                controller: "crmCtrl",
                templateUrl: "../../templates/view-MailCRM-Invoices.html"
            })
            .when("/location", {
                controller: "geoCtrl",
                templateUrl: "../../templates/view-MailCRM-Location.html"
            })
            .when("/notes", {
                controller: "crmCtrl",
                templateUrl: "../../templates/view-MailCRM-Notes.html"
            })
            .when("/attachments", {
                controller: "crmCtrl",
                templateUrl: "../../templates/view-MailCRM-Attachments.html"
            })
            .otherwise({ redirectTo: "/login" });
        })

        .directive("sideNav", ["$rootScope", "$location", "stateSvc", function ($rootScope, $location, stateSvc) {
            return {
                restrict: "E",
                templateUrl: "../../templates/directive-MailCRM-SideNav.html",
                scope: {
                    ngModel: "="
                },
                link: function (scope, element, attrs) {
                    scope.user = stateSvc.idToken.user;
                    scope.activeNavIndex = stateSvc.activeNavIndex;
                    scope.nav = function (path, index) {
                        stateSvc.activeNavIndex = index;
                        $location.path(path);
                    };
                }
            };
        }])

        .directive("spinner", function () {
            return {
                restrict: "E",
                templateUrl: "../../templates/directive-MailCRM-Spinner.html"
            };
        });
})();