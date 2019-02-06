// Xrm.WebApi Polyfill for D365 version below 9.x
// https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi
// Uses jQuery version 1.6 and above to provide promises support for old browsers using $.Deferred
// Author: phoenixinobi
if (typeof (Xrm) == "undefined") { var Xrm = {}; }// namespace Xrm
(function() {
    var self = this;
    self.WebApi = {}; //hard-wire temporarily for testing
    //if (typeof (self.WebApi) == "undefined") { self.WebApi = {}; }// class WebApi
    (function() {
        var self = this;
        //private
        self.private = {
            //properties
            apiEndpointUrl: null,
            initialize: function() {
                console.log("Initializing Xrm.WebApi polyfill...");
                var version = Xrm.Page.context.getVersion();
                var versionTokens = version.split(".");
                this.apiEndpointUrl = Xrm.Page.context.getClientUrl() + "/api/data/v" + versionTokens[0] + "." + versionTokens[1] + "/";
                console.log("Xrm.WebApi.private.apiEndpointUrl is set to " + this.apiEndpointUrl);
                console.log("Xrm.WebApi polyfill initialized!");
            },
            getEntitySetName: function(entityLogicalName) {
                var deferred = $.Deferred();
                var xhr = new XMLHttpRequest();
                xhr.open("GET", this.apiEndpointUrl + "EntityDefinitions(LogicalName='" + entityLogicalName + "')?$select=EntitySetName", true);
                xhr.setRequestHeader("OData-MaxVersion", "4.0");
                xhr.setRequestHeader("OData-Version", "4.0");
                xhr.setRequestHeader("Accept", "application/json");
                xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                xhr.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        xhr.onreadystatechange = null;
                        if (this.status === 200) {
                            deferred.resolve(JSON.parse(this.response));
                        }
                        else {
                            deferred.reject(JSON.parse(this.response).error);
                        }
                    }
                };
                xhr.send();
    
                return deferred.promise();
            }
        }
        self.createRecord = function(entityLogicalName, data) {
            var deferred = $.Deferred();
            Xrm.WebApi.private.getEntitySetName(entityLogicalName).then(
                function success(result) {
                    var xhr = new XMLHttpRequest();
                    var entitySetName = result.EntitySetName;
                    xhr.open("POST", Xrm.WebApi.private.apiEndpointUrl + entitySetName, true);
                    xhr.setRequestHeader("OData-MaxVersion", "4.0");
                    xhr.setRequestHeader("OData-Version", "4.0");
                    xhr.setRequestHeader("Accept", "application/json");
                    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                    xhr.onreadystatechange = function () {
                        if (this.readyState === 4) {
                            xhr.onreadystatechange = null;
                            if (this.status === 204) {
                                var id = this.getResponseHeader("OData-EntityId")
                                    .replace(Xrm.WebApi.private.apiEndpointUrl, "")
                                    .replace(entitySetName, "")
                                    .replace("(", "")
                                    .replace(")", "");
                                deferred.resolve({
                                    entityType: entityLogicalName,
                                    id: id
                                });
                            }
                            else {
                                deferred.reject(JSON.parse(this.response).error);
                            }
                        }
                    };
                    xhr.send(JSON.stringify(data));
                },
                function (error) {
                    deferred.reject(error);
                }
            );
            
            return deferred.promise();
        };

        self.retrieveRecord = function(entityLogicalName, id, options) {
            var deferred = $.Deferred();
            Xrm.WebApi.private.getEntitySetName(entityLogicalName).then(
                function success(result) {
                    var xhr = new XMLHttpRequest();
                    xhr.open("GET", Xrm.WebApi.private.apiEndpointUrl + result.EntitySetName + "(" + id + ")" + options, true);
                    xhr.setRequestHeader("OData-MaxVersion", "4.0");
                    xhr.setRequestHeader("OData-Version", "4.0");
                    xhr.setRequestHeader("Accept", "application/json");
                    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                    xhr.onreadystatechange = function () {
                        if (this.readyState === 4) {
                            xhr.onreadystatechange = null;
                            if (this.status === 200) {
                                deferred.resolve(JSON.parse(this.response));
                            }
                            else {
                                deferred.reject(JSON.parse(this.response).error);
                            }
                        }
                    };
                    xhr.send();
                },
                function (error) {
                    deferred.reject(error);
                }
            );
            
            return deferred.promise();
        };

        self.updateRecord = function(entityLogicalName, id, data) {
            var deferred = $.Deferred();
            Xrm.WebApi.private.getEntitySetName(entityLogicalName).then(
                function success(result) {
                    var xhr = new XMLHttpRequest();
                    var entitySetName = result.EntitySetName;
                    xhr.open("PATCH", Xrm.WebApi.private.apiEndpointUrl + entitySetName + "(" + id + ")", true);
                    xhr.setRequestHeader("OData-MaxVersion", "4.0");
                    xhr.setRequestHeader("OData-Version", "4.0");
                    xhr.setRequestHeader("Accept", "application/json");
                    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                    xhr.onreadystatechange = function () {
                        if (this.readyState === 4) {
                            xhr.onreadystatechange = null;
                            if (this.status === 204) {
                                var id = this.getResponseHeader("OData-EntityId")
                                    .replace(Xrm.WebApi.private.apiEndpointUrl, "")
                                    .replace(entitySetName, "")
                                    .replace("(", "")
                                    .replace(")", "");
                                    deferred.resolve({
                                        entityType: entityLogicalName,
                                        id: id
                                    });
                            }
                            else {
                                deferred.reject(JSON.parse(this.response).error);
                            }
                        }
                    };
                    xhr.send(JSON.stringify(data));
                },
                function (error) {
                    deferred.reject(error);
                }
            );
            
            return deferred.promise();
        };

        self.deleteRecord = function(entityLogicalName, id) {
            var deferred = $.Deferred();
            Xrm.WebApi.private.getEntitySetName(entityLogicalName).then(
                function success(result) {
                    var xhr = new XMLHttpRequest();
                    xhr.open("GET", Xrm.WebApi.private.apiEndpointUrl + result.EntitySetName + "(" + id + ")", true);
                    xhr.setRequestHeader("OData-MaxVersion", "4.0");
                    xhr.setRequestHeader("OData-Version", "4.0");
                    xhr.setRequestHeader("Accept", "application/json");
                    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                    xhr.onreadystatechange = function () {
                        if (this.readyState === 4) {
                            xhr.onreadystatechange = null;
                            if (this.status === 200) {
                                deferred.resolve({
                                    entityType: entityLogicalName,
                                    id: id,
                                    name: JSON.parse(this.response).name
                                });
                                deferred.resolve();
                            }
                            else {
                                deferred.reject(JSON.parse(this.response).error);
                            }
                        }
                    };
                    xhr.send();
                },
                function (error) {
                    deferred.reject(error);
                }
            );
            
            return deferred.promise();
        };
        //TODO: find time to test paging
        self.retrieveMultipleRecords = function(entityLogicalName, options, maxPageSize) {
            var deferred = $.Deferred();
            var xhr = new XMLHttpRequest();
            xhr.open("GET", Xrm.WebApi.private.apiEndpointUrl + entityLogicalName + options, true);
            xhr.setRequestHeader("OData-MaxVersion", "4.0");
            xhr.setRequestHeader("OData-Version", "4.0");
            xhr.setRequestHeader("Accept", "application/json");
            xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            xhr.setRequestHeader("Prefer", "odata.maxpagesize=" + maxPageSize);
            xhr.onreadystatechange = function () {
                if (this.readyState === 4) {
                    xhr.onreadystatechange = null;
                    if (this.status === 200) {
                        deferred.resolve(JSON.parse(this.response));
                    }
                    else {
                        deferred.reject(JSON.parse(this.response).error);
                    }
                }
            };
            xhr.send();

            return deferred.promise();
        };
        //initialize
        Xrm.WebApi.private.initialize();
    }).apply(self.WebApi);
}).apply(Xrm);