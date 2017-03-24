if (typeof (pxbiD365) == "undefined") {
    var pxbiD365 = {};
}

(function() {
    var self = this;
	
    self.Picklist = {};
    (function() { 
        var self = this;
        var context = {};

        self.filterDependentPicklist = function(context, dependentPicklistName, jsonConfig) {
            var attribute = context.getEventSource();
            var attributeValue = attribute.getValue();
            var dependentPicklist = Xrm.Page.getControl(dependentPicklistName);

            
        }

        self.displaySubGrid = function(context, subGridName, jsonConfig) {
            var attribute = context.getEventSource();
            var attributeValue = attribute.getValue();
            var subGrid = Xrm.Page.getControl(subGridName);

            
        }
    }).apply(self.Picklist);
}).apply(pxbiD365);