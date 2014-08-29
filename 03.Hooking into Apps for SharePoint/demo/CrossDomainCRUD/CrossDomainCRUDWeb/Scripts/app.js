"use strict";

var Csom = window.Csom || {};

var getQueryStringParameter = function (p) {
    var params =
       document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == p)
            return decodeURIComponent(singleParam[1]);
    }
}

var appWebUrl = getQueryStringParameter("SPAppWebUrl");
var hostWebUrl = getQueryStringParameter("SPHostUrl");

(function () {
    "use strict";

    jQuery(function () {

        //Load required libraries
        jQuery.getScript(hostWebUrl + "/_layouts/15/MicrosoftAjax.js")
         .then(function () {
             return jQuery.getScript(hostWebUrl + "/_layouts/15/SP.Runtime.js");
         })
         .then(function () {
             return jQuery.getScript(hostWebUrl + "/_layouts/15/SP.js");
         })
         .then(function () {
             return jQuery.getScript(hostWebUrl + "/_layouts/15/SP.RequestExecutor.js");
         })
         .then(function () {
             return jQuery.getScript(hostWebUrl + "/_layouts/15/SP.UI.Controls.js");
         })
         .then(function () {

             //Display chrome control
             ChromeControl.render();

             //Create announcements list
             Csom.Lists.create(
                 appWebUrl,
                 hostWebUrl,
                 "Announcements",
                 "A team announcements list",
                 SP.ListTemplateType.announcements)
                 .always(function () {
                     //Load announcements list
                     return Csom.ViewModel.load(appWebUrl, hostWebUrl, "Announcements");
                 }
                 )
                 .fail(function (jqXHR, status) {
                     $("#loadingDiv").text(status.get_message());
                 })
                 .always(function () {
                     //bind to ui
                     ko.applyBindings(Csom.ViewModel, document.getElementById("announcementsTable"));
                     $("#mainDiv").show();
                     $("#loadingDiv").hide();
                 });

             //Add New Item
             $("#newItemButton").click(function () {
                 Csom.ListItems.create(appWebUrl, hostWebUrl, "Announcements", $("#inputText").val())
                 .always(function () {
                     return Csom.ViewModel.load(appWebUrl, hostWebUrl, "Announcements");
                 }
                 )
                 .fail(function (jqXHR, status) {
                     $("#message").text(status.get_message());
                 })
             });

         });

    });
}());


var ChromeControl = function () {

    var render = function () {
        var options = {
            "appIconUrl": "../Images/AppIcon.png",
            "appTitle": "Cross-Domain CRUD App",
            "settingsLinks": [
                {
                    "linkUrl": "../Pages/Default.aspx?" + document.URL.split("?")[1],
                    "displayName": "Home"
                }
            ]
        };

        var nav = new SP.UI.Controls.Navigation(
                                "chrome_ctrl_placeholder",
                                options
                          );
        nav.setVisible(true);

    };

    return {
        render: render
    };
}();

