/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Global constant: binding ID
    app.bindingID = 'myBinding';

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        write("getting here 0");
    };

    return app;

})();