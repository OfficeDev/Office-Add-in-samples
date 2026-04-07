/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/home/home.js?98ea":
/*!**************************!*\
  !*** ./src/home/home.js ***!
  \**************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "d0639b64073e42251e85.js";

/***/ }),

/***/ "./src/shared/shared.js?7c7c":
/*!******************************!*\
  !*** ./src/shared/shared.js ***!
  \******************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "a8288c95c024cb832f78.js";

/***/ }),

/***/ "./src/shared/visualization.js":
/*!*************************************!*\
  !*** ./src/shared/visualization.js ***!
  \*************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "93b7d16fc07c4cf4f8c6.js";

/***/ }),

/***/ "./src/home/home.css":
/*!***************************!*\
  !*** ./src/home/home.css ***!
  \***************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "f3465d850a41c7f35253.css";

/***/ }),

/***/ "./src/shared/shared.css":
/*!*******************************!*\
  !*** ./src/shared/shared.css ***!
  \*******************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "26bfcaa6a1bb507a6929.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		if (!(moduleId in __webpack_modules__)) {
/******/ 			delete __webpack_module_cache__[moduleId];
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		__webpack_require__.p = "/";
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"home": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!**************************!*\
  !*** ./src/home/home.js ***!
  \**************************/
(function () {
  "use strict";

  // The onReady function must be run each time a new page is loaded.
  Office.onReady(function (info) {
    shared.initialize();
    displayDataOrRedirect();
  });

  // Checks if a binding exists, and either displays the visualization
  //        or redirects to the data-binding page.
  function displayDataOrRedirect() {
    Office.context.document.bindings.getByIdAsync(shared.bindingID, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const binding = result.value;
        let handler = function () {
          displayDataForBinding(binding);
        };
        binding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, handler);
      } else {
        window.location.href = 'data-binding.html';
      }
    });
  }

  // Queries the binding for its data, then delegates to the visualization script.
  function displayDataForBinding(binding) {
    binding.getDataAsync({
      coercionType: Office.CoercionType.Table,
      valueFormat: Office.ValueFormat.Unformatted,
      filterType: Office.FilterType.OnlyVisible
    }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        visualization.display(document.getElementById('display-data'), result.value, showError);
      } else {
        showError('Could not read data.');
      }
    });
    function showError(message) {
      document.getElementById('display-data').innerHTML = '<div class="notice">' + '    <h3>Error</h3>' + $('<p/>', {
        text: message
      })[0].outerHTML + '    <a href="data-binding.html">' + '        <b>Bind to a different data range?</b>' + '    </a>' + '</div>';
    }
  }
})();
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!****************************!*\
  !*** ./src/home/home.html ***!
  \****************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/shared.js */ "./src/shared/shared.js?7c7c"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/visualization.js */ "./src/shared/visualization.js"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_2___ = new URL(/* asset import */ __webpack_require__(/*! ./home.js */ "./src/home/home.js?98ea"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_3___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/shared.css */ "./src/shared/shared.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_4___ = new URL(/* asset import */ __webpack_require__(/*! ./home.css */ "./src/home/home.css"), __webpack_require__.b);
// Module
var code = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <title></title>\r\n    <" + "script src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\" type=\"text/javascript\"><" + "/script>\r\n\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_0___ + "\" type=\"text/javascript\"><" + "/script>\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_1___ + "\" type=\"text/javascript\"><" + "/script>\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_2___ + "\" type=\"text/javascript\"><" + "/script>\r\n\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_3___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_4___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n<body dir=\"ltr\">\r\n    <div id=\"content-main\">\r\n        <div class=\"padding\">\r\n            <br />\r\n            <a id=\"back-button\" href=\"data-binding.html\">Back</a>\r\n            <br /><br />\r\n            <p><b>Add visualization content here.</b> For example:</p>\r\n            <div id=\"display-data\">\r\n                Loading...\r\n            </div>\r\n            <br /><br />\r\n            <a target=\"_blank\" href=\"https://go.microsoft.com/fwlink/?LinkId=276813\">Find more samples online...</a>\r\n        </div>\r\n    </div>\r\n    <img src=\"https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-content-add-in-data-visualization-run\" />\r\n</body>\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=home.js.map