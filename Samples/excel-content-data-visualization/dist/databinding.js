/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/data-binding/data-binding.js?04f7":
/*!******************************************!*\
  !*** ./src/data-binding/data-binding.js ***!
  \******************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "bc7542269d29ea00cd3b.js";

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

/***/ "./src/data-binding/data-binding.css":
/*!*******************************************!*\
  !*** ./src/data-binding/data-binding.css ***!
  \*******************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "5bb60f667ae1ca89ead4.css";

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
/******/ 			"databinding": 0
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
/*!******************************************!*\
  !*** ./src/data-binding/data-binding.js ***!
  \******************************************/
(function () {
  'use strict';

  // The onReady function must be defined for each page in your add-in.
  Office.onReady(function (info) {
    shared.initialize();
    document.getElementById("bind-to-existing-data").addEventListener("click", () => bindToExistingData());
    if (dataInsertionSupported()) {
      document.getElementById("insert-sample-data").addEventListener("click", () => insertSampleData());
      document.getElementById("insert-data-available").style.display = "block";
      document.getElementById("insert-data-unavailable").style.display = "none";
    } else {
      document.getElementById("insert-data-available").style.display = "none";
      document.getElementById("insert-data-unavailable").style.display = "block";
    }
  });

  // Binds the visualization to existing data.
  function bindToExistingData() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Table, {
      id: shared.bindingID,
      sampleData: visualization.generateSampleData()
    }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        window.location.href = 'home.html';
      } else {
        shared.showNotification(result.error.name, result.error.message);
      }
    });
  }

  // Checks whether the current application supports setting selected data.
  function dataInsertionSupported() {
    return Office.context.document.setSelectedDataAsync && Office.context.document.bindings && Office.context.document.bindings.addFromSelectionAsync;
  }

  // Inserts sample data into the current selection (if supported).
  function insertSampleData() {
    Office.context.document.setSelectedDataAsync(visualization.generateSampleData(), function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, {
          id: shared.bindingID
        }, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            window.location.href = 'home.html';
          } else {
            shared.showNotification(result.error.name, result.error.message);
          }
        });
      } else {
        shared.showNotification(result.error.name, result.error.message);
      }
    });
  }
})();
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!********************************************!*\
  !*** ./src/data-binding/data-binding.html ***!
  \********************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/shared.js */ "./src/shared/shared.js?7c7c"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/visualization.js */ "./src/shared/visualization.js"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_2___ = new URL(/* asset import */ __webpack_require__(/*! ./data-binding.js */ "./src/data-binding/data-binding.js?04f7"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_3___ = new URL(/* asset import */ __webpack_require__(/*! ../shared/shared.css */ "./src/shared/shared.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_4___ = new URL(/* asset import */ __webpack_require__(/*! ./data-binding.css */ "./src/data-binding/data-binding.css"), __webpack_require__.b);
// Module
var code = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <title></title>\r\n    <" + "script src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\" type=\"text/javascript\"><" + "/script>\r\n\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_0___ + "\" type=\"text/javascript\"><" + "/script>\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_1___ + "\" type=\"text/javascript\"><" + "/script>\r\n    <" + "script src=\"" + ___HTML_LOADER_IMPORT_2___ + "\" type=\"text/javascript\"><" + "/script>\r\n\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_3___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_4___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n<body dir=\"ltr\">\r\n    <div id=\"content-main\">\r\n        <div class=\"padding\">\r\n            <br />\r\n            <h1>Bind to data</h1>\r\n            <p>Please choose from one of the following options.</p>\r\n            <ul>\r\n                <li>Bind to existing data and display it\r\n                    <span><button id=\"bind-to-existing-data\">Bind to existing data</button></span>\r\n                </li>\r\n                <li>Insert sample data and display visualization\r\n                    <p id=\"insert-data-available\" style=\"display: none;\"><button id=\"insert-sample-data\">Insert sample data</button></p>\r\n                    <p id=\"insert-data-unavailable\"><i>Option not supported on this Excel version!</i></p>\r\n                </li>\r\n            </ul>\r\n        </div>\r\n    </div>\r\n    <div id=\"content-footer\">\r\n        <div class=\"padding\">\r\n            <a href=\"home.html\">Back to Visualization</a>\r\n        </div>\r\n    </div>\r\n    <img src=\"https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-content-add-in-data-visualization-binding-run\" />\r\n</body>\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=databinding.js.map