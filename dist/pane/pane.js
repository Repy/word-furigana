/******/ (() => { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/pane/pane.ts":
/*!**************************!*\
  !*** ./src/pane/pane.ts ***!
  \**************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : __webpack_require__.g;
}
const g = getGlobal();
Office.onReady(() => {
    g.document.getElementById("button").addEventListener("click", () => tryCatch(addRubi), false);
});
function addRubi() {
    return __awaiter(this, void 0, void 0, function* () {
        yield Word.run((context) => __awaiter(this, void 0, void 0, function* () {
            const range = context.document.getSelection();
            range.load("text");
            yield context.sync();
            const rubidata = rubi(range.text);
            let nowRange = range;
            for (const iterator of rubidata) {
                const text = iterator.s;
                const rubitext = iterator.r;
                const code = "\\* jc2 \\* hps10 \\o(\\s\\up9(" + rubitext + ")," + text + ")";
                console.log(code);
                nowRange.insertField(Word.InsertLocation.before, Word.FieldType.eq, code.trim(), true);
            }
            range.clear();
            yield context.sync();
        }));
        yield Word.run((context) => __awaiter(this, void 0, void 0, function* () {
            const range = context.document.getSelection();
            range.load("fields");
            yield context.sync();
            for (const iterator of range.fields.items) {
                iterator.code = iterator.code.trim();
            }
            yield context.sync();
        }));
    });
}
function tryCatch(callback) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield callback();
        }
        catch (error) {
            console.error(error);
        }
    });
}


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = {};
/******/ 	__webpack_modules__["./src/pane/pane.ts"](0, __webpack_exports__, __webpack_require__);
/******/ 	
/******/ })()
;
//# sourceMappingURL=pane.js.map