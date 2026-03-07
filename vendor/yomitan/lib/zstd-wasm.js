// @ts-nocheck

// node_modules/@bokuweb/zstd-wasm/dist/web/zstd.js
var Module = typeof Module !== "undefined" ? Module : {};
var moduleOverrides = {};
var key;
for (key in Module) {
  if (Module.hasOwnProperty(key)) {
    moduleOverrides[key] = Module[key];
  }
}
var arguments_ = [];
var err = Module["printErr"] || console.warn.bind(console);
for (key in moduleOverrides) {
  if (moduleOverrides.hasOwnProperty(key)) {
    Module[key] = moduleOverrides[key];
  }
}
var quit_ = (status, toThrow) => {
  throw toThrow;
};
moduleOverrides = null;
if (Module["arguments"])
  arguments_ = Module["arguments"];
if (Module["thisProgram"])
  thisProgram = Module["thisProgram"];
if (Module["quit"])
  quit_ = Module["quit"];
if (typeof WebAssembly !== "object") {
  abort("no native wasm support detected");
}
var wasmMemory;
var ABORT = false;
var EXITSTATUS;
var HEAPU8;
var HEAP8;
function updateMemoryViews() {
  var b = wasmMemory.buffer;
  Module["HEAP8"] = HEAP8 = new Int8Array(b);
  Module["HEAPU8"] = HEAPU8 = new Uint8Array(b);
}
var __ATPRERUN__ = [];
var __ATINIT__ = [];
var __ATPOSTRUN__ = [];
var runtimeInitialized = false;
function preRun() {
  if (Module["preRun"]) {
    if (typeof Module["preRun"] == "function")
      Module["preRun"] = [Module["preRun"]];
    while (Module["preRun"].length) {
      addOnPreRun(Module["preRun"].shift());
    }
  }
  callRuntimeCallbacks(__ATPRERUN__);
}
function initRuntime() {
  runtimeInitialized = true;
  callRuntimeCallbacks(__ATINIT__);
}
function postRun() {
  if (Module["postRun"]) {
    if (typeof Module["postRun"] == "function")
      Module["postRun"] = [Module["postRun"]];
    while (Module["postRun"].length) {
      addOnPostRun(Module["postRun"].shift());
    }
  }
  callRuntimeCallbacks(__ATPOSTRUN__);
}
function addOnPreRun(cb) {
  __ATPRERUN__.unshift(cb);
}
function addOnInit(cb) {
  __ATINIT__.unshift(cb);
}
function addOnPostRun(cb) {
  __ATPOSTRUN__.unshift(cb);
}
var runDependencies = 0;
var dependenciesFulfilled = null;
function addRunDependency(id) {
  var _a;
  runDependencies++;
  (_a = Module["monitorRunDependencies"]) === null || _a === void 0 ? void 0 : _a.call(Module, runDependencies);
}
function removeRunDependency(id) {
  var _a;
  runDependencies--;
  (_a = Module["monitorRunDependencies"]) === null || _a === void 0 ? void 0 : _a.call(Module, runDependencies);
  if (runDependencies == 0) {
    if (dependenciesFulfilled) {
      var callback = dependenciesFulfilled;
      dependenciesFulfilled = null;
      callback();
    }
  }
}
function abort(what) {
  var _a;
  (_a = Module["onAbort"]) === null || _a === void 0 ? void 0 : _a.call(Module, what);
  what = "Aborted(" + what + ")";
  err(what);
  ABORT = true;
  what += ". Build with -sASSERTIONS for more info.";
  var e = new WebAssembly.RuntimeError(what);
  throw e;
}
function getWasmImports() {
  return { a: wasmImports };
}
function getBinaryPromise(url) {
  return fetch(url, { credentials: "same-origin" }).then(function(response) {
    if (!response["ok"]) {
      throw "failed to load wasm binary file at '" + url + "'";
    }
    return response["arrayBuffer"]();
  });
}
function init(filePathOrBuf) {
  var info = getWasmImports();
  function receiveInstance(instance, module) {
    wasmExports = instance.exports;
    wasmMemory = wasmExports["f"];
    updateMemoryViews();
    addOnInit(wasmExports["g"]);
    removeRunDependency("wasm-instantiate");
    return wasmExports;
  }
  addRunDependency("wasm-instantiate");
  function receiveInstantiationResult(result) {
    receiveInstance(result["instance"]);
  }
  function instantiateArrayBuffer(receiver) {
    return getBinaryPromise(filePathOrBuf).then(function(binary) {
      var result = WebAssembly.instantiate(binary, info);
      return result;
    }).then(receiver, function(reason) {
      err("failed to asynchronously prepare wasm: " + reason);
      abort(reason);
    });
  }
  function instantiateAsync() {
    if (filePathOrBuf && filePathOrBuf.byteLength > 0) {
      return WebAssembly.instantiate(filePathOrBuf, info).then(receiveInstantiationResult, function(reason) {
        err("wasm compile failed: " + reason);
      });
    } else if (typeof WebAssembly.instantiateStreaming === "function" && typeof filePathOrBuf === "string" && typeof fetch === "function") {
      return fetch(filePathOrBuf, { credentials: "same-origin" }).then(function(response) {
        var result = WebAssembly.instantiateStreaming(response, info);
        return result.then(receiveInstantiationResult, function(reason) {
          err("wasm streaming compile failed: " + reason);
          err("falling back to ArrayBuffer instantiation");
          return instantiateArrayBuffer(receiveInstantiationResult);
        });
      });
    } else {
      return instantiateArrayBuffer(receiveInstantiationResult);
    }
  }
  if (Module["instantiateWasm"]) {
    try {
      var exports = Module["instantiateWasm"](info, receiveInstance);
      return exports;
    } catch (e) {
      err("Module.instantiateWasm callback failed with error: " + e);
      return false;
    }
  }
  instantiateAsync();
  return {};
}
var ExitStatus = class {
  constructor(status) {
    this.name = "ExitStatus";
    this.message = `Program terminated with exit(${status})`;
    this.status = status;
  }
};
var callRuntimeCallbacks = (callbacks) => {
  while (callbacks.length > 0) {
    callbacks.shift()(Module);
  }
};
var noExitRuntime = Module["noExitRuntime"] || true;
var __abort_js = () => abort("");
var runtimeKeepaliveCounter = 0;
var __emscripten_runtime_keepalive_clear = () => {
  noExitRuntime = false;
  runtimeKeepaliveCounter = 0;
};
var timers = {};
var handleException = (e) => {
  if (e instanceof ExitStatus || e == "unwind") {
    return EXITSTATUS;
  }
  quit_(1, e);
};
var keepRuntimeAlive = () => noExitRuntime || runtimeKeepaliveCounter > 0;
var _proc_exit = (code) => {
  var _a;
  EXITSTATUS = code;
  if (!keepRuntimeAlive()) {
    (_a = Module["onExit"]) === null || _a === void 0 ? void 0 : _a.call(Module, code);
    ABORT = true;
  }
  quit_(code, new ExitStatus(code));
};
var exitJS = (status, implicit) => {
  EXITSTATUS = status;
  _proc_exit(status);
};
var _exit = exitJS;
var maybeExit = () => {
  if (!keepRuntimeAlive()) {
    try {
      _exit(EXITSTATUS);
    } catch (e) {
      handleException(e);
    }
  }
};
var callUserCallback = (func) => {
  if (ABORT) {
    return;
  }
  try {
    func();
    maybeExit();
  } catch (e) {
    handleException(e);
  }
};
var _emscripten_get_now = () => performance.now();
var __setitimer_js = (which, timeout_ms) => {
  if (timers[which]) {
    clearTimeout(timers[which].id);
    delete timers[which];
  }
  if (!timeout_ms)
    return 0;
  var id = setTimeout(() => {
    delete timers[which];
    callUserCallback(() => __emscripten_timeout(which, _emscripten_get_now()));
  }, timeout_ms);
  timers[which] = { id, timeout_ms };
  return 0;
};
var getHeapMax = () => 2147483648;
var alignMemory = (size, alignment) => Math.ceil(size / alignment) * alignment;
var growMemory = (size) => {
  var b = wasmMemory.buffer;
  var pages = (size - b.byteLength + 65535) / 65536 | 0;
  try {
    wasmMemory.grow(pages);
    updateMemoryViews();
    return 1;
  } catch (e) {
  }
};
var _emscripten_resize_heap = (requestedSize) => {
  var oldSize = HEAPU8.length;
  requestedSize >>>= 0;
  var maxHeapSize = getHeapMax();
  if (requestedSize > maxHeapSize) {
    return false;
  }
  for (var cutDown = 1; cutDown <= 4; cutDown *= 2) {
    var overGrownHeapSize = oldSize * (1 + 0.2 / cutDown);
    overGrownHeapSize = Math.min(overGrownHeapSize, requestedSize + 100663296);
    var newSize = Math.min(maxHeapSize, alignMemory(Math.max(requestedSize, overGrownHeapSize), 65536));
    var replacement = growMemory(newSize);
    if (replacement) {
      return true;
    }
  }
  return false;
};
var wasmImports = {
  c: __abort_js,
  b: __emscripten_runtime_keepalive_clear,
  d: __setitimer_js,
  e: _emscripten_resize_heap,
  a: _proc_exit
};
var wasmExports;
var _ZSTD_isError = Module["_ZSTD_isError"] = (a0) => (_ZSTD_isError = Module["_ZSTD_isError"] = wasmExports["h"])(a0);
var _ZSTD_compressBound = Module["_ZSTD_compressBound"] = (a0) => (_ZSTD_compressBound = Module["_ZSTD_compressBound"] = wasmExports["i"])(a0);
var _ZSTD_createCCtx = Module["_ZSTD_createCCtx"] = () => (_ZSTD_createCCtx = Module["_ZSTD_createCCtx"] = wasmExports["j"])();
var _ZSTD_freeCCtx = Module["_ZSTD_freeCCtx"] = (a0) => (_ZSTD_freeCCtx = Module["_ZSTD_freeCCtx"] = wasmExports["k"])(a0);
var _ZSTD_compress_usingDict = Module["_ZSTD_compress_usingDict"] = (a0, a1, a2, a3, a4, a5, a6, a7) => (_ZSTD_compress_usingDict = Module["_ZSTD_compress_usingDict"] = wasmExports["l"])(a0, a1, a2, a3, a4, a5, a6, a7);
var _ZSTD_compress = Module["_ZSTD_compress"] = (a0, a1, a2, a3, a4) => (_ZSTD_compress = Module["_ZSTD_compress"] = wasmExports["m"])(a0, a1, a2, a3, a4);
var _ZSTD_createDCtx = Module["_ZSTD_createDCtx"] = () => (_ZSTD_createDCtx = Module["_ZSTD_createDCtx"] = wasmExports["n"])();
var _ZSTD_freeDCtx = Module["_ZSTD_freeDCtx"] = (a0) => (_ZSTD_freeDCtx = Module["_ZSTD_freeDCtx"] = wasmExports["o"])(a0);
var _ZSTD_getFrameContentSize = Module["_ZSTD_getFrameContentSize"] = (a0, a1) => (_ZSTD_getFrameContentSize = Module["_ZSTD_getFrameContentSize"] = wasmExports["p"])(a0, a1);
var _ZSTD_decompress_usingDict = Module["_ZSTD_decompress_usingDict"] = (a0, a1, a2, a3, a4, a5, a6) => (_ZSTD_decompress_usingDict = Module["_ZSTD_decompress_usingDict"] = wasmExports["q"])(a0, a1, a2, a3, a4, a5, a6);
var _ZSTD_decompress = Module["_ZSTD_decompress"] = (a0, a1, a2, a3) => (_ZSTD_decompress = Module["_ZSTD_decompress"] = wasmExports["r"])(a0, a1, a2, a3);
var _malloc = Module["_malloc"] = (a0) => (_malloc = Module["_malloc"] = wasmExports["s"])(a0);
var _free = Module["_free"] = (a0) => (_free = Module["_free"] = wasmExports["t"])(a0);
var __emscripten_timeout = (a0, a1) => (__emscripten_timeout = wasmExports["v"])(a0, a1);
var calledRun;
dependenciesFulfilled = function runCaller() {
  if (!calledRun)
    run();
  if (!calledRun)
    dependenciesFulfilled = runCaller;
};
function run() {
  if (runDependencies > 0) {
    return;
  }
  preRun();
  if (runDependencies > 0) {
    return;
  }
  function doRun() {
    var _a;
    if (calledRun)
      return;
    calledRun = true;
    Module["calledRun"] = true;
    if (ABORT)
      return;
    initRuntime();
    (_a = Module["onRuntimeInitialized"]) === null || _a === void 0 ? void 0 : _a.call(Module);
    postRun();
  }
  if (Module["setStatus"]) {
    Module["setStatus"]("Running...");
    setTimeout(() => {
      setTimeout(() => Module["setStatus"](""), 1);
      doRun();
    }, 1);
  } else {
    doRun();
  }
}
Module["run"] = run;
if (Module["preInit"]) {
  if (typeof Module["preInit"] == "function")
    Module["preInit"] = [Module["preInit"]];
  while (Module["preInit"].length > 0) {
    Module["preInit"].pop()();
  }
}
Module["init"] = init;

// node_modules/@bokuweb/zstd-wasm/dist/web/module.js
var __awaiter = function(thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function(resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function(resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var initialized = (() => new Promise((resolve) => {
  Module.onRuntimeInitialized = resolve;
}))();
var waitInitialized = () => __awaiter(void 0, void 0, void 0, function* () {
  yield initialized;
});

// node_modules/@bokuweb/zstd-wasm/dist/web/errors/index.js
var isError = (code) => {
  const _isError = Module["_ZSTD_isError"];
  return _isError(code);
};

// node_modules/@bokuweb/zstd-wasm/dist/web/simple/decompress.js
var getFrameContentSize = (src, size) => {
  const getSize = Module["_ZSTD_getFrameContentSize"];
  return getSize(src, size);
};
var decompress = (buf, opts = { defaultHeapSize: 1024 * 1024 }) => {
  const malloc = Module["_malloc"];
  const src = malloc(buf.byteLength);
  Module.HEAP8.set(buf, src);
  const contentSize = getFrameContentSize(src, buf.byteLength);
  const size = contentSize === -1 ? opts.defaultHeapSize : contentSize;
  const free = Module["_free"];
  const heap = malloc(size);
  try {
    const _decompress = Module["_ZSTD_decompress"];
    const sizeOrError = _decompress(heap, size, src, buf.byteLength);
    if (isError(sizeOrError)) {
      throw new Error(`Failed to compress with code ${sizeOrError}`);
    }
    const data = new Uint8Array(Module.HEAPU8.buffer, heap, sizeOrError).slice();
    free(heap, size);
    free(src, buf.byteLength);
    return data;
  } catch (e) {
    free(heap, size);
    free(src, buf.byteLength);
    throw e;
  }
};

// node_modules/@bokuweb/zstd-wasm/dist/web/simple/compress.js
var compressBound = (size) => {
  const bound = Module["_ZSTD_compressBound"];
  return bound(size);
};
var compress = (buf, level) => {
  const bound = compressBound(buf.byteLength);
  const malloc = Module["_malloc"];
  const compressed = malloc(bound);
  const src = malloc(buf.byteLength);
  Module.HEAP8.set(buf, src);
  const free = Module["_free"];
  try {
    const _compress = Module["_ZSTD_compress"];
    const sizeOrError = _compress(compressed, bound, src, buf.byteLength, level !== null && level !== void 0 ? level : 3);
    if (isError(sizeOrError)) {
      throw new Error(`Failed to compress with code ${sizeOrError}`);
    }
    const data = new Uint8Array(Module.HEAPU8.buffer, compressed, sizeOrError).slice();
    free(compressed, bound);
    free(src, buf.byteLength);
    return data;
  } catch (e) {
    free(compressed, bound);
    free(src, buf.byteLength);
    throw e;
  }
};

// node_modules/@bokuweb/zstd-wasm/dist/web/simple/decompress_using_dict.js
var getFrameContentSize2 = (src, size) => {
  const getSize = Module["_ZSTD_getFrameContentSize"];
  return getSize(src, size);
};
var createDCtx = () => {
  return Module["_ZSTD_createDCtx"]();
};
var freeDCtx = (dctx) => {
  return Module["_ZSTD_freeDCtx"](dctx);
};
var decompressUsingDict = (dctx, buf, dict, opts = { defaultHeapSize: 1024 * 1024 }) => {
  const malloc = Module["_malloc"];
  const src = malloc(buf.byteLength);
  Module.HEAP8.set(buf, src);
  const pdict = malloc(dict.byteLength);
  Module.HEAP8.set(dict, pdict);
  const contentSize = getFrameContentSize2(src, buf.byteLength);
  const size = contentSize === -1 ? opts.defaultHeapSize : contentSize;
  const free = Module["_free"];
  const heap = malloc(size);
  try {
    const _decompress = Module["_ZSTD_decompress_usingDict"];
    const sizeOrError = _decompress(dctx, heap, size, src, buf.byteLength, pdict, dict.byteLength);
    if (isError(sizeOrError)) {
      throw new Error(`Failed to compress with code ${sizeOrError}`);
    }
    const data = new Uint8Array(Module.HEAPU8.buffer, heap, sizeOrError).slice();
    free(heap, size);
    free(src, buf.byteLength);
    free(pdict, dict.byteLength);
    return data;
  } catch (e) {
    free(heap, size);
    free(src, buf.byteLength);
    free(pdict, dict.byteLength);
    throw e;
  }
};

// node_modules/@bokuweb/zstd-wasm/dist/web/simple/compress_using_dict.js
var compressBound2 = (size) => {
  const bound = Module["_ZSTD_compressBound"];
  return bound(size);
};
var createCCtx = () => {
  return Module["_ZSTD_createCCtx"]();
};
var freeCCtx = (cctx) => {
  return Module["_ZSTD_freeCCtx"](cctx);
};
var compressUsingDict = (cctx, buf, dict, level) => {
  const bound = compressBound2(buf.byteLength);
  const malloc = Module["_malloc"];
  const compressed = malloc(bound);
  const src = malloc(buf.byteLength);
  Module.HEAP8.set(buf, src);
  const pdict = malloc(dict.byteLength);
  Module.HEAP8.set(dict, pdict);
  const free = Module["_free"];
  try {
    const _compress = Module["_ZSTD_compress_usingDict"];
    const sizeOrError = _compress(cctx, compressed, bound, src, buf.byteLength, pdict, dict.byteLength, level !== null && level !== void 0 ? level : 3);
    if (isError(sizeOrError)) {
      throw new Error(`Failed to compress with code ${sizeOrError}`);
    }
    const data = new Uint8Array(Module.HEAPU8.buffer, compressed, sizeOrError).slice();
    free(compressed, bound);
    free(src, buf.byteLength);
    free(pdict, dict.byteLength);
    return data;
  } catch (e) {
    free(compressed, bound);
    free(src, buf.byteLength);
    free(pdict, dict.byteLength);
    throw e;
  }
};

// node_modules/@bokuweb/zstd-wasm/dist/web/index.web.js
var __awaiter2 = function(thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function(resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function(resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var init2 = (path) => __awaiter2(void 0, void 0, void 0, function* () {
  const url = new URL(`./zstd.wasm`, import.meta.url).href;
  Module["init"](path !== null && path !== void 0 ? path : url);
  yield waitInitialized();
});
export {
  compress,
  compressUsingDict,
  createCCtx,
  createDCtx,
  decompress,
  decompressUsingDict,
  freeCCtx,
  freeDCtx,
  init2 as init
};
//# sourceMappingURL=zstd-wasm.js.map
