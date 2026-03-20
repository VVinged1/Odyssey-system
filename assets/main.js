var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __commonJS = (cb, mod) => function __require() {
  return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// node_modules/events/events.js
var require_events = __commonJS({
  "node_modules/events/events.js"(exports, module) {
    "use strict";
    var R = typeof Reflect === "object" ? Reflect : null;
    var ReflectApply = R && typeof R.apply === "function" ? R.apply : function ReflectApply2(target, receiver, args) {
      return Function.prototype.apply.call(target, receiver, args);
    };
    var ReflectOwnKeys;
    if (R && typeof R.ownKeys === "function") {
      ReflectOwnKeys = R.ownKeys;
    } else if (Object.getOwnPropertySymbols) {
      ReflectOwnKeys = function ReflectOwnKeys2(target) {
        return Object.getOwnPropertyNames(target).concat(Object.getOwnPropertySymbols(target));
      };
    } else {
      ReflectOwnKeys = function ReflectOwnKeys2(target) {
        return Object.getOwnPropertyNames(target);
      };
    }
    function ProcessEmitWarning(warning) {
      if (console && console.warn) console.warn(warning);
    }
    var NumberIsNaN = Number.isNaN || function NumberIsNaN2(value) {
      return value !== value;
    };
    function EventEmitter2() {
      EventEmitter2.init.call(this);
    }
    module.exports = EventEmitter2;
    module.exports.once = once;
    EventEmitter2.EventEmitter = EventEmitter2;
    EventEmitter2.prototype._events = void 0;
    EventEmitter2.prototype._eventsCount = 0;
    EventEmitter2.prototype._maxListeners = void 0;
    var defaultMaxListeners = 10;
    function checkListener(listener) {
      if (typeof listener !== "function") {
        throw new TypeError('The "listener" argument must be of type Function. Received type ' + typeof listener);
      }
    }
    Object.defineProperty(EventEmitter2, "defaultMaxListeners", {
      enumerable: true,
      get: function() {
        return defaultMaxListeners;
      },
      set: function(arg) {
        if (typeof arg !== "number" || arg < 0 || NumberIsNaN(arg)) {
          throw new RangeError('The value of "defaultMaxListeners" is out of range. It must be a non-negative number. Received ' + arg + ".");
        }
        defaultMaxListeners = arg;
      }
    });
    EventEmitter2.init = function() {
      if (this._events === void 0 || this._events === Object.getPrototypeOf(this)._events) {
        this._events = /* @__PURE__ */ Object.create(null);
        this._eventsCount = 0;
      }
      this._maxListeners = this._maxListeners || void 0;
    };
    EventEmitter2.prototype.setMaxListeners = function setMaxListeners(n) {
      if (typeof n !== "number" || n < 0 || NumberIsNaN(n)) {
        throw new RangeError('The value of "n" is out of range. It must be a non-negative number. Received ' + n + ".");
      }
      this._maxListeners = n;
      return this;
    };
    function _getMaxListeners(that) {
      if (that._maxListeners === void 0)
        return EventEmitter2.defaultMaxListeners;
      return that._maxListeners;
    }
    EventEmitter2.prototype.getMaxListeners = function getMaxListeners() {
      return _getMaxListeners(this);
    };
    EventEmitter2.prototype.emit = function emit(type) {
      var args = [];
      for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);
      var doError = type === "error";
      var events = this._events;
      if (events !== void 0)
        doError = doError && events.error === void 0;
      else if (!doError)
        return false;
      if (doError) {
        var er;
        if (args.length > 0)
          er = args[0];
        if (er instanceof Error) {
          throw er;
        }
        var err = new Error("Unhandled error." + (er ? " (" + er.message + ")" : ""));
        err.context = er;
        throw err;
      }
      var handler = events[type];
      if (handler === void 0)
        return false;
      if (typeof handler === "function") {
        ReflectApply(handler, this, args);
      } else {
        var len = handler.length;
        var listeners = arrayClone(handler, len);
        for (var i = 0; i < len; ++i)
          ReflectApply(listeners[i], this, args);
      }
      return true;
    };
    function _addListener(target, type, listener, prepend) {
      var m;
      var events;
      var existing;
      checkListener(listener);
      events = target._events;
      if (events === void 0) {
        events = target._events = /* @__PURE__ */ Object.create(null);
        target._eventsCount = 0;
      } else {
        if (events.newListener !== void 0) {
          target.emit(
            "newListener",
            type,
            listener.listener ? listener.listener : listener
          );
          events = target._events;
        }
        existing = events[type];
      }
      if (existing === void 0) {
        existing = events[type] = listener;
        ++target._eventsCount;
      } else {
        if (typeof existing === "function") {
          existing = events[type] = prepend ? [listener, existing] : [existing, listener];
        } else if (prepend) {
          existing.unshift(listener);
        } else {
          existing.push(listener);
        }
        m = _getMaxListeners(target);
        if (m > 0 && existing.length > m && !existing.warned) {
          existing.warned = true;
          var w = new Error("Possible EventEmitter memory leak detected. " + existing.length + " " + String(type) + " listeners added. Use emitter.setMaxListeners() to increase limit");
          w.name = "MaxListenersExceededWarning";
          w.emitter = target;
          w.type = type;
          w.count = existing.length;
          ProcessEmitWarning(w);
        }
      }
      return target;
    }
    EventEmitter2.prototype.addListener = function addListener(type, listener) {
      return _addListener(this, type, listener, false);
    };
    EventEmitter2.prototype.on = EventEmitter2.prototype.addListener;
    EventEmitter2.prototype.prependListener = function prependListener(type, listener) {
      return _addListener(this, type, listener, true);
    };
    function onceWrapper() {
      if (!this.fired) {
        this.target.removeListener(this.type, this.wrapFn);
        this.fired = true;
        if (arguments.length === 0)
          return this.listener.call(this.target);
        return this.listener.apply(this.target, arguments);
      }
    }
    function _onceWrap(target, type, listener) {
      var state = { fired: false, wrapFn: void 0, target, type, listener };
      var wrapped = onceWrapper.bind(state);
      wrapped.listener = listener;
      state.wrapFn = wrapped;
      return wrapped;
    }
    EventEmitter2.prototype.once = function once2(type, listener) {
      checkListener(listener);
      this.on(type, _onceWrap(this, type, listener));
      return this;
    };
    EventEmitter2.prototype.prependOnceListener = function prependOnceListener(type, listener) {
      checkListener(listener);
      this.prependListener(type, _onceWrap(this, type, listener));
      return this;
    };
    EventEmitter2.prototype.removeListener = function removeListener(type, listener) {
      var list, events, position, i, originalListener;
      checkListener(listener);
      events = this._events;
      if (events === void 0)
        return this;
      list = events[type];
      if (list === void 0)
        return this;
      if (list === listener || list.listener === listener) {
        if (--this._eventsCount === 0)
          this._events = /* @__PURE__ */ Object.create(null);
        else {
          delete events[type];
          if (events.removeListener)
            this.emit("removeListener", type, list.listener || listener);
        }
      } else if (typeof list !== "function") {
        position = -1;
        for (i = list.length - 1; i >= 0; i--) {
          if (list[i] === listener || list[i].listener === listener) {
            originalListener = list[i].listener;
            position = i;
            break;
          }
        }
        if (position < 0)
          return this;
        if (position === 0)
          list.shift();
        else {
          spliceOne(list, position);
        }
        if (list.length === 1)
          events[type] = list[0];
        if (events.removeListener !== void 0)
          this.emit("removeListener", type, originalListener || listener);
      }
      return this;
    };
    EventEmitter2.prototype.off = EventEmitter2.prototype.removeListener;
    EventEmitter2.prototype.removeAllListeners = function removeAllListeners(type) {
      var listeners, events, i;
      events = this._events;
      if (events === void 0)
        return this;
      if (events.removeListener === void 0) {
        if (arguments.length === 0) {
          this._events = /* @__PURE__ */ Object.create(null);
          this._eventsCount = 0;
        } else if (events[type] !== void 0) {
          if (--this._eventsCount === 0)
            this._events = /* @__PURE__ */ Object.create(null);
          else
            delete events[type];
        }
        return this;
      }
      if (arguments.length === 0) {
        var keys = Object.keys(events);
        var key;
        for (i = 0; i < keys.length; ++i) {
          key = keys[i];
          if (key === "removeListener") continue;
          this.removeAllListeners(key);
        }
        this.removeAllListeners("removeListener");
        this._events = /* @__PURE__ */ Object.create(null);
        this._eventsCount = 0;
        return this;
      }
      listeners = events[type];
      if (typeof listeners === "function") {
        this.removeListener(type, listeners);
      } else if (listeners !== void 0) {
        for (i = listeners.length - 1; i >= 0; i--) {
          this.removeListener(type, listeners[i]);
        }
      }
      return this;
    };
    function _listeners(target, type, unwrap) {
      var events = target._events;
      if (events === void 0)
        return [];
      var evlistener = events[type];
      if (evlistener === void 0)
        return [];
      if (typeof evlistener === "function")
        return unwrap ? [evlistener.listener || evlistener] : [evlistener];
      return unwrap ? unwrapListeners(evlistener) : arrayClone(evlistener, evlistener.length);
    }
    EventEmitter2.prototype.listeners = function listeners(type) {
      return _listeners(this, type, true);
    };
    EventEmitter2.prototype.rawListeners = function rawListeners(type) {
      return _listeners(this, type, false);
    };
    EventEmitter2.listenerCount = function(emitter, type) {
      if (typeof emitter.listenerCount === "function") {
        return emitter.listenerCount(type);
      } else {
        return listenerCount.call(emitter, type);
      }
    };
    EventEmitter2.prototype.listenerCount = listenerCount;
    function listenerCount(type) {
      var events = this._events;
      if (events !== void 0) {
        var evlistener = events[type];
        if (typeof evlistener === "function") {
          return 1;
        } else if (evlistener !== void 0) {
          return evlistener.length;
        }
      }
      return 0;
    }
    EventEmitter2.prototype.eventNames = function eventNames() {
      return this._eventsCount > 0 ? ReflectOwnKeys(this._events) : [];
    };
    function arrayClone(arr, n) {
      var copy = new Array(n);
      for (var i = 0; i < n; ++i)
        copy[i] = arr[i];
      return copy;
    }
    function spliceOne(list, index) {
      for (; index + 1 < list.length; index++)
        list[index] = list[index + 1];
      list.pop();
    }
    function unwrapListeners(arr) {
      var ret = new Array(arr.length);
      for (var i = 0; i < ret.length; ++i) {
        ret[i] = arr[i].listener || arr[i];
      }
      return ret;
    }
    function once(emitter, name) {
      return new Promise(function(resolve, reject) {
        function errorListener(err) {
          emitter.removeListener(name, resolver);
          reject(err);
        }
        function resolver() {
          if (typeof emitter.removeListener === "function") {
            emitter.removeListener("error", errorListener);
          }
          resolve([].slice.call(arguments));
        }
        ;
        eventTargetAgnosticAddListener(emitter, name, resolver, { once: true });
        if (name !== "error") {
          addErrorHandlerIfEventEmitter(emitter, errorListener, { once: true });
        }
      });
    }
    function addErrorHandlerIfEventEmitter(emitter, handler, flags) {
      if (typeof emitter.on === "function") {
        eventTargetAgnosticAddListener(emitter, "error", handler, flags);
      }
    }
    function eventTargetAgnosticAddListener(emitter, name, listener, flags) {
      if (typeof emitter.on === "function") {
        if (flags.once) {
          emitter.once(name, listener);
        } else {
          emitter.on(name, listener);
        }
      } else if (typeof emitter.addEventListener === "function") {
        emitter.addEventListener(name, function wrapListener(arg) {
          if (flags.once) {
            emitter.removeEventListener(name, wrapListener);
          }
          listener(arg);
        });
      } else {
        throw new TypeError('The "emitter" argument must be of type EventEmitter. Received type ' + typeof emitter);
      }
    }
  }
});

// node_modules/@owlbear-rodeo/sdk/lib/api/PlayerApi.js
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
var PlayerApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  get id() {
    if (!this.messageBus.userId) {
      throw Error("Unable to get user ID: not ready");
    }
    return this.messageBus.userId;
  }
  getSelection() {
    return __awaiter(this, void 0, void 0, function* () {
      const { selection } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_SELECTION", {});
      return selection;
    });
  }
  select(items, replace) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_SELECT", { items, replace });
    });
  }
  deselect(items) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_DESELECT", { items });
    });
  }
  getName() {
    return __awaiter(this, void 0, void 0, function* () {
      const { name } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_NAME", {});
      return name;
    });
  }
  setName(name) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_SET_NAME", { name });
    });
  }
  getColor() {
    return __awaiter(this, void 0, void 0, function* () {
      const { color } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_COLOR", {});
      return color;
    });
  }
  setColor(color) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_SET_COLOR", { color });
    });
  }
  getSyncView() {
    return __awaiter(this, void 0, void 0, function* () {
      const { syncView } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_SYNC_VIEW", {});
      return syncView;
    });
  }
  setSyncView(syncView) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_SET_SYNC_VIEW", { syncView });
    });
  }
  getId() {
    return __awaiter(this, void 0, void 0, function* () {
      const { id } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_ID", {});
      return id;
    });
  }
  getRole() {
    return __awaiter(this, void 0, void 0, function* () {
      const { role } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_ROLE", {});
      return role;
    });
  }
  getMetadata() {
    return __awaiter(this, void 0, void 0, function* () {
      const { metadata } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_METADATA", {});
      return metadata;
    });
  }
  setMetadata(update) {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_PLAYER_SET_METADATA", { update });
    });
  }
  hasPermission(permission) {
    return __awaiter(this, void 0, void 0, function* () {
      const role = yield this.getRole();
      if (role === "GM") {
        return true;
      }
      const { permissions } = yield this.messageBus.sendAsync("OBR_ROOM_GET_PERMISSIONS", {});
      return permissions.indexOf(permission) > -1;
    });
  }
  getConnectionId() {
    return __awaiter(this, void 0, void 0, function* () {
      const { connectionId } = yield this.messageBus.sendAsync("OBR_PLAYER_GET_CONNECTION_ID", {});
      return connectionId;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.player);
    };
    this.messageBus.send("OBR_PLAYER_SUBSCRIBE", {});
    this.messageBus.on("OBR_PLAYER_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_PLAYER_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_PLAYER_EVENT_CHANGE", handleChange);
    };
  }
};
var PlayerApi_default = PlayerApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/ViewportApi.js
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
var ViewportApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  reset() {
    return __awaiter2(this, void 0, void 0, function* () {
      const { transform } = yield this.messageBus.sendAsync("OBR_VIEWPORT_RESET", {});
      return transform;
    });
  }
  animateTo(transform) {
    return __awaiter2(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_VIEWPORT_ANIMATE_TO", { transform });
    });
  }
  animateToBounds(bounds) {
    return __awaiter2(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_VIEWPORT_ANIMATE_TO_BOUNDS", {
        bounds
      });
    });
  }
  getPosition() {
    return __awaiter2(this, void 0, void 0, function* () {
      const { position } = yield this.messageBus.sendAsync("OBR_VIEWPORT_GET_POSITION", {});
      return position;
    });
  }
  setPosition(position) {
    return __awaiter2(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_VIEWPORT_SET_POSITION", { position });
    });
  }
  getScale() {
    return __awaiter2(this, void 0, void 0, function* () {
      const { scale } = yield this.messageBus.sendAsync("OBR_VIEWPORT_GET_SCALE", {});
      return scale;
    });
  }
  setScale(scale) {
    return __awaiter2(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_VIEWPORT_SET_SCALE", { scale });
    });
  }
  getWidth() {
    return __awaiter2(this, void 0, void 0, function* () {
      const { width } = yield this.messageBus.sendAsync("OBR_VIEWPORT_GET_WIDTH", {});
      return width;
    });
  }
  getHeight() {
    return __awaiter2(this, void 0, void 0, function* () {
      const { height } = yield this.messageBus.sendAsync("OBR_VIEWPORT_GET_HEIGHT", {});
      return height;
    });
  }
  transformPoint(point) {
    return __awaiter2(this, void 0, void 0, function* () {
      const { point: transformed } = yield this.messageBus.sendAsync("OBR_VIEWPORT_TRANSFORM_POINT", { point });
      return transformed;
    });
  }
  inverseTransformPoint(point) {
    return __awaiter2(this, void 0, void 0, function* () {
      const { point: transformed } = yield this.messageBus.sendAsync("OBR_VIEWPORT_INVERSE_TRANSFORM_POINT", { point });
      return transformed;
    });
  }
};
var ViewportApi_default = ViewportApi;

// node_modules/@owlbear-rodeo/sdk/lib/messages/Message.js
function isMessage(message) {
  return typeof message.id === "string";
}

// node_modules/@owlbear-rodeo/sdk/lib/messages/MessageBus.js
var import_events = __toESM(require_events());

// node_modules/uuid/dist/esm-browser/rng.js
var getRandomValues;
var rnds8 = new Uint8Array(16);
function rng() {
  if (!getRandomValues) {
    getRandomValues = typeof crypto !== "undefined" && crypto.getRandomValues && crypto.getRandomValues.bind(crypto);
    if (!getRandomValues) {
      throw new Error("crypto.getRandomValues() not supported. See https://github.com/uuidjs/uuid#getrandomvalues-not-supported");
    }
  }
  return getRandomValues(rnds8);
}

// node_modules/uuid/dist/esm-browser/stringify.js
var byteToHex = [];
for (let i = 0; i < 256; ++i) {
  byteToHex.push((i + 256).toString(16).slice(1));
}
function unsafeStringify(arr, offset = 0) {
  return byteToHex[arr[offset + 0]] + byteToHex[arr[offset + 1]] + byteToHex[arr[offset + 2]] + byteToHex[arr[offset + 3]] + "-" + byteToHex[arr[offset + 4]] + byteToHex[arr[offset + 5]] + "-" + byteToHex[arr[offset + 6]] + byteToHex[arr[offset + 7]] + "-" + byteToHex[arr[offset + 8]] + byteToHex[arr[offset + 9]] + "-" + byteToHex[arr[offset + 10]] + byteToHex[arr[offset + 11]] + byteToHex[arr[offset + 12]] + byteToHex[arr[offset + 13]] + byteToHex[arr[offset + 14]] + byteToHex[arr[offset + 15]];
}

// node_modules/uuid/dist/esm-browser/native.js
var randomUUID = typeof crypto !== "undefined" && crypto.randomUUID && crypto.randomUUID.bind(crypto);
var native_default = {
  randomUUID
};

// node_modules/uuid/dist/esm-browser/v4.js
function v4(options, buf, offset) {
  if (native_default.randomUUID && !buf && !options) {
    return native_default.randomUUID();
  }
  options = options || {};
  const rnds = options.random || (options.rng || rng)();
  rnds[6] = rnds[6] & 15 | 64;
  rnds[8] = rnds[8] & 63 | 128;
  if (buf) {
    offset = offset || 0;
    for (let i = 0; i < 16; ++i) {
      buf[offset + i] = rnds[i];
    }
    return buf;
  }
  return unsafeStringify(rnds);
}
var v4_default = v4;

// node_modules/@owlbear-rodeo/sdk/lib/messages/MessageBus.js
var MessageBus = class extends import_events.EventEmitter {
  constructor(origin, roomId) {
    super();
    this.ready = false;
    this.userId = null;
    this.ref = null;
    this.handleMessage = (event) => {
      const message = event.data;
      if (event.origin === this.targetOrigin && isMessage(message)) {
        if (message.id === "OBR_READY") {
          this.ready = true;
          const data = message.data;
          this.ref = data.ref;
          this.userId = data.userId;
        }
        this.emit(message.id, message.data);
      }
    };
    this.send = (id, data, nonce) => {
      var _a;
      if (!this.ref) {
        throw Error("Unable to send message: not ready");
      }
      (_a = window.parent) === null || _a === void 0 ? void 0 : _a.postMessage({
        id,
        data,
        ref: this.ref,
        nonce
      }, this.targetOrigin);
    };
    this.sendAsync = (id, data, timeout = 5e3) => {
      const nonce = `_${v4_default()}`;
      this.send(id, data, nonce);
      return Promise.race([
        new Promise((resolve, reject) => {
          const self = this;
          function onResponse(value) {
            self.off(`${id}_RESPONSE${nonce}`, onResponse);
            self.off(`${id}_ERROR${nonce}`, onError);
            resolve(value);
          }
          function onError(error) {
            self.off(`${id}_RESPONSE${nonce}`, onResponse);
            self.off(`${id}_ERROR${nonce}`, onError);
            reject(error);
          }
          this.on(`${id}_RESPONSE${nonce}`, onResponse);
          this.on(`${id}_ERROR${nonce}`, onError);
        }),
        ...timeout > 0 ? [
          new Promise((_, reject) => window.setTimeout(() => reject(new Error(`Message ${id} took longer than ${timeout}ms to get a result`)), timeout))
        ] : []
      ]);
    };
    this.roomId = roomId;
    this.targetOrigin = origin;
    window.addEventListener("message", this.handleMessage);
    this.setMaxListeners(100);
  }
  destroy() {
    window.removeEventListener("message", this.handleMessage);
  }
};
var MessageBus_default = MessageBus;

// node_modules/@owlbear-rodeo/sdk/lib/api/NotificationApi.js
var __awaiter3 = function(thisArg, _arguments, P, generator) {
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
var NotificationApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  show(message, variant) {
    return __awaiter3(this, void 0, void 0, function* () {
      const { id } = yield this.messageBus.sendAsync("OBR_NOTIFICATION_SHOW", { message, variant });
      return id;
    });
  }
  close(id) {
    return __awaiter3(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_NOTIFICATION_CLOSE", { id });
    });
  }
};
var NotificationApi_default = NotificationApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneFogApi.js
var __awaiter4 = function(thisArg, _arguments, P, generator) {
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
var SceneFogApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getColor() {
    return __awaiter4(this, void 0, void 0, function* () {
      const { color } = yield this.messageBus.sendAsync("OBR_SCENE_FOG_GET_COLOR", {});
      return color;
    });
  }
  setColor(color) {
    return __awaiter4(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_FOG_SET_COLOR", { color });
    });
  }
  getStrokeWidth() {
    return __awaiter4(this, void 0, void 0, function* () {
      const { strokeWidth } = yield this.messageBus.sendAsync("OBR_SCENE_FOG_GET_STROKE_WIDTH", {});
      return strokeWidth;
    });
  }
  setStrokeWidth(strokeWidth) {
    return __awaiter4(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_FOG_SET_STROKE_WIDTH", {
        strokeWidth
      });
    });
  }
  getFilled() {
    return __awaiter4(this, void 0, void 0, function* () {
      const { filled } = yield this.messageBus.sendAsync("OBR_SCENE_FOG_GET_FILLED", {});
      return filled;
    });
  }
  setFilled(filled) {
    return __awaiter4(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_FOG_SET_FILLED", { filled });
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.fog);
    };
    this.messageBus.send("OBR_SCENE_FOG_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_FOG_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_FOG_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_FOG_EVENT_CHANGE", handleChange);
    };
  }
};
var SceneFogApi_default = SceneFogApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneGridApi.js
var __awaiter5 = function(thisArg, _arguments, P, generator) {
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
var SceneGridApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getDpi() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { dpi } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_DPI", {});
      return dpi;
    });
  }
  getScale() {
    return __awaiter5(this, void 0, void 0, function* () {
      const scale = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_SCALE", {});
      return scale;
    });
  }
  setScale(scale) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_SCALE", { scale });
    });
  }
  getColor() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { color } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_COLOR", {});
      return color;
    });
  }
  setColor(color) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_COLOR", { color });
    });
  }
  getOpacity() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { opacity } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_OPACITY", {});
      return opacity;
    });
  }
  setOpacity(opacity) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_OPACITY", { opacity });
    });
  }
  getType() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { type } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_TYPE", {});
      return type;
    });
  }
  setType(type) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_TYPE", { type });
    });
  }
  getLineType() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { lineType } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_LINE_TYPE", {});
      return lineType;
    });
  }
  setLineType(lineType) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_LINE_TYPE", {
        lineType
      });
    });
  }
  getMeasurement() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { measurement } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_MEASUREMENT", {});
      return measurement;
    });
  }
  setMeasurement(measurement) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_MEASUREMENT", {
        measurement
      });
    });
  }
  getLineWidth() {
    return __awaiter5(this, void 0, void 0, function* () {
      const { lineWidth } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_LINE_WIDTH", {});
      return lineWidth;
    });
  }
  setLineWidth(lineWidth) {
    return __awaiter5(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_GRID_SET_LINE_WIDTH", {
        lineWidth
      });
    });
  }
  snapPosition(position, snappingSensitivity, useCorners, useCenter) {
    return __awaiter5(this, void 0, void 0, function* () {
      const { position: snapped } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_SNAP_POSITION", {
        position,
        snappingSensitivity,
        useCorners,
        useCenter
      });
      return snapped;
    });
  }
  getDistance(from, to) {
    return __awaiter5(this, void 0, void 0, function* () {
      const { distance } = yield this.messageBus.sendAsync("OBR_SCENE_GRID_GET_DISTANCE", { from, to });
      return distance;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.grid);
    };
    this.messageBus.send("OBR_SCENE_GRID_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_GRID_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_GRID_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_GRID_EVENT_CHANGE", handleChange);
    };
  }
};
var SceneGridApi_default = SceneGridApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneHistoryApi.js
var __awaiter6 = function(thisArg, _arguments, P, generator) {
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
var SceneHistoryApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  undo() {
    return __awaiter6(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_HISTORY_UNDO", {});
    });
  }
  redo() {
    return __awaiter6(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_HISTORY_REDO", {});
    });
  }
  canUndo() {
    return __awaiter6(this, void 0, void 0, function* () {
      const { canUndo } = yield this.messageBus.sendAsync("OBR_SCENE_HISTORY_CAN_UNDO", {});
      return canUndo;
    });
  }
  canRedo() {
    return __awaiter6(this, void 0, void 0, function* () {
      const { canRedo } = yield this.messageBus.sendAsync("OBR_SCENE_HISTORY_CAN_REDO", {});
      return canRedo;
    });
  }
};
var SceneHistoryApi_default = SceneHistoryApi;

// node_modules/immer/dist/immer.mjs
var NOTHING = Symbol.for("immer-nothing");
var DRAFTABLE = Symbol.for("immer-draftable");
var DRAFT_STATE = Symbol.for("immer-state");
var errors = true ? [
  // All error codes, starting by 0:
  function(plugin) {
    return `The plugin for '${plugin}' has not been loaded into Immer. To enable the plugin, import and call \`enable${plugin}()\` when initializing your application.`;
  },
  function(thing) {
    return `produce can only be called on things that are draftable: plain objects, arrays, Map, Set or classes that are marked with '[immerable]: true'. Got '${thing}'`;
  },
  "This object has been frozen and should not be mutated",
  function(data) {
    return "Cannot use a proxy that has been revoked. Did you pass an object from inside an immer function to an async process? " + data;
  },
  "An immer producer returned a new value *and* modified its draft. Either return a new value *or* modify the draft.",
  "Immer forbids circular references",
  "The first or second argument to `produce` must be a function",
  "The third argument to `produce` must be a function or undefined",
  "First argument to `createDraft` must be a plain object, an array, or an immerable object",
  "First argument to `finishDraft` must be a draft returned by `createDraft`",
  function(thing) {
    return `'current' expects a draft, got: ${thing}`;
  },
  "Object.defineProperty() cannot be used on an Immer draft",
  "Object.setPrototypeOf() cannot be used on an Immer draft",
  "Immer only supports deleting array indices",
  "Immer only supports setting array indices and the 'length' property",
  function(thing) {
    return `'original' expects a draft, got: ${thing}`;
  }
  // Note: if more errors are added, the errorOffset in Patches.ts should be increased
  // See Patches.ts for additional errors
] : [];
function die(error, ...args) {
  if (true) {
    const e = errors[error];
    const msg = typeof e === "function" ? e.apply(null, args) : e;
    throw new Error(`[Immer] ${msg}`);
  }
  throw new Error(
    `[Immer] minified error nr: ${error}. Full error at: https://bit.ly/3cXEKWf`
  );
}
var getPrototypeOf = Object.getPrototypeOf;
function isDraft(value) {
  return !!value && !!value[DRAFT_STATE];
}
function isDraftable(value) {
  if (!value)
    return false;
  return isPlainObject(value) || Array.isArray(value) || !!value[DRAFTABLE] || !!value.constructor?.[DRAFTABLE] || isMap(value) || isSet(value);
}
var objectCtorString = Object.prototype.constructor.toString();
var cachedCtorStrings = /* @__PURE__ */ new WeakMap();
function isPlainObject(value) {
  if (!value || typeof value !== "object")
    return false;
  const proto = Object.getPrototypeOf(value);
  if (proto === null || proto === Object.prototype)
    return true;
  const Ctor = Object.hasOwnProperty.call(proto, "constructor") && proto.constructor;
  if (Ctor === Object)
    return true;
  if (typeof Ctor !== "function")
    return false;
  let ctorString = cachedCtorStrings.get(Ctor);
  if (ctorString === void 0) {
    ctorString = Function.toString.call(Ctor);
    cachedCtorStrings.set(Ctor, ctorString);
  }
  return ctorString === objectCtorString;
}
function each(obj, iter, strict = true) {
  if (getArchtype(obj) === 0) {
    const keys = strict ? Reflect.ownKeys(obj) : Object.keys(obj);
    keys.forEach((key) => {
      iter(key, obj[key], obj);
    });
  } else {
    obj.forEach((entry, index) => iter(index, entry, obj));
  }
}
function getArchtype(thing) {
  const state = thing[DRAFT_STATE];
  return state ? state.type_ : Array.isArray(thing) ? 1 : isMap(thing) ? 2 : isSet(thing) ? 3 : 0;
}
function has(thing, prop) {
  return getArchtype(thing) === 2 ? thing.has(prop) : Object.prototype.hasOwnProperty.call(thing, prop);
}
function get(thing, prop) {
  return getArchtype(thing) === 2 ? thing.get(prop) : thing[prop];
}
function set(thing, propOrOldValue, value) {
  const t = getArchtype(thing);
  if (t === 2)
    thing.set(propOrOldValue, value);
  else if (t === 3) {
    thing.add(value);
  } else
    thing[propOrOldValue] = value;
}
function is(x, y) {
  if (x === y) {
    return x !== 0 || 1 / x === 1 / y;
  } else {
    return x !== x && y !== y;
  }
}
function isMap(target) {
  return target instanceof Map;
}
function isSet(target) {
  return target instanceof Set;
}
function latest(state) {
  return state.copy_ || state.base_;
}
function shallowCopy(base, strict) {
  if (isMap(base)) {
    return new Map(base);
  }
  if (isSet(base)) {
    return new Set(base);
  }
  if (Array.isArray(base))
    return Array.prototype.slice.call(base);
  const isPlain = isPlainObject(base);
  if (strict === true || strict === "class_only" && !isPlain) {
    const descriptors = Object.getOwnPropertyDescriptors(base);
    delete descriptors[DRAFT_STATE];
    let keys = Reflect.ownKeys(descriptors);
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      const desc = descriptors[key];
      if (desc.writable === false) {
        desc.writable = true;
        desc.configurable = true;
      }
      if (desc.get || desc.set)
        descriptors[key] = {
          configurable: true,
          writable: true,
          // could live with !!desc.set as well here...
          enumerable: desc.enumerable,
          value: base[key]
        };
    }
    return Object.create(getPrototypeOf(base), descriptors);
  } else {
    const proto = getPrototypeOf(base);
    if (proto !== null && isPlain) {
      return { ...base };
    }
    const obj = Object.create(proto);
    return Object.assign(obj, base);
  }
}
function freeze(obj, deep = false) {
  if (isFrozen(obj) || isDraft(obj) || !isDraftable(obj))
    return obj;
  if (getArchtype(obj) > 1) {
    Object.defineProperties(obj, {
      set: dontMutateMethodOverride,
      add: dontMutateMethodOverride,
      clear: dontMutateMethodOverride,
      delete: dontMutateMethodOverride
    });
  }
  Object.freeze(obj);
  if (deep)
    Object.values(obj).forEach((value) => freeze(value, true));
  return obj;
}
function dontMutateFrozenCollections() {
  die(2);
}
var dontMutateMethodOverride = {
  value: dontMutateFrozenCollections
};
function isFrozen(obj) {
  if (obj === null || typeof obj !== "object")
    return true;
  return Object.isFrozen(obj);
}
var plugins = {};
function getPlugin(pluginKey) {
  const plugin = plugins[pluginKey];
  if (!plugin) {
    die(0, pluginKey);
  }
  return plugin;
}
function loadPlugin(pluginKey, implementation) {
  if (!plugins[pluginKey])
    plugins[pluginKey] = implementation;
}
var currentScope;
function getCurrentScope() {
  return currentScope;
}
function createScope(parent_, immer_) {
  return {
    drafts_: [],
    parent_,
    immer_,
    // Whenever the modified draft contains a draft from another scope, we
    // need to prevent auto-freezing so the unowned draft can be finalized.
    canAutoFreeze_: true,
    unfinalizedDrafts_: 0
  };
}
function usePatchesInScope(scope, patchListener) {
  if (patchListener) {
    getPlugin("Patches");
    scope.patches_ = [];
    scope.inversePatches_ = [];
    scope.patchListener_ = patchListener;
  }
}
function revokeScope(scope) {
  leaveScope(scope);
  scope.drafts_.forEach(revokeDraft);
  scope.drafts_ = null;
}
function leaveScope(scope) {
  if (scope === currentScope) {
    currentScope = scope.parent_;
  }
}
function enterScope(immer2) {
  return currentScope = createScope(currentScope, immer2);
}
function revokeDraft(draft2) {
  const state = draft2[DRAFT_STATE];
  if (state.type_ === 0 || state.type_ === 1)
    state.revoke_();
  else
    state.revoked_ = true;
}
function processResult(result, scope) {
  scope.unfinalizedDrafts_ = scope.drafts_.length;
  const baseDraft = scope.drafts_[0];
  const isReplaced = result !== void 0 && result !== baseDraft;
  if (isReplaced) {
    if (baseDraft[DRAFT_STATE].modified_) {
      revokeScope(scope);
      die(4);
    }
    if (isDraftable(result)) {
      result = finalize(scope, result);
      if (!scope.parent_)
        maybeFreeze(scope, result);
    }
    if (scope.patches_) {
      getPlugin("Patches").generateReplacementPatches_(
        baseDraft[DRAFT_STATE].base_,
        result,
        scope.patches_,
        scope.inversePatches_
      );
    }
  } else {
    result = finalize(scope, baseDraft, []);
  }
  revokeScope(scope);
  if (scope.patches_) {
    scope.patchListener_(scope.patches_, scope.inversePatches_);
  }
  return result !== NOTHING ? result : void 0;
}
function finalize(rootScope, value, path) {
  if (isFrozen(value))
    return value;
  const useStrictIteration = rootScope.immer_.shouldUseStrictIteration();
  const state = value[DRAFT_STATE];
  if (!state) {
    each(
      value,
      (key, childValue) => finalizeProperty(rootScope, state, value, key, childValue, path),
      useStrictIteration
    );
    return value;
  }
  if (state.scope_ !== rootScope)
    return value;
  if (!state.modified_) {
    maybeFreeze(rootScope, state.base_, true);
    return state.base_;
  }
  if (!state.finalized_) {
    state.finalized_ = true;
    state.scope_.unfinalizedDrafts_--;
    const result = state.copy_;
    let resultEach = result;
    let isSet2 = false;
    if (state.type_ === 3) {
      resultEach = new Set(result);
      result.clear();
      isSet2 = true;
    }
    each(
      resultEach,
      (key, childValue) => finalizeProperty(
        rootScope,
        state,
        result,
        key,
        childValue,
        path,
        isSet2
      ),
      useStrictIteration
    );
    maybeFreeze(rootScope, result, false);
    if (path && rootScope.patches_) {
      getPlugin("Patches").generatePatches_(
        state,
        path,
        rootScope.patches_,
        rootScope.inversePatches_
      );
    }
  }
  return state.copy_;
}
function finalizeProperty(rootScope, parentState, targetObject, prop, childValue, rootPath, targetIsSet) {
  if (childValue == null) {
    return;
  }
  if (typeof childValue !== "object" && !targetIsSet) {
    return;
  }
  const childIsFrozen = isFrozen(childValue);
  if (childIsFrozen && !targetIsSet) {
    return;
  }
  if (childValue === targetObject)
    die(5);
  if (isDraft(childValue)) {
    const path = rootPath && parentState && parentState.type_ !== 3 && // Set objects are atomic since they have no keys.
    !has(parentState.assigned_, prop) ? rootPath.concat(prop) : void 0;
    const res = finalize(rootScope, childValue, path);
    set(targetObject, prop, res);
    if (isDraft(res)) {
      rootScope.canAutoFreeze_ = false;
    } else
      return;
  } else if (targetIsSet) {
    targetObject.add(childValue);
  }
  if (isDraftable(childValue) && !childIsFrozen) {
    if (!rootScope.immer_.autoFreeze_ && rootScope.unfinalizedDrafts_ < 1) {
      return;
    }
    if (parentState && parentState.base_ && parentState.base_[prop] === childValue && childIsFrozen) {
      return;
    }
    finalize(rootScope, childValue);
    if ((!parentState || !parentState.scope_.parent_) && typeof prop !== "symbol" && (isMap(targetObject) ? targetObject.has(prop) : Object.prototype.propertyIsEnumerable.call(targetObject, prop)))
      maybeFreeze(rootScope, childValue);
  }
}
function maybeFreeze(scope, value, deep = false) {
  if (!scope.parent_ && scope.immer_.autoFreeze_ && scope.canAutoFreeze_) {
    freeze(value, deep);
  }
}
function createProxyProxy(base, parent) {
  const isArray = Array.isArray(base);
  const state = {
    type_: isArray ? 1 : 0,
    // Track which produce call this is associated with.
    scope_: parent ? parent.scope_ : getCurrentScope(),
    // True for both shallow and deep changes.
    modified_: false,
    // Used during finalization.
    finalized_: false,
    // Track which properties have been assigned (true) or deleted (false).
    assigned_: {},
    // The parent draft state.
    parent_: parent,
    // The base state.
    base_: base,
    // The base proxy.
    draft_: null,
    // set below
    // The base copy with any updated values.
    copy_: null,
    // Called by the `produce` function.
    revoke_: null,
    isManual_: false
  };
  let target = state;
  let traps = objectTraps;
  if (isArray) {
    target = [state];
    traps = arrayTraps;
  }
  const { revoke, proxy } = Proxy.revocable(target, traps);
  state.draft_ = proxy;
  state.revoke_ = revoke;
  return proxy;
}
var objectTraps = {
  get(state, prop) {
    if (prop === DRAFT_STATE)
      return state;
    const source = latest(state);
    if (!has(source, prop)) {
      return readPropFromProto(state, source, prop);
    }
    const value = source[prop];
    if (state.finalized_ || !isDraftable(value)) {
      return value;
    }
    if (value === peek(state.base_, prop)) {
      prepareCopy(state);
      return state.copy_[prop] = createProxy(value, state);
    }
    return value;
  },
  has(state, prop) {
    return prop in latest(state);
  },
  ownKeys(state) {
    return Reflect.ownKeys(latest(state));
  },
  set(state, prop, value) {
    const desc = getDescriptorFromProto(latest(state), prop);
    if (desc?.set) {
      desc.set.call(state.draft_, value);
      return true;
    }
    if (!state.modified_) {
      const current2 = peek(latest(state), prop);
      const currentState = current2?.[DRAFT_STATE];
      if (currentState && currentState.base_ === value) {
        state.copy_[prop] = value;
        state.assigned_[prop] = false;
        return true;
      }
      if (is(value, current2) && (value !== void 0 || has(state.base_, prop)))
        return true;
      prepareCopy(state);
      markChanged(state);
    }
    if (state.copy_[prop] === value && // special case: handle new props with value 'undefined'
    (value !== void 0 || prop in state.copy_) || // special case: NaN
    Number.isNaN(value) && Number.isNaN(state.copy_[prop]))
      return true;
    state.copy_[prop] = value;
    state.assigned_[prop] = true;
    return true;
  },
  deleteProperty(state, prop) {
    if (peek(state.base_, prop) !== void 0 || prop in state.base_) {
      state.assigned_[prop] = false;
      prepareCopy(state);
      markChanged(state);
    } else {
      delete state.assigned_[prop];
    }
    if (state.copy_) {
      delete state.copy_[prop];
    }
    return true;
  },
  // Note: We never coerce `desc.value` into an Immer draft, because we can't make
  // the same guarantee in ES5 mode.
  getOwnPropertyDescriptor(state, prop) {
    const owner = latest(state);
    const desc = Reflect.getOwnPropertyDescriptor(owner, prop);
    if (!desc)
      return desc;
    return {
      writable: true,
      configurable: state.type_ !== 1 || prop !== "length",
      enumerable: desc.enumerable,
      value: owner[prop]
    };
  },
  defineProperty() {
    die(11);
  },
  getPrototypeOf(state) {
    return getPrototypeOf(state.base_);
  },
  setPrototypeOf() {
    die(12);
  }
};
var arrayTraps = {};
each(objectTraps, (key, fn) => {
  arrayTraps[key] = function() {
    arguments[0] = arguments[0][0];
    return fn.apply(this, arguments);
  };
});
arrayTraps.deleteProperty = function(state, prop) {
  if (isNaN(parseInt(prop)))
    die(13);
  return arrayTraps.set.call(this, state, prop, void 0);
};
arrayTraps.set = function(state, prop, value) {
  if (prop !== "length" && isNaN(parseInt(prop)))
    die(14);
  return objectTraps.set.call(this, state[0], prop, value, state[0]);
};
function peek(draft2, prop) {
  const state = draft2[DRAFT_STATE];
  const source = state ? latest(state) : draft2;
  return source[prop];
}
function readPropFromProto(state, source, prop) {
  const desc = getDescriptorFromProto(source, prop);
  return desc ? `value` in desc ? desc.value : (
    // This is a very special case, if the prop is a getter defined by the
    // prototype, we should invoke it with the draft as context!
    desc.get?.call(state.draft_)
  ) : void 0;
}
function getDescriptorFromProto(source, prop) {
  if (!(prop in source))
    return void 0;
  let proto = getPrototypeOf(source);
  while (proto) {
    const desc = Object.getOwnPropertyDescriptor(proto, prop);
    if (desc)
      return desc;
    proto = getPrototypeOf(proto);
  }
  return void 0;
}
function markChanged(state) {
  if (!state.modified_) {
    state.modified_ = true;
    if (state.parent_) {
      markChanged(state.parent_);
    }
  }
}
function prepareCopy(state) {
  if (!state.copy_) {
    state.copy_ = shallowCopy(
      state.base_,
      state.scope_.immer_.useStrictShallowCopy_
    );
  }
}
var Immer2 = class {
  constructor(config) {
    this.autoFreeze_ = true;
    this.useStrictShallowCopy_ = false;
    this.useStrictIteration_ = true;
    this.produce = (base, recipe, patchListener) => {
      if (typeof base === "function" && typeof recipe !== "function") {
        const defaultBase = recipe;
        recipe = base;
        const self = this;
        return function curriedProduce(base2 = defaultBase, ...args) {
          return self.produce(base2, (draft2) => recipe.call(this, draft2, ...args));
        };
      }
      if (typeof recipe !== "function")
        die(6);
      if (patchListener !== void 0 && typeof patchListener !== "function")
        die(7);
      let result;
      if (isDraftable(base)) {
        const scope = enterScope(this);
        const proxy = createProxy(base, void 0);
        let hasError = true;
        try {
          result = recipe(proxy);
          hasError = false;
        } finally {
          if (hasError)
            revokeScope(scope);
          else
            leaveScope(scope);
        }
        usePatchesInScope(scope, patchListener);
        return processResult(result, scope);
      } else if (!base || typeof base !== "object") {
        result = recipe(base);
        if (result === void 0)
          result = base;
        if (result === NOTHING)
          result = void 0;
        if (this.autoFreeze_)
          freeze(result, true);
        if (patchListener) {
          const p = [];
          const ip = [];
          getPlugin("Patches").generateReplacementPatches_(base, result, p, ip);
          patchListener(p, ip);
        }
        return result;
      } else
        die(1, base);
    };
    this.produceWithPatches = (base, recipe) => {
      if (typeof base === "function") {
        return (state, ...args) => this.produceWithPatches(state, (draft2) => base(draft2, ...args));
      }
      let patches, inversePatches;
      const result = this.produce(base, recipe, (p, ip) => {
        patches = p;
        inversePatches = ip;
      });
      return [result, patches, inversePatches];
    };
    if (typeof config?.autoFreeze === "boolean")
      this.setAutoFreeze(config.autoFreeze);
    if (typeof config?.useStrictShallowCopy === "boolean")
      this.setUseStrictShallowCopy(config.useStrictShallowCopy);
    if (typeof config?.useStrictIteration === "boolean")
      this.setUseStrictIteration(config.useStrictIteration);
  }
  createDraft(base) {
    if (!isDraftable(base))
      die(8);
    if (isDraft(base))
      base = current(base);
    const scope = enterScope(this);
    const proxy = createProxy(base, void 0);
    proxy[DRAFT_STATE].isManual_ = true;
    leaveScope(scope);
    return proxy;
  }
  finishDraft(draft2, patchListener) {
    const state = draft2 && draft2[DRAFT_STATE];
    if (!state || !state.isManual_)
      die(9);
    const { scope_: scope } = state;
    usePatchesInScope(scope, patchListener);
    return processResult(void 0, scope);
  }
  /**
   * Pass true to automatically freeze all copies created by Immer.
   *
   * By default, auto-freezing is enabled.
   */
  setAutoFreeze(value) {
    this.autoFreeze_ = value;
  }
  /**
   * Pass true to enable strict shallow copy.
   *
   * By default, immer does not copy the object descriptors such as getter, setter and non-enumrable properties.
   */
  setUseStrictShallowCopy(value) {
    this.useStrictShallowCopy_ = value;
  }
  /**
   * Pass false to use faster iteration that skips non-enumerable properties
   * but still handles symbols for compatibility.
   *
   * By default, strict iteration is enabled (includes all own properties).
   */
  setUseStrictIteration(value) {
    this.useStrictIteration_ = value;
  }
  shouldUseStrictIteration() {
    return this.useStrictIteration_;
  }
  applyPatches(base, patches) {
    let i;
    for (i = patches.length - 1; i >= 0; i--) {
      const patch = patches[i];
      if (patch.path.length === 0 && patch.op === "replace") {
        base = patch.value;
        break;
      }
    }
    if (i > -1) {
      patches = patches.slice(i + 1);
    }
    const applyPatchesImpl = getPlugin("Patches").applyPatches_;
    if (isDraft(base)) {
      return applyPatchesImpl(base, patches);
    }
    return this.produce(
      base,
      (draft2) => applyPatchesImpl(draft2, patches)
    );
  }
};
function createProxy(value, parent) {
  const draft2 = isMap(value) ? getPlugin("MapSet").proxyMap_(value, parent) : isSet(value) ? getPlugin("MapSet").proxySet_(value, parent) : createProxyProxy(value, parent);
  const scope = parent ? parent.scope_ : getCurrentScope();
  scope.drafts_.push(draft2);
  return draft2;
}
function current(value) {
  if (!isDraft(value))
    die(10, value);
  return currentImpl(value);
}
function currentImpl(value) {
  if (!isDraftable(value) || isFrozen(value))
    return value;
  const state = value[DRAFT_STATE];
  let copy;
  let strict = true;
  if (state) {
    if (!state.modified_)
      return state.base_;
    state.finalized_ = true;
    copy = shallowCopy(value, state.scope_.immer_.useStrictShallowCopy_);
    strict = state.scope_.immer_.shouldUseStrictIteration();
  } else {
    copy = shallowCopy(value, true);
  }
  each(
    copy,
    (key, childValue) => {
      set(copy, key, currentImpl(childValue));
    },
    strict
  );
  if (state) {
    state.finalized_ = false;
  }
  return copy;
}
function enablePatches() {
  const errorOffset = 16;
  if (true) {
    errors.push(
      'Sets cannot have "replace" patches.',
      function(op) {
        return "Unsupported patch operation: " + op;
      },
      function(path) {
        return "Cannot apply patch, path doesn't resolve: " + path;
      },
      "Patching reserved attributes like __proto__, prototype and constructor is not allowed"
    );
  }
  const REPLACE = "replace";
  const ADD = "add";
  const REMOVE = "remove";
  function generatePatches_(state, basePath, patches, inversePatches) {
    switch (state.type_) {
      case 0:
      case 2:
        return generatePatchesFromAssigned(
          state,
          basePath,
          patches,
          inversePatches
        );
      case 1:
        return generateArrayPatches(state, basePath, patches, inversePatches);
      case 3:
        return generateSetPatches(
          state,
          basePath,
          patches,
          inversePatches
        );
    }
  }
  function generateArrayPatches(state, basePath, patches, inversePatches) {
    let { base_, assigned_ } = state;
    let copy_ = state.copy_;
    if (copy_.length < base_.length) {
      ;
      [base_, copy_] = [copy_, base_];
      [patches, inversePatches] = [inversePatches, patches];
    }
    for (let i = 0; i < base_.length; i++) {
      if (assigned_[i] && copy_[i] !== base_[i]) {
        const path = basePath.concat([i]);
        patches.push({
          op: REPLACE,
          path,
          // Need to maybe clone it, as it can in fact be the original value
          // due to the base/copy inversion at the start of this function
          value: clonePatchValueIfNeeded(copy_[i])
        });
        inversePatches.push({
          op: REPLACE,
          path,
          value: clonePatchValueIfNeeded(base_[i])
        });
      }
    }
    for (let i = base_.length; i < copy_.length; i++) {
      const path = basePath.concat([i]);
      patches.push({
        op: ADD,
        path,
        // Need to maybe clone it, as it can in fact be the original value
        // due to the base/copy inversion at the start of this function
        value: clonePatchValueIfNeeded(copy_[i])
      });
    }
    for (let i = copy_.length - 1; base_.length <= i; --i) {
      const path = basePath.concat([i]);
      inversePatches.push({
        op: REMOVE,
        path
      });
    }
  }
  function generatePatchesFromAssigned(state, basePath, patches, inversePatches) {
    const { base_, copy_ } = state;
    each(state.assigned_, (key, assignedValue) => {
      const origValue = get(base_, key);
      const value = get(copy_, key);
      const op = !assignedValue ? REMOVE : has(base_, key) ? REPLACE : ADD;
      if (origValue === value && op === REPLACE)
        return;
      const path = basePath.concat(key);
      patches.push(op === REMOVE ? { op, path } : { op, path, value });
      inversePatches.push(
        op === ADD ? { op: REMOVE, path } : op === REMOVE ? { op: ADD, path, value: clonePatchValueIfNeeded(origValue) } : { op: REPLACE, path, value: clonePatchValueIfNeeded(origValue) }
      );
    });
  }
  function generateSetPatches(state, basePath, patches, inversePatches) {
    let { base_, copy_ } = state;
    let i = 0;
    base_.forEach((value) => {
      if (!copy_.has(value)) {
        const path = basePath.concat([i]);
        patches.push({
          op: REMOVE,
          path,
          value
        });
        inversePatches.unshift({
          op: ADD,
          path,
          value
        });
      }
      i++;
    });
    i = 0;
    copy_.forEach((value) => {
      if (!base_.has(value)) {
        const path = basePath.concat([i]);
        patches.push({
          op: ADD,
          path,
          value
        });
        inversePatches.unshift({
          op: REMOVE,
          path,
          value
        });
      }
      i++;
    });
  }
  function generateReplacementPatches_(baseValue, replacement, patches, inversePatches) {
    patches.push({
      op: REPLACE,
      path: [],
      value: replacement === NOTHING ? void 0 : replacement
    });
    inversePatches.push({
      op: REPLACE,
      path: [],
      value: baseValue
    });
  }
  function applyPatches_(draft2, patches) {
    patches.forEach((patch) => {
      const { path, op } = patch;
      let base = draft2;
      for (let i = 0; i < path.length - 1; i++) {
        const parentType = getArchtype(base);
        let p = path[i];
        if (typeof p !== "string" && typeof p !== "number") {
          p = "" + p;
        }
        if ((parentType === 0 || parentType === 1) && (p === "__proto__" || p === "constructor"))
          die(errorOffset + 3);
        if (typeof base === "function" && p === "prototype")
          die(errorOffset + 3);
        base = get(base, p);
        if (typeof base !== "object")
          die(errorOffset + 2, path.join("/"));
      }
      const type = getArchtype(base);
      const value = deepClonePatchValue(patch.value);
      const key = path[path.length - 1];
      switch (op) {
        case REPLACE:
          switch (type) {
            case 2:
              return base.set(key, value);
            case 3:
              die(errorOffset);
            default:
              return base[key] = value;
          }
        case ADD:
          switch (type) {
            case 1:
              return key === "-" ? base.push(value) : base.splice(key, 0, value);
            case 2:
              return base.set(key, value);
            case 3:
              return base.add(value);
            default:
              return base[key] = value;
          }
        case REMOVE:
          switch (type) {
            case 1:
              return base.splice(key, 1);
            case 2:
              return base.delete(key);
            case 3:
              return base.delete(patch.value);
            default:
              return delete base[key];
          }
        default:
          die(errorOffset + 1, op);
      }
    });
    return draft2;
  }
  function deepClonePatchValue(obj) {
    if (!isDraftable(obj))
      return obj;
    if (Array.isArray(obj))
      return obj.map(deepClonePatchValue);
    if (isMap(obj))
      return new Map(
        Array.from(obj.entries()).map(([k, v]) => [k, deepClonePatchValue(v)])
      );
    if (isSet(obj))
      return new Set(Array.from(obj).map(deepClonePatchValue));
    const cloned = Object.create(getPrototypeOf(obj));
    for (const key in obj)
      cloned[key] = deepClonePatchValue(obj[key]);
    if (has(obj, DRAFTABLE))
      cloned[DRAFTABLE] = obj[DRAFTABLE];
    return cloned;
  }
  function clonePatchValueIfNeeded(obj) {
    if (isDraft(obj)) {
      return deepClonePatchValue(obj);
    } else
      return obj;
  }
  loadPlugin("Patches", {
    applyPatches_,
    generatePatches_,
    generateReplacementPatches_
  });
}
var immer = new Immer2();
var produce = immer.produce;
var produceWithPatches = /* @__PURE__ */ immer.produceWithPatches.bind(
  immer
);

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneItemsApi.js
var __awaiter7 = function(thisArg, _arguments, P, generator) {
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
enablePatches();
var SceneItemsApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getItems(filter) {
    return __awaiter7(this, void 0, void 0, function* () {
      if (Array.isArray(filter)) {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_GET_ITEMS", { ids: filter });
        return items;
      } else if (filter) {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_GET_ALL_ITEMS", {});
        return items.filter(filter);
      } else {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_GET_ALL_ITEMS", {});
        return items;
      }
    });
  }
  isItemArray(value) {
    return Array.isArray(value) && value.every((item) => typeof item !== "string");
  }
  updateItems(filterOrItems, update) {
    return __awaiter7(this, void 0, void 0, function* () {
      let items;
      if (this.isItemArray(filterOrItems)) {
        items = filterOrItems;
      } else {
        items = yield this.getItems(filterOrItems);
      }
      const [nextState, patches] = produceWithPatches(items, update);
      const nextUpdates = nextState.map((item) => ({
        id: item.id,
        type: item.type
      }));
      for (const patch of patches) {
        const [index, key] = patch.path;
        if (typeof index === "number" && typeof key === "string") {
          nextUpdates[index][key] = nextState[index][key];
        }
      }
      const updates = nextUpdates.filter(
        // Ensure that there are updates besides the default ID and type
        (update2) => Object.keys(update2).length > 2
      );
      if (updates.length === 0) {
        return;
      }
      yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_UPDATE_ITEMS", {
        updates
      });
    });
  }
  addItems(items) {
    return __awaiter7(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_ADD_ITEMS", {
        items
      });
    });
  }
  deleteItems(ids) {
    return __awaiter7(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_DELETE_ITEMS", {
        ids
      });
    });
  }
  getItemAttachments(ids) {
    return __awaiter7(this, void 0, void 0, function* () {
      const { items } = yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_GET_ITEM_ATTACHMENTS", { ids });
      return items;
    });
  }
  getItemBounds(ids) {
    return __awaiter7(this, void 0, void 0, function* () {
      const { bounds } = yield this.messageBus.sendAsync("OBR_SCENE_ITEMS_GET_ITEM_BOUNDS", { ids });
      return bounds;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.items);
    };
    this.messageBus.send("OBR_SCENE_ITEMS_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_ITEMS_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_ITEMS_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_ITEMS_EVENT_CHANGE", handleChange);
    };
  }
};
var SceneItemsApi_default = SceneItemsApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneLocalApi.js
var __awaiter8 = function(thisArg, _arguments, P, generator) {
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
enablePatches();
var SceneLocalApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getItems(filter) {
    return __awaiter8(this, void 0, void 0, function* () {
      if (Array.isArray(filter)) {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_GET_ITEMS", { ids: filter });
        return items;
      } else if (filter) {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_GET_ALL_ITEMS", {});
        return items.filter(filter);
      } else {
        const { items } = yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_GET_ALL_ITEMS", {});
        return items;
      }
    });
  }
  isItemArray(value) {
    return Array.isArray(value) && value.every((item) => typeof item !== "string");
  }
  updateItems(filterOrItems, update, fastUpdate) {
    return __awaiter8(this, void 0, void 0, function* () {
      let items;
      if (this.isItemArray(filterOrItems)) {
        items = filterOrItems;
      } else {
        items = yield this.getItems(filterOrItems);
      }
      const [nextState, patches] = produceWithPatches(items, update);
      const nextUpdates = nextState.map((item) => ({
        id: item.id,
        type: item.type
      }));
      for (const patch of patches) {
        const [index, key] = patch.path;
        if (typeof index === "number" && typeof key === "string") {
          nextUpdates[index][key] = nextState[index][key];
        }
      }
      const updates = nextUpdates.filter(
        // Ensure that there are updates besides the default ID and type
        (update2) => Object.keys(update2).length > 2
      );
      if (updates.length === 0) {
        return;
      }
      yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_UPDATE_ITEMS", {
        updates,
        fastUpdate
      });
    });
  }
  addItems(items) {
    return __awaiter8(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_ADD_ITEMS", {
        items
      });
    });
  }
  deleteItems(ids) {
    return __awaiter8(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_DELETE_ITEMS", {
        ids
      });
    });
  }
  getItemAttachments(ids) {
    return __awaiter8(this, void 0, void 0, function* () {
      const { items } = yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_GET_ITEM_ATTACHMENTS", { ids });
      return items;
    });
  }
  getItemBounds(ids) {
    return __awaiter8(this, void 0, void 0, function* () {
      const { bounds } = yield this.messageBus.sendAsync("OBR_SCENE_LOCAL_GET_ITEM_BOUNDS", { ids });
      return bounds;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.items);
    };
    this.messageBus.send("OBR_SCENE_LOCAL_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_LOCAL_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_LOCAL_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_LOCAL_EVENT_CHANGE", handleChange);
    };
  }
};
var SceneLocalApi_default = SceneLocalApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/scene/SceneApi.js
var __awaiter9 = function(thisArg, _arguments, P, generator) {
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
var SceneApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
    this.grid = new SceneGridApi_default(messageBus2);
    this.fog = new SceneFogApi_default(messageBus2);
    this.history = new SceneHistoryApi_default(messageBus2);
    this.items = new SceneItemsApi_default(messageBus2);
    this.local = new SceneLocalApi_default(messageBus2);
  }
  isReady() {
    return __awaiter9(this, void 0, void 0, function* () {
      const { ready } = yield this.messageBus.sendAsync("OBR_SCENE_IS_READY", {});
      return ready;
    });
  }
  onReadyChange(callback) {
    const handleChange = (data) => {
      callback(data.ready);
    };
    this.messageBus.send("OBR_SCENE_READY_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_EVENT_READY_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_READY_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_EVENT_READY_CHANGE", handleChange);
    };
  }
  getMetadata() {
    return __awaiter9(this, void 0, void 0, function* () {
      const { metadata } = yield this.messageBus.sendAsync("OBR_SCENE_GET_METADATA", {});
      return metadata;
    });
  }
  setMetadata(update) {
    return __awaiter9(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_SCENE_SET_METADATA", { update });
    });
  }
  onMetadataChange(callback) {
    const handleChange = (data) => {
      callback(data.metadata);
    };
    this.messageBus.send("OBR_SCENE_METADATA_SUBSCRIBE", {});
    this.messageBus.on("OBR_SCENE_METADATA_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_SCENE_METADATA_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_SCENE_METADATA_EVENT_CHANGE", handleChange);
    };
  }
};
var SceneApi_default = SceneApi;

// node_modules/@owlbear-rodeo/sdk/lib/common/normalize.js
function normalizeUrl(url) {
  return url.startsWith("http") ? url : `${window.location.origin}${url}`;
}
function normalizeIconPaths(icons) {
  return icons.map((base) => Object.assign(Object.assign({}, base), { icon: normalizeUrl(base.icon) }));
}
function normalizeUrlObject(urlObject) {
  return Object.assign(Object.assign({}, urlObject), { url: normalizeUrl(urlObject.url) });
}

// node_modules/@owlbear-rodeo/sdk/lib/api/ContextMenuApi.js
var __awaiter10 = function(thisArg, _arguments, P, generator) {
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
var ContextMenuApi = class {
  constructor(messageBus2) {
    this.contextMenus = {};
    this.handleClick = (event) => {
      var _a;
      const menu = this.contextMenus[event.id];
      if (menu) {
        (_a = menu.onClick) === null || _a === void 0 ? void 0 : _a.call(menu, event.context, event.elementId);
      }
    };
    this.messageBus = messageBus2;
    messageBus2.on("OBR_CONTEXT_MENU_EVENT_CLICK", this.handleClick);
  }
  create(contextMenu) {
    return __awaiter10(this, void 0, void 0, function* () {
      this.messageBus.sendAsync("OBR_CONTEXT_MENU_CREATE", {
        id: contextMenu.id,
        shortcut: contextMenu.shortcut,
        icons: normalizeIconPaths(contextMenu.icons),
        embed: contextMenu.embed && normalizeUrlObject(contextMenu.embed)
      });
      this.contextMenus[contextMenu.id] = contextMenu;
    });
  }
  remove(id) {
    return __awaiter10(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_CONTEXT_MENU_REMOVE", { id });
      delete this.contextMenus[id];
    });
  }
};
var ContextMenuApi_default = ContextMenuApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/ToolApi.js
var __awaiter11 = function(thisArg, _arguments, P, generator) {
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
var ToolApi = class {
  constructor(messageBus2) {
    this.tools = {};
    this.toolActions = {};
    this.toolModes = {};
    this.handleToolClick = (event) => {
      const tool = this.tools[event.id];
      if (tool) {
        if (tool.onClick) {
          const result = tool.onClick(event.context, event.elementId);
          Promise.resolve(result).then((activate) => {
            if (activate) {
              this.messageBus.send("OBR_TOOL_ACTIVATE", {
                id: event.id
              });
            }
          });
        } else {
          this.messageBus.send("OBR_TOOL_ACTIVATE", {
            id: event.id
          });
        }
      }
    };
    this.handleToolActionClick = (event) => {
      var _a;
      const action = this.toolActions[event.id];
      if (action) {
        (_a = action.onClick) === null || _a === void 0 ? void 0 : _a.call(action, event.context, event.elementId);
      }
    };
    this.handleToolModeClick = (event) => {
      const mode = this.toolModes[event.id];
      if (mode) {
        if (mode.onClick) {
          const result = mode.onClick(event.context, event.elementId);
          Promise.resolve(result).then((activate) => {
            if (activate) {
              this.messageBus.send("OBR_TOOL_MODE_ACTIVATE", {
                toolId: event.context.activeTool,
                modeId: event.id
              });
            }
          });
        } else {
          this.messageBus.send("OBR_TOOL_MODE_ACTIVATE", {
            toolId: event.context.activeTool,
            modeId: event.id
          });
        }
      }
    };
    this.handleToolModeToolClick = (event) => {
      const mode = this.toolModes[event.id];
      if (mode) {
        if (mode.onToolClick) {
          const result = mode.onToolClick(event.context, event.event);
          Promise.resolve(result).then((select) => {
            if (select && event.event.target && !event.event.target.locked) {
              this.messageBus.sendAsync("OBR_PLAYER_SELECT", {
                items: [event.event.target.id]
              });
            }
          });
        } else {
          if (event.event.target && !event.event.target.locked) {
            this.messageBus.sendAsync("OBR_PLAYER_SELECT", {
              items: [event.event.target.id]
            });
          }
        }
      }
    };
    this.handleToolModeToolDoubleClick = (event) => {
      const mode = this.toolModes[event.id];
      if (mode) {
        if (mode.onToolDoubleClick) {
          const result = mode.onToolDoubleClick(event.context, event.event);
          Promise.resolve(result).then((select) => {
            if (select && event.event.target) {
              this.messageBus.sendAsync("OBR_PLAYER_SELECT", {
                items: [event.event.target.id]
              });
            }
          });
        } else {
          if (event.event.target) {
            this.messageBus.sendAsync("OBR_PLAYER_SELECT", {
              items: [event.event.target.id]
            });
          }
        }
      }
    };
    this.handleToolModeToolDown = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolDown) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolMove = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolMove) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolUp = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolUp) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolDragStart = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolDragStart) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolDragMove = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolDragMove) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolDragEnd = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolDragEnd) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeToolDragCancel = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onToolDragCancel) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeKeyDown = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onKeyDown) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeKeyUp = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onKeyUp) === null || _a === void 0 ? void 0 : _a.call(mode, event.context, event.event);
      }
    };
    this.handleToolModeActivate = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onActivate) === null || _a === void 0 ? void 0 : _a.call(mode, event.context);
      }
    };
    this.handleToolModeDeactivate = (event) => {
      var _a;
      const mode = this.toolModes[event.id];
      if (mode) {
        (_a = mode.onDeactivate) === null || _a === void 0 ? void 0 : _a.call(mode, event.context);
      }
    };
    this.messageBus = messageBus2;
    messageBus2.on("OBR_TOOL_EVENT_CLICK", this.handleToolClick);
    messageBus2.on("OBR_TOOL_ACTION_EVENT_CLICK", this.handleToolActionClick);
    messageBus2.on("OBR_TOOL_MODE_EVENT_CLICK", this.handleToolModeClick);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_CLICK", this.handleToolModeToolClick);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DOUBLE_CLICK", this.handleToolModeToolDoubleClick);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DOWN", this.handleToolModeToolDown);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_MOVE", this.handleToolModeToolMove);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_UP", this.handleToolModeToolUp);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DRAG_START", this.handleToolModeToolDragStart);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DRAG_MOVE", this.handleToolModeToolDragMove);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DRAG_END", this.handleToolModeToolDragEnd);
    messageBus2.on("OBR_TOOL_MODE_EVENT_TOOL_DRAG_CANCEL", this.handleToolModeToolDragCancel);
    messageBus2.on("OBR_TOOL_MODE_EVENT_KEY_DOWN", this.handleToolModeKeyDown);
    messageBus2.on("OBR_TOOL_MODE_EVENT_KEY_UP", this.handleToolModeKeyUp);
    messageBus2.on("OBR_TOOL_MODE_EVENT_ACTIVATE", this.handleToolModeActivate);
    messageBus2.on("OBR_TOOL_MODE_EVENT_DEACTIVATE", this.handleToolModeDeactivate);
  }
  create(tool) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_CREATE", {
        id: tool.id,
        shortcut: tool.shortcut,
        defaultMode: tool.defaultMode,
        defaultMetadata: tool.defaultMetadata,
        icons: normalizeIconPaths(tool.icons),
        disabled: tool.disabled
      });
      this.tools[tool.id] = tool;
    });
  }
  remove(id) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_REMOVE", { id });
      delete this.tools[id];
    });
  }
  activateTool(id) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_ACTIVATE", { id });
    });
  }
  getActiveTool() {
    return __awaiter11(this, void 0, void 0, function* () {
      const { id } = yield this.messageBus.sendAsync("OBR_TOOL_GET_ACTIVE", {});
      return id;
    });
  }
  onToolChange(callback) {
    const handleChange = (data) => {
      callback(data.id);
    };
    this.messageBus.send("OBR_TOOL_ACTIVE_SUBSCRIBE", {});
    this.messageBus.on("OBR_TOOL_ACTIVE_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_TOOL_ACTIVE_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_TOOL_ACTIVE_EVENT_CHANGE", handleChange);
    };
  }
  getMetadata(id) {
    return __awaiter11(this, void 0, void 0, function* () {
      const { metadata } = yield this.messageBus.sendAsync("OBR_TOOL_GET_METADATA", { id });
      return metadata;
    });
  }
  setMetadata(toolId, update) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_SET_METADATA", {
        toolId,
        update
      });
    });
  }
  createAction(action) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_ACTION_CREATE", {
        id: action.id,
        shortcut: action.shortcut,
        icons: normalizeIconPaths(action.icons),
        disabled: action.disabled
      });
      this.toolActions[action.id] = action;
    });
  }
  removeAction(id) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_ACTION_REMOVE", { id });
      delete this.tools[id];
    });
  }
  createMode(mode) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_MODE_CREATE", {
        id: mode.id,
        shortcut: mode.shortcut,
        icons: normalizeIconPaths(mode.icons),
        preventDrag: mode.preventDrag,
        disabled: mode.disabled,
        cursors: mode.cursors
      });
      this.toolModes[mode.id] = mode;
    });
  }
  removeMode(id) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_MODE_REMOVE", { id });
      delete this.tools[id];
    });
  }
  activateMode(toolId, modeId) {
    return __awaiter11(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_TOOL_MODE_ACTIVATE", {
        toolId,
        modeId
      });
    });
  }
  getActiveToolMode() {
    return __awaiter11(this, void 0, void 0, function* () {
      const { id } = yield this.messageBus.sendAsync("OBR_TOOL_MODE_GET_ACTIVE", {});
      return id;
    });
  }
  onToolModeChange(callback) {
    const handleChange = (data) => {
      callback(data.id);
    };
    this.messageBus.send("OBR_TOOL_MODE_ACTIVE_SUBSCRIBE", {});
    this.messageBus.on("OBR_TOOL_MODE_ACTIVE_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_TOOL_MODE_ACTIVE_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_TOOL_MODE_ACTIVE_EVENT_CHANGE", handleChange);
    };
  }
};
var ToolApi_default = ToolApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/PopoverApi.js
var __awaiter12 = function(thisArg, _arguments, P, generator) {
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
var PopoverApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  open(popover) {
    return __awaiter12(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_POPOVER_OPEN", Object.assign({}, normalizeUrlObject(popover)));
    });
  }
  close(id) {
    return __awaiter12(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_POPOVER_CLOSE", { id });
    });
  }
  getWidth(id) {
    return __awaiter12(this, void 0, void 0, function* () {
      const { width } = yield this.messageBus.sendAsync("OBR_POPOVER_GET_WIDTH", { id });
      return width;
    });
  }
  setWidth(id, width) {
    return __awaiter12(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_POPOVER_SET_WIDTH", { id, width });
    });
  }
  getHeight(id) {
    return __awaiter12(this, void 0, void 0, function* () {
      const { height } = yield this.messageBus.sendAsync("OBR_POPOVER_GET_HEIGHT", { id });
      return height;
    });
  }
  setHeight(id, height) {
    return __awaiter12(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_POPOVER_SET_HEIGHT", { id, height });
    });
  }
};
var PopoverApi_default = PopoverApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/ModalApi.js
var __awaiter13 = function(thisArg, _arguments, P, generator) {
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
var ModalApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  open(modal) {
    return __awaiter13(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_MODAL_OPEN", Object.assign({}, normalizeUrlObject(modal)));
    });
  }
  close(id) {
    return __awaiter13(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_MODAL_CLOSE", { id });
    });
  }
};
var ModalApi_default = ModalApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/ActionApi.js
var __awaiter14 = function(thisArg, _arguments, P, generator) {
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
var ActionApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getWidth() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { width } = yield this.messageBus.sendAsync("OBR_ACTION_GET_WIDTH", {});
      return width;
    });
  }
  setWidth(width) {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_SET_WIDTH", { width });
    });
  }
  getHeight() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { height } = yield this.messageBus.sendAsync("OBR_ACTION_GET_HEIGHT", {});
      return height;
    });
  }
  setHeight(height) {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_SET_HEIGHT", { height });
    });
  }
  getBadgeText() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { badgeText } = yield this.messageBus.sendAsync("OBR_ACTION_GET_BADGE_TEXT", {});
      return badgeText;
    });
  }
  setBadgeText(badgeText) {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_SET_BADGE_TEXT", { badgeText });
    });
  }
  getBadgeBackgroundColor() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { badgeBackgroundColor } = yield this.messageBus.sendAsync("OBR_ACTION_GET_BADGE_BACKGROUND_COLOR", {});
      return badgeBackgroundColor;
    });
  }
  setBadgeBackgroundColor(badgeBackgroundColor) {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_SET_BADGE_BACKGROUND_COLOR", {
        badgeBackgroundColor
      });
    });
  }
  getIcon() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { icon } = yield this.messageBus.sendAsync("OBR_ACTION_GET_ICON", {});
      return icon;
    });
  }
  setIcon(icon) {
    return __awaiter14(this, void 0, void 0, function* () {
      const data = normalizeIconPaths([{ icon }]);
      yield this.messageBus.sendAsync("OBR_ACTION_SET_ICON", {
        icon: data[0].icon
      });
    });
  }
  getTitle() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { title } = yield this.messageBus.sendAsync("OBR_ACTION_GET_TITLE", {});
      return title;
    });
  }
  setTitle(title) {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_SET_TITLE", { title });
    });
  }
  isOpen() {
    return __awaiter14(this, void 0, void 0, function* () {
      const { isOpen } = yield this.messageBus.sendAsync("OBR_ACTION_GET_IS_OPEN", {});
      return isOpen;
    });
  }
  open() {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_OPEN", {});
    });
  }
  close() {
    return __awaiter14(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ACTION_CLOSE", {});
    });
  }
  onOpenChange(callback) {
    const handleChange = (data) => {
      callback(data.isOpen);
    };
    this.messageBus.send("OBR_ACTION_IS_OPEN_SUBSCRIBE", {});
    this.messageBus.on("OBR_ACTION_IS_OPEN_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_IS_OPEN_ACTION_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_ACTION_IS_OPEN_EVENT_CHANGE", handleChange);
    };
  }
};
var ActionApi_default = ActionApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/InteractionApi.js
var __awaiter15 = function(thisArg, _arguments, P, generator) {
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
enablePatches();
var InteractionApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  startItemInteraction(baseState) {
    return __awaiter15(this, void 0, void 0, function* () {
      const { id } = yield this.messageBus.sendAsync("OBR_INTERACTION_START_ITEM_INTERACTION", { baseState });
      let prev = baseState;
      const dispatcher = (update) => {
        const [next, patches] = produceWithPatches(prev, update);
        prev = next;
        this.messageBus.send("OBR_INTERACTION_UPDATE_ITEM_INTERACTION", {
          id,
          patches
        });
        return next;
      };
      const stop = () => {
        this.messageBus.send("OBR_INTERACTION_STOP_ITEM_INTERACTION", { id });
      };
      return [dispatcher, stop];
    });
  }
};
var InteractionApi_default = InteractionApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/PartyApi.js
var __awaiter16 = function(thisArg, _arguments, P, generator) {
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
var PartyApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getPlayers() {
    return __awaiter16(this, void 0, void 0, function* () {
      const { players } = yield this.messageBus.sendAsync("OBR_PARTY_GET_PLAYERS", {});
      return players;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.players);
    };
    this.messageBus.send("OBR_PARTY_SUBSCRIBE", {});
    this.messageBus.on("OBR_PARTY_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_PARTY_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_PARTY_EVENT_CHANGE", handleChange);
    };
  }
};
var PartyApi_default = PartyApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/RoomApi.js
var __awaiter17 = function(thisArg, _arguments, P, generator) {
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
var RoomApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  get id() {
    return this.messageBus.roomId;
  }
  getPermissions() {
    return __awaiter17(this, void 0, void 0, function* () {
      const { permissions } = yield this.messageBus.sendAsync("OBR_ROOM_GET_PERMISSIONS", {});
      return permissions;
    });
  }
  getMetadata() {
    return __awaiter17(this, void 0, void 0, function* () {
      const { metadata } = yield this.messageBus.sendAsync("OBR_ROOM_GET_METADATA", {});
      return metadata;
    });
  }
  setMetadata(update) {
    return __awaiter17(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ROOM_SET_METADATA", { update });
    });
  }
  onMetadataChange(callback) {
    const handleChange = (data) => {
      callback(data.metadata);
    };
    this.messageBus.send("OBR_ROOM_METADATA_SUBSCRIBE", {});
    this.messageBus.on("OBR_ROOM_METADATA_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_METADATA_ROOM_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_ROOM_METADATA_EVENT_CHANGE", handleChange);
    };
  }
  onPermissionsChange(callback) {
    const handleChange = (data) => {
      callback(data.permissions);
    };
    this.messageBus.send("OBR_ROOM_PERMISSIONS_SUBSCRIBE", {});
    this.messageBus.on("OBR_ROOM_PERMISSIONS_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_PERMISSIONS_ROOM_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_ROOM_PERMISSIONS_EVENT_CHANGE", handleChange);
    };
  }
};
var RoomApi_default = RoomApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/ThemeApi.js
var __awaiter18 = function(thisArg, _arguments, P, generator) {
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
var ThemeApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  getTheme() {
    return __awaiter18(this, void 0, void 0, function* () {
      const { theme } = yield this.messageBus.sendAsync("OBR_THEME_GET_THEME", {});
      return theme;
    });
  }
  onChange(callback) {
    const handleChange = (data) => {
      callback(data.theme);
    };
    this.messageBus.send("OBR_THEME_SUBSCRIBE", {});
    this.messageBus.on("OBR_THEME_EVENT_CHANGE", handleChange);
    return () => {
      this.messageBus.send("OBR_THEME_UNSUBSCRIBE", {});
      this.messageBus.off("OBR_THEME_EVENT_CHANGE", handleChange);
    };
  }
};
var ThemeApi_default = ThemeApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/AssetsApi.js
var __awaiter19 = function(thisArg, _arguments, P, generator) {
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
var AssetsApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  uploadImages(images, typeHint) {
    return __awaiter19(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ASSETS_UPLOAD_IMAGES", {
        images,
        typeHint
      });
    });
  }
  uploadScenes(scenes, disableShowScenes) {
    return __awaiter19(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_ASSETS_UPLOAD_SCENES", {
        scenes,
        disableShowScenes
      });
    });
  }
  downloadImages(multiple, defaultSearch, typeHint) {
    return __awaiter19(this, void 0, void 0, function* () {
      const { images } = yield this.messageBus.sendAsync("OBR_ASSETS_DOWNLOAD_IMAGES", { multiple, defaultSearch, typeHint }, -1);
      return images;
    });
  }
  downloadScenes(multiple, defaultSearch) {
    return __awaiter19(this, void 0, void 0, function* () {
      const { scenes } = yield this.messageBus.sendAsync("OBR_ASSETS_DOWNLOAD_SCENES", { multiple, defaultSearch }, -1);
      return scenes;
    });
  }
};
var AssetsApi_default = AssetsApi;

// node_modules/@owlbear-rodeo/sdk/lib/api/BroadcastApi.js
var __awaiter20 = function(thisArg, _arguments, P, generator) {
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
var BroadcastApi = class {
  constructor(messageBus2) {
    this.messageBus = messageBus2;
  }
  sendMessage(channel, data, options) {
    return __awaiter20(this, void 0, void 0, function* () {
      yield this.messageBus.sendAsync("OBR_BROADCAST_SEND_MESSAGE", {
        channel,
        data,
        options
      });
    });
  }
  onMessage(channel, callback) {
    this.messageBus.send("OBR_BROADCAST_SUBSCRIBE", { channel });
    this.messageBus.on(`OBR_BROADCAST_MESSAGE_${channel}`, callback);
    return () => {
      this.messageBus.send("OBR_BROADCAST_UNSUBSCRIBE", { channel });
      this.messageBus.off(`OBR_BROADCAST_MESSAGE_${channel}`, callback);
    };
  }
};
var BroadcastApi_default = BroadcastApi;

// node_modules/@owlbear-rodeo/sdk/lib/builders/GenericItemBuilder.js
var GenericItemBuilder = class {
  constructor(player) {
    this._item = {
      createdUserId: player.id,
      id: v4_default(),
      name: "Item",
      zIndex: Date.now(),
      lastModified: (/* @__PURE__ */ new Date()).toISOString(),
      lastModifiedUserId: player.id,
      locked: false,
      metadata: {},
      position: { x: 0, y: 0 },
      rotation: 0,
      scale: { x: 1, y: 1 },
      type: "ITEM",
      visible: true,
      layer: "POPOVER"
    };
  }
  createdUserId(createdUserId) {
    this._item.createdUserId = createdUserId;
    return this.self();
  }
  id(id) {
    this._item.id = id;
    return this.self();
  }
  name(name) {
    this._item.name = name;
    return this.self();
  }
  description(description) {
    this._item.description = description;
    return this.self();
  }
  lastModified(lastModified) {
    this._item.lastModified = lastModified;
    return this.self();
  }
  zIndex(zIndex) {
    this._item.zIndex = zIndex;
    return this.self();
  }
  lastModifiedUserId(lastModifiedUserId) {
    this._item.lastModifiedUserId = lastModifiedUserId;
    return this.self();
  }
  locked(locked) {
    this._item.locked = locked;
    return this.self();
  }
  metadata(metadata) {
    this._item.metadata = metadata;
    return this.self();
  }
  position(position) {
    this._item.position = position;
    return this.self();
  }
  rotation(rotation) {
    this._item.rotation = rotation;
    return this.self();
  }
  scale(scale) {
    this._item.scale = scale;
    return this.self();
  }
  visible(visible) {
    this._item.visible = visible;
    return this.self();
  }
  attachedTo(attachedTo) {
    this._item.attachedTo = attachedTo;
    return this.self();
  }
  layer(layer) {
    this._item.layer = layer;
    return this.self();
  }
  disableHit(disable) {
    this._item.disableHit = disable;
    return this.self();
  }
  disableAutoZIndex(disable) {
    this._item.disableAutoZIndex = disable;
    return this.self();
  }
  disableAttachmentBehavior(disable) {
    this._item.disableAttachmentBehavior = disable;
    return this.self();
  }
  self() {
    return this;
  }
};

// node_modules/@owlbear-rodeo/sdk/lib/builders/PathBuilder.js
var PathBuilder = class extends GenericItemBuilder {
  constructor(player) {
    super(player);
    this._commands = [];
    this._fillRule = "nonzero";
    this._style = {
      fillColor: "black",
      fillOpacity: 1,
      strokeColor: "white",
      strokeOpacity: 1,
      strokeWidth: 5,
      strokeDash: []
    };
    this._item.name = "Path";
    this._item.layer = "DRAWING";
  }
  commands(commands) {
    this._commands = commands;
    return this.self();
  }
  fillRule(fillRule) {
    this._fillRule = fillRule;
    return this.self();
  }
  style(style) {
    this._style = style;
    return this.self();
  }
  fillColor(fillColor) {
    this._style.fillColor = fillColor;
    return this.self();
  }
  fillOpacity(fillOpacity) {
    this._style.fillOpacity = fillOpacity;
    return this.self();
  }
  strokeColor(strokeColor) {
    this._style.strokeColor = strokeColor;
    return this.self();
  }
  strokeOpacity(strokeOpacity) {
    this._style.strokeOpacity = strokeOpacity;
    return this.self();
  }
  strokeWidth(strokeWidth) {
    this._style.strokeWidth = strokeWidth;
    return this.self();
  }
  strokeDash(strokeDash) {
    this._style.strokeDash = strokeDash;
    return this.self();
  }
  build() {
    return Object.assign(Object.assign({}, this._item), { type: "PATH", commands: this._commands, fillRule: this._fillRule, style: this._style });
  }
};

// node_modules/js-base64/base64.mjs
var _hasBuffer = typeof Buffer === "function";
var _TD = typeof TextDecoder === "function" ? new TextDecoder() : void 0;
var _TE = typeof TextEncoder === "function" ? new TextEncoder() : void 0;
var b64ch = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
var b64chs = Array.prototype.slice.call(b64ch);
var b64tab = ((a) => {
  let tab = {};
  a.forEach((c, i) => tab[c] = i);
  return tab;
})(b64chs);
var b64re = /^(?:[A-Za-z\d+\/]{4})*?(?:[A-Za-z\d+\/]{2}(?:==)?|[A-Za-z\d+\/]{3}=?)?$/;
var _fromCC = String.fromCharCode.bind(String);
var _U8Afrom = typeof Uint8Array.from === "function" ? Uint8Array.from.bind(Uint8Array) : (it) => new Uint8Array(Array.prototype.slice.call(it, 0));
var _tidyB64 = (s) => s.replace(/[^A-Za-z0-9\+\/]/g, "");
var re_btou = /[\xC0-\xDF][\x80-\xBF]|[\xE0-\xEF][\x80-\xBF]{2}|[\xF0-\xF7][\x80-\xBF]{3}/g;
var cb_btou = (cccc) => {
  switch (cccc.length) {
    case 4:
      var cp = (7 & cccc.charCodeAt(0)) << 18 | (63 & cccc.charCodeAt(1)) << 12 | (63 & cccc.charCodeAt(2)) << 6 | 63 & cccc.charCodeAt(3), offset = cp - 65536;
      return _fromCC((offset >>> 10) + 55296) + _fromCC((offset & 1023) + 56320);
    case 3:
      return _fromCC((15 & cccc.charCodeAt(0)) << 12 | (63 & cccc.charCodeAt(1)) << 6 | 63 & cccc.charCodeAt(2));
    default:
      return _fromCC((31 & cccc.charCodeAt(0)) << 6 | 63 & cccc.charCodeAt(1));
  }
};
var btou = (b) => b.replace(re_btou, cb_btou);
var atobPolyfill = (asc) => {
  asc = asc.replace(/\s+/g, "");
  if (!b64re.test(asc))
    throw new TypeError("malformed base64.");
  asc += "==".slice(2 - (asc.length & 3));
  let u24, r1, r2;
  let binArray = [];
  for (let i = 0; i < asc.length; ) {
    u24 = b64tab[asc.charAt(i++)] << 18 | b64tab[asc.charAt(i++)] << 12 | (r1 = b64tab[asc.charAt(i++)]) << 6 | (r2 = b64tab[asc.charAt(i++)]);
    if (r1 === 64) {
      binArray.push(_fromCC(u24 >> 16 & 255));
    } else if (r2 === 64) {
      binArray.push(_fromCC(u24 >> 16 & 255, u24 >> 8 & 255));
    } else {
      binArray.push(_fromCC(u24 >> 16 & 255, u24 >> 8 & 255, u24 & 255));
    }
  }
  return binArray.join("");
};
var _atob = typeof atob === "function" ? (asc) => atob(_tidyB64(asc)) : _hasBuffer ? (asc) => Buffer.from(asc, "base64").toString("binary") : atobPolyfill;
var _toUint8Array = _hasBuffer ? (a) => _U8Afrom(Buffer.from(a, "base64")) : (a) => _U8Afrom(_atob(a).split("").map((c) => c.charCodeAt(0)));
var _decode = _hasBuffer ? (a) => Buffer.from(a, "base64").toString("utf8") : _TD ? (a) => _TD.decode(_toUint8Array(a)) : (a) => btou(_atob(a));
var _unURI = (a) => _tidyB64(a.replace(/[-_]/g, (m0) => m0 == "-" ? "+" : "/"));
var decode = (src) => _decode(_unURI(src));

// node_modules/@owlbear-rodeo/sdk/lib/common/getDetails.js
function getDetails() {
  const urlSearchParams = new URLSearchParams(window.location.search);
  const ref = urlSearchParams.get("obrref");
  let origin = "";
  let roomId = "";
  if (ref) {
    const decodedRef = decode(ref);
    const parts = decodedRef.split(" ");
    if (parts.length === 2) {
      origin = parts[0];
      roomId = parts[1];
    }
  }
  return { origin, roomId };
}

// node_modules/@owlbear-rodeo/sdk/lib/types/items/Image.js
function isImage(item) {
  return item.type === "IMAGE";
}

// node_modules/@owlbear-rodeo/sdk/lib/types/items/Path.js
var Command;
(function(Command2) {
  Command2[Command2["MOVE"] = 0] = "MOVE";
  Command2[Command2["LINE"] = 1] = "LINE";
  Command2[Command2["QUAD"] = 2] = "QUAD";
  Command2[Command2["CONIC"] = 3] = "CONIC";
  Command2[Command2["CUBIC"] = 4] = "CUBIC";
  Command2[Command2["CLOSE"] = 5] = "CLOSE";
})(Command || (Command = {}));

// node_modules/@owlbear-rodeo/sdk/lib/index.js
var details = getDetails();
var messageBus = new MessageBus_default(details.origin, details.roomId);
var viewportApi = new ViewportApi_default(messageBus);
var playerApi = new PlayerApi_default(messageBus);
var partyApi = new PartyApi_default(messageBus);
var notificationApi = new NotificationApi_default(messageBus);
var sceneApi = new SceneApi_default(messageBus);
var contextMenuApi = new ContextMenuApi_default(messageBus);
var toolApi = new ToolApi_default(messageBus);
var popoverApi = new PopoverApi_default(messageBus);
var modalApi = new ModalApi_default(messageBus);
var actionApi = new ActionApi_default(messageBus);
var interactionApi = new InteractionApi_default(messageBus);
var roomApi = new RoomApi_default(messageBus);
var themeApi = new ThemeApi_default(messageBus);
var assetsApi = new AssetsApi_default(messageBus);
var broadcastApi = new BroadcastApi_default(messageBus);
var OBR = {
  onReady: (callback) => {
    if (messageBus.ready) {
      callback();
    } else {
      messageBus.once("OBR_READY", () => callback());
    }
  },
  get isReady() {
    return messageBus.ready;
  },
  viewport: viewportApi,
  player: playerApi,
  party: partyApi,
  notification: notificationApi,
  scene: sceneApi,
  contextMenu: contextMenuApi,
  tool: toolApi,
  popover: popoverApi,
  modal: modalApi,
  action: actionApi,
  interaction: interactionApi,
  room: roomApi,
  theme: themeApi,
  assets: assetsApi,
  broadcast: broadcastApi,
  /** True if the current site is embedded in an instance of Owlbear Rodeo */
  isAvailable: Boolean(details.origin)
};
function buildPath() {
  return new PathBuilder(playerApi);
}
var lib_default = OBR;

// shared.js
var EXTENSION_ID = "com.codex.body-hp";
var META_KEY = `${EXTENSION_ID}/data`;
var OVERLAY_KEY = `${EXTENSION_ID}/overlayFor`;
var BODY_ORDER = ["Head", "L.Arm", "R.Arm", "Torso", "L.Leg", "R.Leg"];
var ROLL_HISTORY_LIMIT = 12;
var COMBAT_SKILL_CATEGORY = "combat";
var APPLIED_SKILL_CATEGORY = "applied";
var MELEE_SKILL_NAME = "Melee";
var PARRY_SKILL_NAME = "Parry";
var LEGACY_MELEE_SKILL_NAMES = /* @__PURE__ */ new Set(["Hand", "Cold", "\u0420\u0443\u043A\u043E\u043F\u0430\u0448\u043D\u044B\u0439"]);
var LEGACY_REMOVED_SKILLS = /* @__PURE__ */ new Set(["Hand", "Cold", "Throwing", "Rifle", "Turrets"]);
var VISUAL_VERSION = 3;
var RING_COLORS = {
  full: "#73FF5A",
  half: "#FFAF22",
  kaputt: "#FF460D",
  base: "#000000",
  border: "#050505"
};
var OUTER_SEGMENTS = [
  { part: "Head", angle: -90, span: 30 },
  { part: "R.Arm", angle: -18, span: 30 },
  { part: "R.Leg", angle: 54, span: 30 },
  { part: "L.Leg", angle: 126, span: 30 },
  { part: "L.Arm", angle: 198, span: 30 }
];
var DEFAULT_ODYSSEY_SKILLS = {
  [MELEE_SKILL_NAME]: 0,
  [PARRY_SKILL_NAME]: 0
};
var DEFAULT_ODYSSEY_SKILL_CATEGORIES = {
  [MELEE_SKILL_NAME]: COMBAT_SKILL_CATEGORY,
  [PARRY_SKILL_NAME]: COMBAT_SKILL_CATEGORY
};
var DEFAULT_ODYSSEY_SKILL_STRENGTH_BONUSES = {
  [MELEE_SKILL_NAME]: true,
  [PARRY_SKILL_NAME]: false
};
var BODY_DEFAULTS = {
  Head: { current: 1, max: 1, armor: 0, minor: 0, serious: 0 },
  "L.Arm": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  "R.Arm": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  Torso: { current: 3, max: 3, armor: 6, minor: 0, serious: 0 },
  "L.Leg": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 },
  "R.Leg": { current: 2, max: 2, armor: 2, minor: 0, serious: 0 }
};
var DEFAULT_TRACKER_DATA = {
  enabled: true,
  minor: 0,
  serious: 0,
  body: structuredClone(BODY_DEFAULTS),
  identity: {
    playerId: "",
    characterId: ""
  },
  lastRoll: null,
  history: [],
  sync: {
    lastEventId: 0,
    lastSyncedAt: null
  },
  odyssey: {
    owner: {
      playerId: "",
      playerName: ""
    },
    skills: structuredClone(DEFAULT_ODYSSEY_SKILLS),
    skillCategories: structuredClone(DEFAULT_ODYSSEY_SKILL_CATEGORIES),
    skillStrengthBonuses: structuredClone(DEFAULT_ODYSSEY_SKILL_STRENGTH_BONUSES),
    attributes: {
      Strength: 0,
      Agility: 0,
      Reaction: 0,
      Endurance: 0,
      Perception: 0,
      Intelligence: 0,
      Charisma: 0,
      Willpower: 0,
      Magic: 0
    },
    weapons: {
      melee: [],
      ranged: []
    }
  }
};
function deepClone(value) {
  return structuredClone(value);
}
function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}
function numberOrFallback(value, fallback) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}
function sanitizeTrackerData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA);
  if (!raw || typeof raw !== "object") return next;
  next.enabled = raw.enabled !== false;
  next.minor = clamp(Number(raw.minor ?? 0) || 0, 0, 4);
  next.serious = clamp(Number(raw.serious ?? 0) || 0, 0, 2);
  next.identity.playerId = String(raw.identity?.playerId ?? "").trim();
  next.identity.characterId = String(raw.identity?.characterId ?? "").trim();
  next.lastRoll = sanitizeRollSummary(raw.lastRoll);
  next.history = Array.isArray(raw.history) ? raw.history.map(sanitizeRollSummary).filter(Boolean).slice(0, ROLL_HISTORY_LIMIT) : [];
  next.sync.lastEventId = Math.max(0, Number(raw.sync?.lastEventId ?? 0) || 0);
  next.sync.lastSyncedAt = raw.sync?.lastSyncedAt ? String(raw.sync.lastSyncedAt) : null;
  next.odyssey = sanitizeOdysseyData(raw.odyssey);
  for (const partName of BODY_ORDER) {
    const source = raw.body?.[partName] ?? {};
    const part = next.body[partName];
    part.max = clamp(numberOrFallback(source.max, part.max), 0, 99);
    part.current = clamp(
      numberOrFallback(source.current, part.current),
      0,
      part.max
    );
    part.armor = clamp(numberOrFallback(source.armor, part.armor), 0, 99);
    part.minor = clamp(numberOrFallback(source.minor, part.minor), 0, 3);
    part.serious = clamp(numberOrFallback(source.serious, part.serious), 0, 1);
  }
  return next;
}
function sanitizeOdysseyData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA.odyssey);
  if (!raw || typeof raw !== "object") return next;
  next.owner.playerId = String(raw.owner?.playerId ?? "").trim();
  next.owner.playerName = String(raw.owner?.playerName ?? "").trim();
  const rawSkills = raw.skills && typeof raw.skills === "object" ? raw.skills : {};
  const rawSkillCategories = raw.skillCategories && typeof raw.skillCategories === "object" ? raw.skillCategories : {};
  const rawSkillStrengthBonuses = raw.skillStrengthBonuses && typeof raw.skillStrengthBonuses === "object" ? raw.skillStrengthBonuses : {};
  const migratedMeleeValue = Math.max(
    Number(rawSkills[MELEE_SKILL_NAME] ?? 0) || 0,
    ...Array.from(LEGACY_MELEE_SKILL_NAMES).map((skillName) => Number(rawSkills[skillName] ?? 0) || 0),
    Number(DEFAULT_ODYSSEY_SKILLS[MELEE_SKILL_NAME] ?? 0) || 0
  );
  const migratedParryValue = Math.max(
    Number(rawSkills[PARRY_SKILL_NAME] ?? 0) || 0,
    Number(raw.attributes?.Parry ?? 0) || 0,
    Number(DEFAULT_ODYSSEY_SKILLS[PARRY_SKILL_NAME] ?? 0) || 0
  );
  next.skills[MELEE_SKILL_NAME] = clamp(migratedMeleeValue, 0, 10);
  next.skillCategories[MELEE_SKILL_NAME] = COMBAT_SKILL_CATEGORY;
  next.skillStrengthBonuses[MELEE_SKILL_NAME] = true;
  next.skills[PARRY_SKILL_NAME] = clamp(migratedParryValue, 0, 10);
  next.skillCategories[PARRY_SKILL_NAME] = COMBAT_SKILL_CATEGORY;
  next.skillStrengthBonuses[PARRY_SKILL_NAME] = false;
  for (const [key, value] of Object.entries(rawSkills)) {
    const normalizedKey = String(key).trim();
    if (!normalizedKey) continue;
    if (normalizedKey === MELEE_SKILL_NAME || normalizedKey === PARRY_SKILL_NAME || LEGACY_MELEE_SKILL_NAMES.has(normalizedKey) || LEGACY_REMOVED_SKILLS.has(normalizedKey)) {
      continue;
    }
    next.skills[normalizedKey] = clamp(Number(value) || 0, 0, 10);
    const categoryValue = String(
      rawSkillCategories[normalizedKey] ?? rawSkillCategories[key] ?? ""
    ).toLowerCase();
    next.skillCategories[normalizedKey] = categoryValue === COMBAT_SKILL_CATEGORY ? COMBAT_SKILL_CATEGORY : APPLIED_SKILL_CATEGORY;
    next.skillStrengthBonuses[normalizedKey] = Boolean(
      rawSkillStrengthBonuses[normalizedKey] ?? rawSkillStrengthBonuses[key] ?? false
    );
  }
  for (const key of Object.keys(next.attributes)) {
    const fallbackValue = key === "Magic" ? raw.attributes?.[key] ?? raw.attributes?.Psionics ?? 0 : raw.attributes?.[key] ?? 0;
    next.attributes[key] = clamp(Number(fallbackValue) || 0, 0, 15);
  }
  next.weapons.melee = sanitizeWeapons(raw.weapons?.melee);
  next.weapons.ranged = sanitizeWeapons(raw.weapons?.ranged);
  return next;
}
function sanitizeWeapons(raw) {
  if (!Array.isArray(raw)) return [];
  return raw.filter((item) => item && typeof item === "object").map((item) => ({
    name: String(item.name ?? "").trim() || "Weapon",
    damage: clamp(Number(item.damage ?? 0) || 0, -99, 99)
  })).slice(0, 20);
}
function sanitizeRollSummary(raw) {
  if (!raw || typeof raw !== "object") return null;
  const eventId = Math.max(0, Number(raw.eventId ?? 0) || 0);
  const summary = String(raw.summary ?? "").trim();
  const actorName = String(raw.actorName ?? "").trim();
  const outcome = String(raw.outcome ?? "").trim();
  const total = raw.total == null ? null : Number(raw.total) || 0;
  const targetPart = String(raw.targetPart ?? "").trim();
  const timestamp = raw.timestamp ? String(raw.timestamp) : null;
  const source = String(raw.source ?? "bridge").trim();
  if (!summary && !actorName && total == null && !outcome) {
    return null;
  }
  return {
    eventId,
    summary,
    actorName,
    outcome,
    total,
    targetPart,
    timestamp,
    source
  };
}
function getTrackerData(item) {
  return sanitizeTrackerData(item?.metadata?.[META_KEY]);
}
function isCharacterToken(item) {
  return Boolean(item) && isImage(item) && item.layer === "CHARACTER";
}
function isTrackedCharacter(item) {
  return isCharacterToken(item) && item.metadata?.[META_KEY]?.enabled === true;
}
function isOverlayItem(item) {
  return Boolean(item?.metadata?.[OVERLAY_KEY]);
}
function getCharacterName(item) {
  if (!item) return "Unnamed character";
  const byName = typeof item.name === "string" ? item.name.trim() : "";
  if (byName) return byName;
  return `Character ${item.id.slice(0, 6)}`;
}
function sortCharacters(items) {
  return [...items].sort(
    (left, right) => getCharacterName(left).localeCompare(getCharacterName(right))
  );
}
function formatOverlayText(data) {
  const body = data.body;
  return [
    `Head ${body["Head"].current}/${body["Head"].max}(${body["Head"].armor}) | L.Arm ${body["L.Arm"].current}/${body["L.Arm"].max}(${body["L.Arm"].armor}) | R.Arm ${body["R.Arm"].current}/${body["R.Arm"].max}(${body["R.Arm"].armor})`,
    `Torso ${body["Torso"].current}/${body["Torso"].max}(${body["Torso"].armor}) | L.Leg ${body["L.Leg"].current}/${body["L.Leg"].max}(${body["L.Leg"].armor}) | R.Leg ${body["R.Leg"].current}/${body["R.Leg"].max}(${body["R.Leg"].armor})`
  ].join("\n");
}
function getOdysseyData(item) {
  return sanitizeOdysseyData(getTrackerData(item).odyssey);
}
function canPlayerControlToken(playerRole2, playerId2, token) {
  if (!token || !isCharacterToken(token)) return false;
  if (playerRole2 === "GM") return true;
  const odyssey = getOdysseyData(token);
  return Boolean(playerId2) && odyssey.owner.playerId === playerId2;
}
function getAvailableWeapons(token, mode = "melee") {
  const odyssey = getOdysseyData(token);
  const list = mode === "ranged" ? odyssey.weapons.ranged : odyssey.weapons.melee;
  return list.length ? list : [{ name: "Default", damage: 0 }];
}
function getBodyTotals(data) {
  return BODY_ORDER.reduce(
    (accumulator, partName) => {
      accumulator.current += data.body[partName].current;
      accumulator.max += data.body[partName].max;
      return accumulator;
    },
    { current: 0, max: 0 }
  );
}
function getEffectiveSize(token) {
  const scaleX = Math.abs(token.scale?.x ?? 1);
  const scaleY = Math.abs(token.scale?.y ?? 1);
  return {
    width: (token.width || 140) * scaleX,
    height: (token.height || 140) * scaleY
  };
}
async function getTokenMetrics(token) {
  const effectiveSize = getEffectiveSize(token);
  let center = token.position;
  let width = effectiveSize.width;
  let height = effectiveSize.height;
  try {
    const bounds = await lib_default.scene.items.getItemBounds([token.id]);
    if (bounds?.width > 0 && bounds?.height > 0) {
      center = bounds.center;
      width = bounds.width;
      height = bounds.height;
    }
  } catch (error) {
    console.warn("[Body HP] Unable to read token bounds, using fallback size", error);
  }
  let gridDpi = 150;
  try {
    gridDpi = await lib_default.scene.grid.getDpi() || gridDpi;
  } catch (error) {
    console.warn("[Body HP] Unable to read grid dpi, using fallback size", error);
  }
  const scaleFactor = Math.max(
    Math.abs(token.scale?.x ?? 1),
    Math.abs(token.scale?.y ?? 1),
    1
  );
  const visibleDiameter = Math.max(
    width,
    height,
    effectiveSize.width,
    effectiveSize.height,
    gridDpi * scaleFactor,
    56
  );
  const tokenRadius = visibleDiameter / 2;
  const tokenGap = 0;
  const torsoThickness = Math.max(5, visibleDiameter * 0.035);
  const torsoInnerRadius = tokenRadius + tokenGap;
  const torsoOuterRadius = torsoInnerRadius + torsoThickness;
  const ringGap = 0;
  const outerThickness = Math.max(8, visibleDiameter * 0.08);
  const outerInnerRadius = torsoOuterRadius + ringGap;
  const outerRadius = outerInnerRadius + outerThickness;
  return {
    center,
    visibleDiameter,
    outerRadius,
    outerInnerRadius,
    torsoOuterRadius,
    torsoInnerRadius
  };
}
function polar(radius, angle) {
  const radians = angle * Math.PI / 180;
  return {
    x: radius * Math.cos(radians),
    y: radius * Math.sin(radians)
  };
}
function arcPoints(radius, startAngle, endAngle, segments = 18) {
  const points = [];
  for (let index = 0; index <= segments; index += 1) {
    const ratio = index / segments;
    const angle = startAngle + (endAngle - startAngle) * ratio;
    points.push(polar(radius, angle));
  }
  return points;
}
function buildAnnulusCommands(radiusOuter, radiusInner) {
  const outer = arcPoints(radiusOuter, -180, 180, 36);
  const inner = arcPoints(radiusInner, -180, 180, 36);
  const commands = [[Command.MOVE, outer[0].x, outer[0].y]];
  for (const point of outer.slice(1)) {
    commands.push([Command.LINE, point.x, point.y]);
  }
  commands.push([Command.CLOSE]);
  commands.push([Command.MOVE, inner[0].x, inner[0].y]);
  for (const point of inner) {
    commands.push([Command.LINE, point.x, point.y]);
  }
  commands.push([Command.CLOSE]);
  return commands;
}
function buildSectorCommands(radiusOuter, radiusInner, centerAngle, spanAngle) {
  const startAngle = centerAngle - spanAngle / 2;
  const endAngle = centerAngle + spanAngle / 2;
  const outer = arcPoints(radiusOuter, startAngle, endAngle, 10);
  const inner = arcPoints(radiusInner, endAngle, startAngle, 10);
  const commands = [[Command.MOVE, outer[0].x, outer[0].y]];
  for (const point of outer.slice(1)) {
    commands.push([Command.LINE, point.x, point.y]);
  }
  for (const point of inner) {
    commands.push([Command.LINE, point.x, point.y]);
  }
  commands.push([Command.CLOSE]);
  return commands;
}
function getPartColor(part) {
  if (part.max <= 0 || part.current <= 0) return RING_COLORS.kaputt;
  if (part.current < part.max) return RING_COLORS.half;
  return RING_COLORS.full;
}
function buildRingItem(token, metrics, kind, commands, fillColor, zIndex = 0, fillRule = "nonzero") {
  return buildPath().name(`${kind}: ${getCharacterName(token)}`).commands(commands).fillRule(fillRule).fillColor(fillColor).fillOpacity(1).strokeColor(RING_COLORS.border).strokeOpacity(1).strokeWidth(0.75).position(metrics.center).rotation(0).zIndex(Date.now() + zIndex).attachedTo(token.id).disableAttachmentBehavior(["ROTATION"]).layer("ATTACHMENT").locked(true).disableHit(true).metadata({
    [OVERLAY_KEY]: token.id,
    kind,
    visualVersion: VISUAL_VERSION
  }).build();
}
function buildOverlayItems(token, data, metrics) {
  const items = [];
  items.push(
    buildRingItem(
      token,
      metrics,
      "outer-base",
      buildAnnulusCommands(metrics.outerRadius, metrics.outerInnerRadius),
      RING_COLORS.base,
      0,
      "evenodd"
    )
  );
  for (const segment of OUTER_SEGMENTS) {
    items.push(
      buildRingItem(
        token,
        metrics,
        `segment-${segment.part}`,
        buildSectorCommands(
          metrics.outerRadius,
          metrics.outerInnerRadius,
          segment.angle,
          segment.span
        ),
        getPartColor(data.body[segment.part]),
        1
      )
    );
  }
  items.push(
    buildRingItem(
      token,
      metrics,
      "torso-ring",
      buildAnnulusCommands(metrics.torsoOuterRadius, metrics.torsoInnerRadius),
      getPartColor(data.body.Torso),
      2,
      "evenodd"
    )
  );
  return items;
}
async function updateTrackerData(tokenId, updater) {
  await lib_default.scene.items.updateItems([tokenId], (items) => {
    const token = items[0];
    if (!token) return;
    token.metadata ?? (token.metadata = {});
    token.metadata[META_KEY] = sanitizeTrackerData(
      updater(getTrackerData(token))
    );
  });
}
async function removeOverlaysForToken(tokenId, items) {
  const sceneItems2 = items ?? await lib_default.scene.items.getItems();
  const overlayIds = sceneItems2.filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId).map((item) => item.id);
  if (overlayIds.length) {
    await lib_default.scene.items.deleteItems(overlayIds);
  }
}
async function ensureOverlayForToken(tokenId, items) {
  const sceneItems2 = items ?? await lib_default.scene.items.getItems();
  const token = sceneItems2.find((item) => item.id === tokenId);
  if (!token || !isCharacterToken(token)) return;
  await removeOverlaysForToken(tokenId, sceneItems2);
  if (!isTrackedCharacter(token)) return;
  const metrics = await getTokenMetrics(token);
  await lib_default.scene.items.addItems(
    buildOverlayItems(token, getTrackerData(token), metrics)
  );
}
async function syncTrackedOverlays() {
  const items = await lib_default.scene.items.getItems();
  const byId = new Map(items.map((item) => [item.id, item]));
  const staleOverlayIds = items.filter(isOverlayItem).filter((item) => {
    const token = byId.get(item.metadata[OVERLAY_KEY]);
    return !token || !isTrackedCharacter(token);
  }).map((item) => item.id);
  if (staleOverlayIds.length) {
    await lib_default.scene.items.deleteItems(staleOverlayIds);
  }
  const trackedTokens = items.filter(isTrackedCharacter);
  for (const token of trackedTokens) {
    await ensureOverlayForToken(token.id, items);
  }
}

// odyssey_rules.js
function rollPercent() {
  return Math.floor(Math.random() * 100) + 1;
}
function rollDice(sides, modifier = 0, count = 1) {
  const safeSides = clamp(Number(sides) || 0, 2, 1e3);
  const safeCount = clamp(Number(count) || 0, 1, 100);
  const rolls = Array.from({ length: safeCount }, () => Math.floor(Math.random() * safeSides) + 1);
  const subtotal = rolls.reduce((sum, roll) => sum + roll, 0);
  return {
    roll: rolls[0] ?? 0,
    rolls,
    count: safeCount,
    sides: safeSides,
    subtotal,
    modifier: Number(modifier) || 0,
    total: subtotal + (Number(modifier) || 0)
  };
}
function calculateAccuracy(attackSkill, attackBonuses = 0, attackPenalties = 0, defenseBonuses = 0, defensePenalties = 0, parry = 0) {
  const attackRoll = rollPercent();
  const defenseRoll = rollPercent();
  const attackTotal = attackRoll + clamp(Number(attackSkill) || 0, 0, 10) * 10 + (Number(attackBonuses) || 0) - (Number(attackPenalties) || 0);
  const defenseTotal = defenseRoll + (Number(defenseBonuses) || 0) - (Number(defensePenalties) || 0) + clamp(Number(parry) || 0, 0, 10) * 10;
  return {
    attackRoll,
    defenseRoll,
    attackTotal,
    defenseTotal
  };
}
function calculateDamage(attackResult, defenseResult, weaponDamage = 0, armor = 0) {
  const totalAttack = (Number(attackResult) || 0) + (Number(weaponDamage) || 0);
  const totalDefense = (Number(defenseResult) || 0) + (Number(armor) || 0);
  const damageDiff = totalAttack - totalDefense;
  let label = "No damage.";
  let crit = 0;
  let serious = 0;
  let minor = 0;
  if (damageDiff >= 31) {
    label = "Critical damage.";
    crit = 1;
  } else if (damageDiff >= 6) {
    label = "Serious hit.";
    serious = 1;
  } else if (damageDiff > 0) {
    label = "Minor damage.";
    minor = 1;
  }
  return {
    totalAttack,
    totalDefense,
    damageDiff,
    label,
    crit,
    serious,
    minor
  };
}
function resolveAttack({
  attackSkill = 0,
  weaponDamage = 0,
  defenseBonuses = 0,
  defensePenalties = 0,
  attackBonuses = 0,
  attackPenalties = 0,
  parry = 0,
  targetPart = "Torso",
  targetArmor = 0
}) {
  const part = BODY_ORDER.includes(targetPart) ? targetPart : "Torso";
  const accuracy = calculateAccuracy(
    attackSkill,
    attackBonuses,
    attackPenalties,
    defenseBonuses,
    defensePenalties,
    parry
  );
  const criticalSuccess = accuracy.attackRoll >= 95;
  const criticalFailure = accuracy.attackRoll <= 5;
  const hit = criticalSuccess || !criticalFailure && accuracy.attackTotal > accuracy.defenseTotal;
  let outcome = "failure";
  let damage = null;
  let bodyDelta = 0;
  if (criticalSuccess) {
    outcome = "critical-success";
    const baseDamage = calculateDamage(
      accuracy.attackTotal,
      accuracy.defenseTotal,
      weaponDamage,
      targetArmor
    );
    damage = {
      ...baseDamage,
      label: "Critical hit: 2 Crit.",
      crit: 2,
      serious: 0,
      minor: 0
    };
    bodyDelta = -2;
  } else if (criticalFailure) {
    outcome = "critical-failure";
  } else if (hit) {
    outcome = "success";
    damage = calculateDamage(accuracy.attackTotal, accuracy.defenseTotal, weaponDamage, targetArmor);
    bodyDelta = -(damage.crit || 0);
  }
  return {
    ...accuracy,
    targetPart: part,
    targetArmor: Number(targetArmor) || 0,
    weaponDamage: Number(weaponDamage) || 0,
    outcome,
    hit,
    damage,
    bodyDelta,
    summary: buildAttackSummary({
      part,
      outcome,
      damage,
      attackRoll: accuracy.attackRoll,
      attackTotal: accuracy.attackTotal,
      defenseTotal: accuracy.defenseTotal
    })
  };
}
function buildAttackSummary({ part, outcome, damage, attackRoll, attackTotal, defenseTotal }) {
  if (outcome === "critical-success") {
    return `Critical success to ${part}. Roll ${attackRoll}; ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  if (outcome === "critical-failure") {
    return `Critical failure. Roll ${attackRoll}.`;
  }
  if (outcome === "success") {
    return `Hit ${part}. ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  return `Missed ${part}. ${attackTotal} vs ${defenseTotal}.`;
}

// main.js
var DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";
var DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
var DEBUG_ENTRY_LIMIT = 250;
var TARGET_PICK_TOOL_ID = "com.codex.body-hp/attack-target-picker";
var TARGET_PICK_MODE_ID = "pick-target";
var TARGET_HIGHLIGHT_KEY = "com.codex.body-hp/local-attack-target";
var EXTENSION_ICON_URL = new URL("./icon.svg", window.location.href).href;
var TARGET_PICK_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
    <circle cx="16" cy="16" r="3.25" fill="#ef4444" stroke="#7f1d1d" stroke-width="1.5"/>
    <path d="M16 2v8M16 22v8M2 16h8M22 16h8" stroke="#ef4444" stroke-width="2.5" stroke-linecap="round"/>
    <path d="M16 6v4M16 22v4M6 16h4M22 16h4" stroke="#7f1d1d" stroke-width="1.2" stroke-linecap="round"/>
  </svg>`
)}") 16 16, crosshair`;
var CORE_COMBAT_SKILLS = Object.keys(DEFAULT_ODYSSEY_SKILLS);
var ATTACK_ONLY_EXCLUDED_SKILLS = /* @__PURE__ */ new Set([PARRY_SKILL_NAME]);
var ATTRIBUTE_UI_FIELDS = [
  ["Strength", "Strength"],
  ["Agility", "Agility"],
  ["Reaction", "Reaction"],
  ["Endurance", "Endurance"],
  ["Perception", "Perception"],
  ["Intelligence", "Intelligence"],
  ["Charisma", "Charisma"],
  ["Willpower", "Willpower"],
  ["Magic", "Magic"]
];
var ui = {
  roleBadge: document.getElementById("roleBadge"),
  refreshBtn: document.getElementById("refreshBtn"),
  syncBtn: document.getElementById("syncBtn"),
  statusBox: document.getElementById("statusBox"),
  selectionHint: document.getElementById("selectionHint"),
  selectedTokenPanel: document.getElementById("selectedTokenPanel"),
  debugConsole: document.getElementById("debugConsole"),
  clearDebugBtn: document.getElementById("clearDebugBtn"),
  trackedSection: document.getElementById("trackedSection"),
  trackedCount: document.getElementById("trackedCount"),
  trackedList: document.getElementById("trackedList"),
  allTokensSection: document.getElementById("allTokensSection"),
  allCount: document.getElementById("allCount"),
  allTokensList: document.getElementById("allTokensList")
};
var playerRole = "PLAYER";
var playerId = "";
var playerName = "";
var playerColor = "#facc15";
var sceneItems = [];
var selectionIds = [];
var activeTokenId = null;
var debugEntries = [];
var partyPlayers = [];
var gmPrivateEntries = [];
var collapsibleSectionState = /* @__PURE__ */ new Map();
var attackFormDrafts = /* @__PURE__ */ new Map();
var inputAutosaveTimers = /* @__PURE__ */ new Map();
var selectionPollTimer = null;
var targetPickState = {
  active: false,
  attackerTokenId: null,
  previousToolId: "",
  previousModeId: void 0,
  toolReady: false,
  restoring: false
};
function sanitizeDebugEntries(raw) {
  if (!Array.isArray(raw)) return [];
  return raw.filter((entry) => entry && typeof entry === "object").map((entry) => ({
    id: Number(entry.id) || Date.now(),
    title: String(entry.title ?? "Debug"),
    body: String(entry.body ?? ""),
    kind: String(entry.kind ?? "info"),
    timestamp: String(entry.timestamp ?? "")
  })).slice(0, DEBUG_ENTRY_LIMIT);
}
function mergeDebugEntries(...entryGroups) {
  const merged = /* @__PURE__ */ new Map();
  for (const group of entryGroups) {
    for (const entry of sanitizeDebugEntries(group)) {
      merged.set(entry.id, entry);
    }
  }
  return [...merged.values()].sort((left, right) => Number(right.id) - Number(left.id)).slice(0, DEBUG_ENTRY_LIMIT);
}
function setStatus(message, kind = "info") {
  ui.statusBox.textContent = message;
  ui.statusBox.className = `status ${kind}`;
  console[kind === "error" ? "error" : "log"](`[Body HP] ${message}`);
}
function escapeHtml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}
function getSkillCategory(odyssey, skillName) {
  return odyssey?.skillCategories?.[skillName] === COMBAT_SKILL_CATEGORY ? COMBAT_SKILL_CATEGORY : APPLIED_SKILL_CATEGORY;
}
function getSkillStrengthBonusFlag(odyssey, skillName) {
  return Boolean(odyssey?.skillStrengthBonuses?.[skillName]);
}
function getSortedSkillEntries(odyssey) {
  return Object.entries(odyssey?.skills ?? {}).sort(
    ([left], [right]) => left.localeCompare(right)
  );
}
function getCombatSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === COMBAT_SKILL_CATEGORY
  );
}
function getAttackSkillEntries(odyssey) {
  return getCombatSkillEntries(odyssey).filter(
    ([skillName]) => !ATTACK_ONLY_EXCLUDED_SKILLS.has(skillName)
  );
}
function getAppliedSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === APPLIED_SKILL_CATEGORY
  );
}
function buildSkillOptions(skillEntries, selectedValue = "") {
  return skillEntries.map(
    ([key, value]) => `<option value="${escapeHtml(key)}" ${key === selectedValue ? "selected" : ""}>${escapeHtml(key)} (${value})</option>`
  ).join("");
}
function buildGroupedSkillOptions(odyssey, selectedValue = "") {
  const combatOptions = buildSkillOptions(getCombatSkillEntries(odyssey), selectedValue);
  const appliedOptions = buildSkillOptions(getAppliedSkillEntries(odyssey), selectedValue);
  return [
    combatOptions ? `<optgroup label="Combat">${combatOptions}</optgroup>` : "",
    appliedOptions ? `<optgroup label="Applied">${appliedOptions}</optgroup>` : ""
  ].filter(Boolean).join("");
}
function getTransientFieldKey(field) {
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }
  if (field.dataset.attackField) return `attack:${field.dataset.attackField}`;
  if (field.dataset.rollField) return `roll:${field.dataset.rollField}`;
  if (field.dataset.rollCharField) return `roll-char:${field.dataset.rollCharField}`;
  if (field.dataset.rollSkillField) return `roll-skill:${field.dataset.rollSkillField}`;
  if (field.dataset.gmRollField) return `gm-roll:${field.dataset.gmRollField}`;
  if (field.dataset.skillField) return `new-skill:${field.dataset.skillField}`;
  if (field.dataset.action === "select-owner-player") return "owner";
  if (field.dataset.action === "set-odyssey-skill") return `skill:${field.dataset.skill ?? ""}`;
  if (field.dataset.action === "set-skill-strength-bonus") {
    return `skill-strength:${field.dataset.skill ?? ""}`;
  }
  if (field.dataset.action === "set-odyssey-attribute") {
    return `attribute:${field.dataset.attribute ?? ""}`;
  }
  if (field.dataset.action === "set-field") {
    return `part:${field.dataset.part ?? ""}:${field.dataset.field ?? ""}`;
  }
  return "";
}
function shouldPreserveFieldValue(fieldKey, focusedKey) {
  return fieldKey === focusedKey || fieldKey.startsWith("attack:") || fieldKey.startsWith("roll:") || fieldKey.startsWith("roll-char:") || fieldKey.startsWith("roll-skill:") || fieldKey.startsWith("gm-roll:") || fieldKey.startsWith("new-skill:");
}
function captureSelectedPanelState() {
  if (!activeTokenId || !ui.selectedTokenPanel.childElementCount) return null;
  let focusedKey = "";
  let selectionStart = null;
  let selectionEnd = null;
  const activeField = document.activeElement;
  if ((activeField instanceof HTMLInputElement || activeField instanceof HTMLSelectElement) && ui.selectedTokenPanel.contains(activeField)) {
    focusedKey = getTransientFieldKey(activeField);
    if (activeField instanceof HTMLInputElement && activeField.type !== "number") {
      selectionStart = activeField.selectionStart;
      selectionEnd = activeField.selectionEnd;
    }
  }
  const fields = [];
  ui.selectedTokenPanel.querySelectorAll("input, select").forEach((field) => {
    if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) return;
    const key = getTransientFieldKey(field);
    if (!key || !shouldPreserveFieldValue(key, focusedKey)) return;
    fields.push([
      key,
      field instanceof HTMLInputElement && field.type === "checkbox" ? { kind: "checkbox", checked: field.checked } : { kind: "value", value: field.value }
    ]);
  });
  return {
    tokenId: activeTokenId,
    fields,
    focusedKey,
    selectionStart,
    selectionEnd
  };
}
function restoreSelectedPanelState(panelState) {
  if (!panelState || panelState.tokenId !== activeTokenId) return;
  const fieldValues = new Map(panelState.fields ?? []);
  let focusedField = null;
  ui.selectedTokenPanel.querySelectorAll("input, select").forEach((field) => {
    if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) return;
    const key = getTransientFieldKey(field);
    if (!key || !fieldValues.has(key)) return;
    const nextValue = fieldValues.get(key);
    if (field instanceof HTMLInputElement && field.type === "checkbox" && nextValue && typeof nextValue === "object" && nextValue.kind === "checkbox") {
      field.checked = Boolean(nextValue.checked);
    } else {
      const normalizedValue = nextValue && typeof nextValue === "object" && nextValue.kind === "value" ? nextValue.value : typeof nextValue === "string" ? nextValue : "";
      if (field instanceof HTMLSelectElement) {
        const hasMatchingOption = Array.from(field.options).some(
          (option) => option.value === normalizedValue
        );
        if (!hasMatchingOption) return;
      }
      field.value = normalizedValue;
    }
    if (key === panelState.focusedKey) {
      focusedField = field;
    }
  });
  if (!focusedField) return;
  focusedField.focus({ preventScroll: true });
  if (focusedField instanceof HTMLInputElement && typeof panelState.selectionStart === "number" && typeof panelState.selectionEnd === "number") {
    try {
      focusedField.setSelectionRange(panelState.selectionStart, panelState.selectionEnd);
    } catch (error) {
      console.warn("[Body HP] Unable to restore cursor position", error);
    }
  }
}
function getSortedPartyPlayers() {
  return [...partyPlayers].sort(
    (left, right) => String(left?.name ?? "").localeCompare(String(right?.name ?? ""))
  );
}
function getCharacters() {
  return sortCharacters(sceneItems.filter(isCharacterToken));
}
function getTrackedCharacters() {
  return getCharacters().filter(isTrackedCharacter);
}
function getCharacterById(tokenId) {
  return getCharacters().find((item) => item.id === tokenId) ?? null;
}
function resolveActiveTokenId() {
  const characters = getCharacters();
  const selectedCharacterId = selectionIds.find(
    (id) => characters.some((character) => character.id === id)
  );
  if (selectedCharacterId) return selectedCharacterId;
  if (activeTokenId && characters.some((character) => character.id === activeTokenId)) {
    return activeTokenId;
  }
  const firstTracked = getTrackedCharacters()[0];
  if (firstTracked) return firstTracked.id;
  return characters[0]?.id ?? null;
}
function isEditable() {
  return playerRole === "GM";
}
function canUseToken(token) {
  return canPlayerControlToken(playerRole, playerId, token);
}
function canEditTokenData(token) {
  return canUseToken(token);
}
async function initializeCharacterToken(tokenId) {
  const token = getCharacterById(tokenId);
  if (!token || !isCharacterToken(token)) return false;
  const shouldInitialize = !isTrackedCharacter(token);
  if (shouldInitialize) {
    await updateTrackerData(tokenId, (current2) => current2);
  }
  await ensureOverlayForToken(tokenId);
  return shouldInitialize;
}
function resolveDefaultTargetTokenId(attackerId) {
  const visibleTargets = getCharacters().filter(
    (token) => token.id !== attackerId && token.visible !== false
  );
  const otherSelected = selectionIds.find(
    (id) => id !== attackerId && visibleTargets.some((token) => token.id === id)
  );
  if (otherSelected) return otherSelected;
  const fallback = visibleTargets[0];
  return fallback?.id ?? "";
}
async function pushDebugEntry(title, body, kind = "info") {
  const entry = {
    id: Date.now() * 1e3 + Math.floor(Math.random() * 1e3),
    title,
    body,
    kind,
    timestamp: (/* @__PURE__ */ new Date()).toLocaleTimeString()
  };
  const metadata = await lib_default.room.getMetadata();
  const nextEntries = mergeDebugEntries([entry], metadata?.[DEBUG_LOG_KEY], debugEntries);
  debugEntries = nextEntries;
  renderDebugConsole();
  await lib_default.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-entry", entry },
    { destination: "REMOTE" }
  );
  if (playerRole === "GM") {
    await lib_default.room.setMetadata({
      [DEBUG_LOG_KEY]: nextEntries
    });
  }
}
async function clearDebugConsole() {
  if (playerRole !== "GM") {
    setStatus("Only the GM can clear the debug console.", "error");
    return;
  }
  debugEntries = [];
  renderDebugConsole();
  await lib_default.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-clear" },
    { destination: "REMOTE" }
  );
  await lib_default.room.setMetadata({
    [DEBUG_LOG_KEY]: []
  });
  setStatus("Debug console cleared.", "success");
}
function renderDebugConsole() {
  if (!debugEntries.length) {
    ui.debugConsole.innerHTML = `
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}

Actions from all players and the GM will appear here after rolls and attacks.</pre>
      </div>`;
    return;
  }
  ui.debugConsole.innerHTML = debugEntries.map(
    (entry) => `
        <div class="debug-entry">
          <div class="debug-head">
            <div class="debug-title">${escapeHtml(entry.title)}</div>
            <div class="muted">${escapeHtml(entry.timestamp)}</div>
          </div>
          <pre class="console-output">${escapeHtml(entry.body)}</pre>
        </div>`
  ).join("");
}
async function loadSharedDebugConsole() {
  const metadata = await lib_default.room.getMetadata();
  debugEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
}
function arraysEqual(left, right) {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}
function startSelectionPolling() {
  if (selectionPollTimer) {
    clearInterval(selectionPollTimer);
  }
  selectionPollTimer = setInterval(() => {
    void lib_default.player.getSelection().then((selection) => selection ?? []).then((selection) => {
      if (!arraysEqual(selectionIds, selection)) {
        return syncState();
      }
      return null;
    }).catch((error) => {
      console.warn("[Body HP] Selection polling failed", error);
    });
  }, 200);
}
function formatAttackDebug({
  attackerName,
  targetName,
  targetPart,
  attackSkillName,
  attackSkillValue,
  weaponDamage,
  strengthBonus,
  attackBonuses,
  manualAttackPenalties,
  automaticTargetPenalty,
  totalAttackPenalties,
  defenseBonuses,
  defensePenalties,
  baseTargetParry,
  targetParry,
  parryMode,
  targetArmor,
  result,
  beforeHp,
  afterHp,
  beforeMinor,
  afterMinor,
  beforeSerious,
  afterSerious,
  critApplied
}) {
  const accuracyTable = formatTextTable(
    ["Side", "Attacking", "Defending"],
    [
      [
        "Accuracy",
        `${result.attackRoll} + ${attackSkillValue * 10} + ${attackBonuses} - ${totalAttackPenalties} = ${result.attackTotal}`,
        `${result.defenseRoll} + ${targetParry * 10} + ${defenseBonuses} - ${defensePenalties} = ${result.defenseTotal}`
      ],
      [
        "Damage",
        `${result.attackTotal} + ${weaponDamage}`,
        `${result.defenseTotal} + ${targetArmor}`
      ],
      [
        "Result",
        `${result.damage?.totalAttack ?? result.attackTotal}`,
        `${result.damage?.totalDefense ?? result.defenseTotal}`
      ]
    ]
  );
  const damageTable = formatTextTable(
    ["Parameter", "Value"],
    [
      ["Attacker", attackerName],
      ["Target", `${targetName} -> ${targetPart}`],
      ["Attack Skill", `${attackSkillName} (${attackSkillValue})`],
      ["Strength Bonus", strengthBonus],
      ["Manual Attack Penalty", manualAttackPenalties],
      ["Auto Target Penalty", automaticTargetPenalty],
      ["Parry Mode", getParryModeLabel(parryMode)],
      ["Base Parry", baseTargetParry],
      ["Effective Parry", targetParry],
      ["Armor", targetArmor],
      ["Outcome", result.outcome],
      ["Damage Diff", result.damage?.damageDiff ?? 0],
      ["Damage Label", result.damage?.label ?? "No damage"],
      ["Applied Min/Sir/Crit", `${result.damage?.minor ?? 0} / ${result.damage?.serious ?? 0} / ${result.damage?.crit ?? 0}`],
      ["Converted Crit", critApplied],
      ["Crit State", formatStateTransition(beforeHp, afterHp)],
      ["Minor State", formatStateTransition(beforeMinor, afterMinor)],
      ["Serious State", formatStateTransition(beforeSerious, afterSerious)]
    ]
  );
  return `${accuracyTable}

${damageTable}`;
}
function projectPartDamage(part, damage) {
  const next = {
    current: Number(part?.current) || 0,
    max: Number(part?.max) || 0,
    armor: Number(part?.armor) || 0,
    minor: Number(part?.minor) || 0,
    serious: Number(part?.serious) || 0
  };
  next.minor = Math.max(0, next.minor + (Number(damage?.minor) || 0));
  next.serious = Math.max(0, next.serious + (Number(damage?.serious) || 0));
  const promotedSerious = Math.floor(next.minor / 4);
  next.minor %= 4;
  next.serious += promotedSerious;
  const convertedCrit = Math.floor(next.serious / 2);
  next.serious %= 2;
  const directCrit = Math.max(0, Number(damage?.crit) || 0);
  const totalCrit = directCrit + convertedCrit;
  next.current = clamp(next.current - totalCrit, 0, next.max);
  return {
    ...next,
    critApplied: totalCrit
  };
}
function formatDiceDebug({ tokenName, result }) {
  return formatTextTable(
    ["Parameter", "Value"],
    [
      ["Actor", tokenName],
      ["Dice", `${result.count}d${result.sides}`],
      ["Rolls", result.rolls.join(", ")],
      ["Subtotal", result.subtotal],
      ["Modifier", result.modifier],
      ["Total", result.total]
    ]
  );
}
function getCurrentPlayerColor() {
  return String(
    playerColor || partyPlayers.find((player) => player?.id === playerId)?.color || "#facc15"
  );
}
function getAutomaticTargetPenalty(targetPart) {
  if (targetPart === "Head") return 30;
  if (targetPart === "L.Arm" || targetPart === "R.Arm" || targetPart === "L.Leg" || targetPart === "R.Leg") {
    return 15;
  }
  return 0;
}
function getParryDivisor(mode) {
  if (mode === "off") return 0;
  const numeric = Number(mode);
  if (Number.isInteger(numeric) && numeric >= 1 && numeric <= 5) {
    return numeric;
  }
  return 1;
}
function getParryModeLabel(mode) {
  if (mode === "off") return "Ignore Parry";
  const divisor = getParryDivisor(mode);
  return `${divisor} Opponent${divisor === 1 ? "" : "s"}`;
}
function formatStateTransition(before, after) {
  if (before == null || after == null) return "-";
  return `${before} -> ${after}`;
}
function formatRollCharDebug({ tokenName, attributeLabel, result }) {
  return [
    `Character: ${tokenName}`,
    `Characteristic: ${attributeLabel}`,
    `${result.result}`,
    "",
    formatTextTable(
      ["Roll", "Base Attribute", "Modifier", "Final Attribute"],
      [[result.roll, result.baseAttribute, result.modifier, result.finalAttribute]]
    )
  ].join("\n");
}
function formatRollSkillDebug({ tokenName, skillName, result }) {
  return [
    `Character: ${tokenName}`,
    `Skill: ${skillName}`,
    `${result.result}`,
    "",
    formatTextTable(
      ["Parameter", "Value"],
      [
        ["First Roll", `${result.rollPrimary} + ${result.baseSkill * 10} + ${result.modifier} = ${result.totalPrimary}`],
        ["Second Roll", `${result.rollSecondary} = ${result.totalSecondary}`]
      ]
    )
  ].join("\n");
}
function formatTextTable(headers, rows) {
  const normalizedHeaders = headers.map((cell) => String(cell ?? ""));
  const normalizedRows = rows.map((row) => row.map((cell) => String(cell ?? "")));
  const widths = normalizedHeaders.map(
    (header, columnIndex) => Math.max(
      header.length,
      ...normalizedRows.map((row) => (row[columnIndex] ?? "").length)
    )
  );
  const renderBorder = (left, middle, right, fill) => `${left}${widths.map((width) => fill.repeat(width + 2)).join(middle)}${right}`;
  const renderRow = (row) => `\u2502 ${row.map((cell, columnIndex) => String(cell ?? "").padEnd(widths[columnIndex], " ")).join(" \u2502 ")} \u2502`;
  return [
    renderBorder("\u2552", "\u2564", "\u2555", "\u2550"),
    renderRow(normalizedHeaders),
    renderBorder("\u255E", "\u256A", "\u2561", "\u2550"),
    ...normalizedRows.map(renderRow),
    renderBorder("\u2558", "\u2567", "\u255B", "\u2550")
  ].join("\n");
}
function getAttackDraft(token, data, targetCharacters) {
  const defaultWeapon = getAvailableWeapons(token, "melee")[0] ?? { damage: 0 };
  const stored = attackFormDrafts.get(token.id) ?? {};
  const combatSkillNames = getAttackSkillEntries(data.odyssey).map(([skillName]) => skillName);
  const fallbackSkill = combatSkillNames[0] ?? CORE_COMBAT_SKILLS.find((key) => key in data.odyssey.skills) ?? CORE_COMBAT_SKILLS[0];
  const targetTokenId = stored.targetTokenId === "" ? "" : targetCharacters.some((target) => target.id === stored.targetTokenId) ? stored.targetTokenId : resolveDefaultTargetTokenId(token.id);
  return {
    skill: combatSkillNames.includes(stored.skill) ? stored.skill : fallbackSkill,
    targetTokenId,
    targetPart: BODY_ORDER.includes(stored.targetPart) ? stored.targetPart : "Torso",
    weaponDamage: stored.weaponDamage ?? String(defaultWeapon.damage),
    attackBonuses: stored.attackBonuses ?? "0",
    attackPenalties: stored.attackPenalties ?? "0",
    defenseBonuses: stored.defenseBonuses ?? "0",
    defensePenalties: stored.defensePenalties ?? "0",
    manualArmor: stored.manualArmor ?? "0",
    manualParry: stored.manualParry ?? "0",
    parryMode: ["off", "1", "2", "3", "4", "5"].includes(String(stored.parryMode)) ? String(stored.parryMode) : "1"
  };
}
function saveAttackDraftValue(tokenId, field, value) {
  if (!tokenId || !field) return;
  const current2 = attackFormDrafts.get(tokenId) ?? {};
  attackFormDrafts.set(tokenId, {
    ...current2,
    [field]: value
  });
}
function buildCircleCommands(radius, segments = 28) {
  const safeRadius = Math.max(radius, 8);
  const commands = [];
  for (let index = 0; index <= segments; index += 1) {
    const angle = Math.PI * 2 * index / segments - Math.PI / 2;
    const x = Math.cos(angle) * safeRadius;
    const y = Math.sin(angle) * safeRadius;
    if (index === 0) {
      commands.push([Command.MOVE, x, y]);
    } else {
      commands.push([Command.LINE, x, y]);
    }
  }
  commands.push([Command.CLOSE]);
  return commands;
}
function isLocalTargetHighlight(item) {
  return Boolean(item?.metadata?.[TARGET_HIGHLIGHT_KEY]);
}
async function clearTargetHighlight() {
  const localItems = await lib_default.scene.local.getItems();
  const highlightIds = localItems.filter(isLocalTargetHighlight).map((item) => item.id);
  if (highlightIds.length) {
    await lib_default.scene.local.deleteItems(highlightIds);
  }
}
async function buildTargetHighlightItem(targetToken) {
  let bounds = null;
  try {
    bounds = await lib_default.scene.items.getItemBounds([targetToken.id]);
  } catch (error) {
    console.warn("[Body HP] Unable to read target bounds for highlight", error);
  }
  const width = Math.max(
    bounds?.width ?? 0,
    (targetToken.width || 140) * Math.abs(targetToken.scale?.x ?? 1),
    56
  );
  const height = Math.max(
    bounds?.height ?? 0,
    (targetToken.height || 140) * Math.abs(targetToken.scale?.y ?? 1),
    56
  );
  const diameter = Math.max(24, Math.min(width, height));
  const radius = diameter / 2;
  const center = bounds?.center ?? targetToken.position;
  return buildPath().name(`Attack Target: ${getCharacterName(targetToken)}`).commands(buildCircleCommands(radius)).position(center).rotation(0).attachedTo(targetToken.id).disableAttachmentBehavior(["ROTATION"]).layer("ATTACHMENT").locked(true).disableHit(true).fillColor(getCurrentPlayerColor()).fillOpacity(0.22).strokeColor(getCurrentPlayerColor()).strokeOpacity(0).strokeWidth(0).metadata({
    [TARGET_HIGHLIGHT_KEY]: {
      playerId,
      playerName,
      attackerTokenId: activeTokenId ?? "",
      targetTokenId: targetToken.id
    }
  }).build();
}
async function syncTargetHighlight() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker || !isCharacterToken(attacker)) {
    await clearTargetHighlight();
    return;
  }
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== attacker.id && item.visible !== false
  );
  const draft2 = getAttackDraft(attacker, getTrackerData(attacker), targetCharacters);
  const target = getCharacterById(draft2.targetTokenId);
  const localItems = await lib_default.scene.local.getItems();
  const existingHighlights = localItems.filter(isLocalTargetHighlight);
  if (!target || !isCharacterToken(target) || target.visible === false || target.id === attacker.id) {
    if (existingHighlights.length) {
      await lib_default.scene.local.deleteItems(existingHighlights.map((item) => item.id));
    }
    return;
  }
  if (existingHighlights.length === 1 && existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.targetTokenId === target.id && existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.attackerTokenId === attacker.id) {
    return;
  }
  if (existingHighlights.length) {
    await lib_default.scene.local.deleteItems(existingHighlights.map((item) => item.id));
  }
  await lib_default.scene.local.addItems([await buildTargetHighlightItem(target)]);
}
async function teardownTargetPickerTool() {
  if (!targetPickState.toolReady) return;
  try {
    await lib_default.tool.removeMode(TARGET_PICK_MODE_ID);
  } catch (error) {
    console.warn("[Body HP] Unable to remove target picker mode", error);
  }
  try {
    await lib_default.tool.remove(TARGET_PICK_TOOL_ID);
  } catch (error) {
    console.warn("[Body HP] Unable to remove target picker tool", error);
  }
  targetPickState.toolReady = false;
}
async function restorePreviousTool() {
  const previousToolId = targetPickState.previousToolId;
  const previousModeId = targetPickState.previousModeId;
  if (!previousToolId || previousToolId === TARGET_PICK_TOOL_ID) return;
  targetPickState.restoring = true;
  try {
    await lib_default.tool.activateTool(previousToolId);
    if (previousModeId) {
      try {
        await lib_default.tool.activateMode(previousToolId, previousModeId);
      } catch (error) {
        console.warn("[Body HP] Unable to restore previous tool mode", error);
      }
    }
  } finally {
    targetPickState.restoring = false;
  }
}
async function stopTargetPick(statusMessage = "", statusKind = "info") {
  const wasActive = targetPickState.active;
  targetPickState.active = false;
  targetPickState.attackerTokenId = null;
  render();
  if (wasActive) {
    await restorePreviousTool();
  }
  await teardownTargetPickerTool();
  targetPickState.previousToolId = "";
  targetPickState.previousModeId = void 0;
  if (statusMessage) {
    setStatus(statusMessage, statusKind);
  }
}
async function ensureTargetPickerTool() {
  if (targetPickState.toolReady) return;
  await lib_default.tool.create({
    id: TARGET_PICK_TOOL_ID,
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick Attack Target" }],
    defaultMode: TARGET_PICK_MODE_ID
  });
  await lib_default.tool.createMode({
    id: TARGET_PICK_MODE_ID,
    icons: [{ icon: EXTENSION_ICON_URL, label: "Pick Attack Target" }],
    cursors: [{ cursor: TARGET_PICK_CURSOR }],
    onToolClick: async (_context, event) => {
      if (!targetPickState.active) return false;
      const attacker = getCharacterById(targetPickState.attackerTokenId);
      const clickedTargetId = event.target?.id ?? "";
      const liveItems = clickedTargetId ? await lib_default.scene.items.getItems() : [];
      const target = (clickedTargetId ? liveItems.find((item) => item.id === clickedTargetId) : null) ?? event.target;
      if (!attacker || !isCharacterToken(attacker)) {
        await stopTargetPick("Select an attacker token first.", "error");
        return false;
      }
      if (!target || !isCharacterToken(target)) {
        setStatus("Click a visible character token to use it as target.", "error");
        return false;
      }
      if (target.visible === false) {
        setStatus("Hidden tokens cannot be targeted.", "error");
        return false;
      }
      if (target.id === attacker.id) {
        setStatus("Attacker and target must be different tokens.", "error");
        return false;
      }
      saveAttackDraftValue(attacker.id, "targetTokenId", target.id);
      activeTokenId = attacker.id;
      render();
      const targetField = ui.selectedTokenPanel.querySelector('[data-attack-field="targetTokenId"]');
      if (targetField instanceof HTMLSelectElement) {
        targetField.value = target.id;
      }
      await syncTargetHighlight();
      await stopTargetPick(`Target set to ${getCharacterName(target)}.`, "success");
      return false;
    },
    onKeyDown: (_context, event) => {
      if (event.key === "Escape" && targetPickState.active) {
        void stopTargetPick("Target picking cancelled.", "info");
      }
    },
    onDeactivate: () => {
      if (targetPickState.active && !targetPickState.restoring) {
        void stopTargetPick("Target picking cancelled.", "info");
      }
    }
  });
  targetPickState.toolReady = true;
}
async function startTargetPick() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== attacker.id && item.visible !== false
  );
  if (!targetCharacters.length) {
    setStatus("Add at least one visible target token.", "error");
    return;
  }
  if (targetPickState.active && targetPickState.attackerTokenId === attacker.id) {
    await stopTargetPick("Target picking cancelled.", "info");
    return;
  }
  if (targetPickState.active) {
    await stopTargetPick();
  }
  targetPickState.previousToolId = await lib_default.tool.getActiveTool();
  targetPickState.previousModeId = await lib_default.tool.getActiveToolMode();
  targetPickState.attackerTokenId = attacker.id;
  targetPickState.active = true;
  await ensureTargetPickerTool();
  await lib_default.tool.activateTool(TARGET_PICK_TOOL_ID);
  await lib_default.tool.activateMode(TARGET_PICK_TOOL_ID, TARGET_PICK_MODE_ID);
  render();
  setStatus("Click a visible target token on the map to assign it.", "info");
}
function renderCollapsibleSection(title, content, open = false, sectionKey = "") {
  const scopedSectionKey = `${activeTokenId ?? "global"}:${sectionKey || title}`;
  const resolvedOpen = collapsibleSectionState.has(scopedSectionKey) ? collapsibleSectionState.get(scopedSectionKey) : open;
  return `
    <details class="collapsible-block" data-section-key="${escapeHtml(scopedSectionKey)}" ${resolvedOpen ? "open" : ""}>
      <summary class="collapsible-title">${escapeHtml(title)}</summary>
      <div class="collapsible-body">${content}</div>
    </details>
  `;
}
function pushPrivateGmEntry(title, body) {
  gmPrivateEntries = [
    {
      id: Date.now(),
      title,
      body,
      timestamp: (/* @__PURE__ */ new Date()).toLocaleTimeString()
    },
    ...gmPrivateEntries
  ].slice(0, 12);
}
function rollCharacterCheck(attributeValue, modifier = 0) {
  const baseAttribute = Number(attributeValue) || 0;
  const finalAttribute = Math.max(0, baseAttribute + (Number(modifier) || 0));
  const roll = Math.floor(Math.random() * 20) + 1;
  const result = roll <= finalAttribute ? "Check Passed" : "Check Failed";
  return {
    roll,
    baseAttribute,
    modifier: Number(modifier) || 0,
    finalAttribute,
    result
  };
}
function rollSkillCheck(skillValue, modifier = 0) {
  const baseSkill = Number(skillValue) || 0;
  const rollPrimary = Math.floor(Math.random() * 100) + 1;
  const rollSecondary = Math.floor(Math.random() * 100) + 1;
  const totalPrimary = rollPrimary + baseSkill * 10 + (Number(modifier) || 0);
  const totalSecondary = rollSecondary;
  const result = totalPrimary > totalSecondary ? "Check Passed" : "Check Failed";
  return {
    rollPrimary,
    rollSecondary,
    baseSkill,
    modifier: Number(modifier) || 0,
    totalPrimary,
    totalSecondary,
    result
  };
}
function renderOwnerFields(data, disabledAttr) {
  const playerOptions = [
    `<option value="">Unassigned</option>`,
    ...getSortedPartyPlayers().map(
      (player) => `
        <option value="${escapeHtml(player.id)}" ${data.odyssey.owner.playerId === player.id ? "selected" : ""}>${escapeHtml(player.name || player.id)}</option>`
    )
  ].join("");
  return renderCollapsibleSection(
    "Ownership",
    `
      <div class="hint-box">
        <div class="field-label">Current viewer</div>
        <pre class="console-output">Name: ${escapeHtml(playerName || "Unknown")}
Player ID: ${escapeHtml(playerId || "Unavailable")}</pre>
      </div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Assigned Player</span>
          <select data-action="select-owner-player" ${disabledAttr}>${playerOptions}</select>
        </label>
      </div>
    `,
    false
  );
}
function renderCharacteristicsBlock(data, disabledAttr) {
  const attributeInputs = ATTRIBUTE_UI_FIELDS.map(
    ([key, label]) => `
      <label class="field-stack">
        <span class="field-label">${escapeHtml(label)}</span>
        <input type="number" min="0" max="15" value="${data.odyssey.attributes[key] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${escapeHtml(key)}" ${disabledAttr}>
      </label>`
  ).join("");
  return renderCollapsibleSection(
    "Characteristics",
    `<div class="form-grid">${attributeInputs}</div>`,
    false
  );
}
function renderOdysseySkillRows(odyssey, skillEntries, disabledAttr) {
  return skillEntries.map(
    ([key, value]) => `
        <div class="skill-row">
          <div class="skill-name">${escapeHtml(key)}</div>
          <input type="number" min="0" max="10" value="${value}" data-action="set-odyssey-skill" data-skill="${escapeHtml(key)}" ${disabledAttr}>
          <label class="skill-toggle">
            <input type="checkbox" data-action="set-skill-strength-bonus" data-skill="${escapeHtml(key)}" ${disabledAttr} ${getSkillStrengthBonusFlag(odyssey, key) ? "checked" : ""}>
            <span>STR Bonus</span>
          </label>
          <button type="button" class="danger" data-action="remove-skill" data-skill="${escapeHtml(key)}" ${CORE_COMBAT_SKILLS.includes(key) ? "disabled" : disabledAttr}>${CORE_COMBAT_SKILLS.includes(key) ? "Core" : "Remove"}</button>
        </div>`
  ).join("");
}
function renderEnglishSkillsBlock(data, disabledAttr) {
  const combatSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getCombatSkillEntries(data.odyssey),
    disabledAttr
  );
  const appliedSkillRows = renderOdysseySkillRows(
    data.odyssey,
    getAppliedSkillEntries(data.odyssey),
    disabledAttr
  );
  return renderCollapsibleSection(
    "Skills",
    `
      <div class="field-label">Combat</div>
      <div class="list">${combatSkillRows || '<div class="empty">No combat skills yet.</div>'}</div>
      <div class="field-label">Applied</div>
      <div class="list">${appliedSkillRows || '<div class="empty">No applied skills yet.</div>'}</div>
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Skill Name</span>
          <input type="text" data-skill-field="new-name" placeholder="New skill" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Value</span>
          <input type="number" min="0" max="10" value="0" data-skill-field="new-value" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Category</span>
          <select data-skill-field="new-category" ${disabledAttr}>
            <option value="${COMBAT_SKILL_CATEGORY}">Combat</option>
            <option value="${APPLIED_SKILL_CATEGORY}" selected>Applied</option>
          </select>
        </label>
        <label class="field-stack checkbox-stack">
          <span class="field-label">Add Strength Bonus?</span>
          <label class="skill-toggle">
            <input type="checkbox" data-skill-field="new-strength-bonus" ${disabledAttr}>
            <span>Enable for attack damage</span>
          </label>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" class="secondary" data-action="add-skill" ${disabledAttr}>Add Skill</button>
      </div>
    `,
    false
  );
}
function renderEnglishAttackBlock(token, data, tokenLocked) {
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false
  );
  const disabledAttr = tokenLocked ? "disabled" : "";
  const pickDisabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft2 = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft2.skill);
  const selectedTarget = targetCharacters.find((target) => target.id === draft2.targetTokenId) ?? null;
  const targetName = selectedTarget ? getCharacterName(selectedTarget) : "Manual Defense";
  const isPickingTarget = targetPickState.active && targetPickState.attackerTokenId === token.id;
  const manualDefenseDisabledAttr = tokenLocked || selectedTarget ? "disabled" : "";
  return renderCollapsibleSection(
    "Attack",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Target Token</span>
          <select data-attack-field="targetTokenId" ${disabledAttr}>
            <option value="" ${draft2.targetTokenId ? "" : "selected"}>No Target (Manual Defense)</option>
            ${targetCharacters.map(
      (target) => `<option value="${target.id}" ${target.id === draft2.targetTokenId ? "selected" : ""}>${escapeHtml(
        getCharacterName(target)
      )}</option>`
    ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Pick On Map</span>
          <button type="button" class="secondary" data-action="pick-attack-target" ${pickDisabledAttr}>
            ${isPickingTarget ? "Cancel Target Pick" : "Pick Target On Map"}
          </button>
        </label>
        <label class="field-stack">
          <span class="field-label">Target Body Part</span>
          <select data-attack-field="targetPart" ${disabledAttr}>
            ${BODY_ORDER.map(
      (part) => `<option value="${part}" ${part === draft2.targetPart ? "selected" : ""}>${part}</option>`
    ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon Damage</span>
          <input type="number" value="${draft2.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Bonus</span>
          <input type="number" value="${draft2.attackBonuses}" data-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Penalty</span>
          <input type="number" value="${draft2.attackPenalties}" data-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense Bonus</span>
          <input type="number" value="${draft2.defenseBonuses}" data-attack-field="defenseBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Defense Penalty</span>
          <input type="number" value="${draft2.defensePenalties}" data-attack-field="defensePenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Parry Mode</span>
          <select data-attack-field="parryMode" ${disabledAttr}>
            <option value="off" ${draft2.parryMode === "off" ? "selected" : ""}>Do Not Count Parry</option>
            ${[1, 2, 3, 4, 5].map(
      (value) => `<option value="${value}" ${String(value) === draft2.parryMode ? "selected" : ""}>${value} Opponent${value === 1 ? "" : "s"}</option>`
    ).join("")}
          </select>
        </label>
        <label class="field-stack">
          <span class="field-label">Manual Armor</span>
          <input type="number" min="0" max="99" value="${draft2.manualArmor}" data-attack-field="manualArmor" ${manualDefenseDisabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Manual Parry</span>
          <input type="number" min="0" max="10" value="${draft2.manualParry}" data-attack-field="manualParry" ${manualDefenseDisabledAttr}>
        </label>
      </div>
      <div class="muted">${targetCharacters.length ? "Attack goes from the selected attacker token to the selected target token, or you can switch to manual defense." : "No visible target tokens found. Attack can still be rolled with manual defense values."}</div>
      <div class="muted">Current target: ${escapeHtml(targetName)}</div>
      <div class="muted">Automatic called-shot penalties: Head -30, arms/legs -15.</div>
      <div class="muted">Strength is added to weapon damage only for attack skills with STR Bonus enabled. ${escapeHtml(PARRY_SKILL_NAME)} is added to defense only for melee attacks. Manual Armor/Parry are used when no target token is selected.</div>
      <div class="row row-gap">
        <button type="button" class="success" data-action="perform-attack" ${disabledAttr}>Attack</button>
      </div>
    `,
    true
  );
}
function renderEnglishDiceBlock(token, data, tokenLocked) {
  const attributeOptions = ATTRIBUTE_UI_FIELDS.map(
    ([key, label]) => `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
  ).join("");
  const skillOptions = buildGroupedSkillOptions(data.odyssey);
  const tokenLockedAttr = tokenLocked ? "disabled" : "";
  return renderCollapsibleSection(
    "Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice Sides</span>
          <input type="number" min="2" max="1000" value="20" data-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Dice Count</span>
          <input type="number" min="1" max="100" value="1" data-roll-field="count">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-dice">Roll Dice</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Characteristic</span>
          <select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Roll Characteristic</button>
      </div>

      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Skill</span>
          <select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Bonus / Penalty</span>
          <input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}>
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Roll Skill</button>
      </div>
    `,
    false
  );
}
function renderPrivateGmDiceBlock() {
  if (!isEditable()) return "";
  const privateLog = gmPrivateEntries.length ? gmPrivateEntries.map(
    (entry) => `
            <div class="debug-entry">
              <div class="debug-head">
                <div class="debug-title">${escapeHtml(entry.title)}</div>
                <div class="muted">${escapeHtml(entry.timestamp)}</div>
              </div>
              <pre class="console-output">${escapeHtml(entry.body)}</pre>
            </div>`
  ).join("") : '<div class="empty">Private GM rolls will stay visible only here.</div>';
  return renderCollapsibleSection(
    "GM Private Dice",
    `
      <div class="form-grid">
        <label class="field-stack">
          <span class="field-label">Dice sides</span>
          <input type="number" min="2" max="1000" value="20" data-gm-roll-field="dice">
        </label>
        <label class="field-stack">
          <span class="field-label">Dice Count</span>
          <input type="number" min="1" max="100" value="1" data-gm-roll-field="count">
        </label>
        <label class="field-stack">
          <span class="field-label">Modifier</span>
          <input type="number" value="0" data-gm-roll-field="modifier">
        </label>
      </div>
      <div class="row row-gap">
        <button type="button" data-action="perform-gm-private-roll">GM Roll</button>
      </div>
      <div class="list private-roll-log">${privateLog}</div>
    `,
    false
  );
}
function renderSelectedToken() {
  const panelState = captureSelectedPanelState();
  activeTokenId = resolveActiveTokenId();
  const token = getCharacterById(activeTokenId);
  if (!token) {
    ui.selectionHint.textContent = "No character token selected";
    ui.selectedTokenPanel.innerHTML = '<div class="empty">Add a character token to the map from Owlbear Rodeo Characters, then select it.</div>';
    return;
  }
  const tracked = isTrackedCharacter(token);
  const data = getTrackerData(token);
  const odyssey = getOdysseyData(token);
  const totals = getBodyTotals(data);
  const selected = selectionIds.includes(token.id);
  const tokenLocked = !canUseToken(token);
  const bodyFieldDisabled = !canEditTokenData(token) ? "disabled" : "";
  const gmOnlyDisabled = !isEditable() ? "disabled" : "";
  const showPartBlock = isEditable();
  const lastRollText = data.lastRoll ? escapeHtml(data.lastRoll.summary || "Last roll recorded") : "No rolls synced yet";
  ui.selectionHint.textContent = selected ? "Selected on map" : "Showing current focus";
  ui.selectedTokenPanel.innerHTML = `
    <div class="selected-card">
      <div class="selected-head">
        <div>
          <div class="token-name">${escapeHtml(getCharacterName(token))}</div>
          <div class="token-meta">${escapeHtml(token.id.slice(0, 8))} - ${tracked ? "Initialized" : "Auto-init on selection"} - ${tokenLocked ? "Read only" : "Controllable"}</div>
        </div>
        <div class="row row-gap">
          <button type="button" data-action="focus-token" class="secondary">Select On Map</button>
        </div>
      </div>

      <div class="summary-strip">
        <div class="stat-chip">
          <span class="chip-label">Body HP</span>
          <span class="chip-value">${totals.current}/${totals.max}</span>
        </div>
        <div class="stat-chip">
          <span class="chip-label">Assigned Player</span>
          <span class="chip-value">${escapeHtml(odyssey.owner.playerName || odyssey.owner.playerId || "Unassigned")}</span>
        </div>
      </div>

      ${isEditable() ? `
            ${renderOwnerFields({ odyssey }, gmOnlyDisabled)}
            ${renderCharacteristicsBlock({ odyssey }, gmOnlyDisabled)}
            ${renderEnglishSkillsBlock({ odyssey }, gmOnlyDisabled)}
          ` : ""}
      ${renderEnglishAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishDiceBlock(token, { odyssey }, tokenLocked)}
      ${renderPrivateGmDiceBlock()}
      ${renderCollapsibleSection(
    "Last Roll",
    `<pre class="console-output">${lastRollText}</pre>`,
    false
  )}
      ${showPartBlock ? renderCollapsibleSection(
    "Body Parts",
    `
                <div class="body-table-wrap">
                  <table class="body-table">
                    <thead>
                      <tr>
                        <th>Body Part</th>
                        <th>Crit</th>
                        <th>Max</th>
                        <th>Armor</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${BODY_ORDER.map((partName) => {
      const part = data.body[partName];
      return `
                          <tr>
                            <td class="part-name">${escapeHtml(partName)}</td>
                            <td>
                              <div class="inline-stepper">
                                <button type="button" data-action="change-part" data-part="${escapeHtml(
        partName
      )}" data-field="current" data-delta="-1" ${bodyFieldDisabled}>-</button>
                                <input type="number" min="0" max="${part.max}" value="${part.current}" data-action="set-field" data-part="${escapeHtml(
        partName
      )}" data-field="current" ${bodyFieldDisabled}>
                                <button type="button" data-action="change-part" data-part="${escapeHtml(
        partName
      )}" data-field="current" data-delta="1" ${bodyFieldDisabled}>+</button>
                              </div>
                            </td>
                            <td>
                              <input class="compact-input" type="number" min="0" max="99" value="${part.max}" data-action="set-field" data-part="${escapeHtml(
        partName
      )}" data-field="max" ${bodyFieldDisabled}>
                            </td>
                            <td>
                              <input class="compact-input" type="number" min="0" max="99" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(
        partName
      )}" data-field="armor" ${bodyFieldDisabled}>
                            </td>
                          </tr>
                        `;
    }).join("")}
                    </tbody>
                  </table>
                </div>
              `,
    true
  ) : ""}
      ${renderCollapsibleSection(
    "Overlay Preview",
    `<pre class="console-output">${escapeHtml(formatOverlayText(data))}</pre>`,
    false
  )}
    </div>`;
  restoreSelectedPanelState(panelState);
}
function renderTrackedList() {
  const trackedCharacters = getTrackedCharacters();
  ui.trackedCount.textContent = String(trackedCharacters.length);
  if (!trackedCharacters.length) {
    ui.trackedList.innerHTML = '<div class="empty">No initialized characters yet. Click a token on the map to initialize it automatically.</div>';
    return;
  }
  ui.trackedList.innerHTML = trackedCharacters.map((token) => {
    const data = getTrackerData(token);
    const totals = getBodyTotals(data);
    const controllable = canUseToken(token);
    return `
        <button type="button" class="list-item${token.id === activeTokenId ? " active" : ""}" data-action="select-character" data-token-id="${token.id}">
          <div class="list-item-head">
            <span>${escapeHtml(getCharacterName(token))}</span>
            <span class="pill hp">${totals.current}/${totals.max}</span>
          </div>
          <div class="list-item-sub">${controllable ? "Playable" : "Read only"}</div>
        </button>`;
  }).join("");
}
function renderAllCharacters() {
  const characters = getCharacters();
  ui.allCount.textContent = String(characters.length);
  if (!characters.length) {
    ui.allTokensList.innerHTML = '<div class="empty">No character tokens are on the scene yet.</div>';
    return;
  }
  ui.allTokensList.innerHTML = characters.map((token) => {
    const tracked = isTrackedCharacter(token);
    const controllable = canUseToken(token);
    return `
        <div class="token-row${token.id === activeTokenId ? " active" : ""}">
          <div>
            <div class="token-row-name">${escapeHtml(getCharacterName(token))}</div>
            <div class="token-row-sub">${escapeHtml(token.id.slice(0, 8))} - ${controllable ? "Playable" : "Read only"}</div>
          </div>
          <div class="row row-gap">
            <button type="button" class="secondary" data-action="select-character" data-token-id="${token.id}">Select</button>
            <span class="pill ${tracked ? "hp" : "armor"}">${tracked ? "Initialized" : "Ready"}</span>
          </div>
        </div>`;
  }).join("");
}
function render() {
  ui.roleBadge.textContent = playerRole === "GM" ? "GM" : "PLAYER";
  ui.trackedSection.classList.toggle("hidden", playerRole !== "GM");
  ui.allTokensSection.classList.toggle("hidden", playerRole !== "GM");
  ui.clearDebugBtn?.classList.toggle("hidden", playerRole !== "GM");
  renderSelectedToken();
  renderDebugConsole();
  if (playerRole === "GM") {
    renderTrackedList();
    renderAllCharacters();
  }
}
async function syncState(showToast = false) {
  const [role, id, name, color, items, selection, players] = await Promise.all([
    lib_default.player.getRole(),
    lib_default.player.getId(),
    lib_default.player.getName(),
    lib_default.player.getColor(),
    lib_default.scene.items.getItems(),
    lib_default.player.getSelection(),
    lib_default.party.getPlayers()
  ]);
  playerRole = role;
  playerId = id;
  playerName = name;
  playerColor = color;
  partyPlayers = players ?? [];
  sceneItems = items;
  selectionIds = selection ?? [];
  const selectedCharacterId = selectionIds.find(
    (selectionId) => sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
  );
  if (selectedCharacterId) {
    activeTokenId = selectedCharacterId;
    const initialized = await initializeCharacterToken(selectedCharacterId);
    if (initialized) {
      sceneItems = await lib_default.scene.items.getItems();
    }
  } else if (activeTokenId && !sceneItems.some((item) => item.id === activeTokenId)) {
    activeTokenId = null;
  }
  render();
  await syncTargetHighlight();
  if (showToast) {
    setStatus(
      `Loaded ${getCharacters().length} character token(s), ${getTrackedCharacters().length} tracked.`,
      "success"
    );
  }
}
async function selectCharacter(tokenId) {
  activeTokenId = tokenId;
  await lib_default.player.select([tokenId], true);
  await initializeCharacterToken(tokenId);
  sceneItems = await lib_default.scene.items.getItems();
  render();
  await syncTargetHighlight();
}
async function changeBodyField(partName, field, delta) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    const part = next.body[partName];
    if (!part) return next;
    if (field === "current") {
      part.current = clamp(part.current + delta, 0, part.max);
    } else if (field === "max") {
      part.max = clamp(part.max + delta, 0, 99);
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      part.armor = clamp(part.armor + delta, 0, 99);
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}
async function setBodyField(partName, field, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    const part = next.body[partName];
    if (!part) return next;
    const numericValue = clamp(Number(value) || 0, 0, 99);
    if (field === "current") {
      part.current = clamp(numericValue, 0, part.max);
    } else if (field === "max") {
      part.max = numericValue;
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      part.armor = numericValue;
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}
async function setOwnerPlayer(ownerPlayerId) {
  if (!isEditable()) {
    setStatus("Only the GM can assign token owners.", "error");
    return;
  }
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  const selectedPlayer = getSortedPartyPlayers().find((player) => player.id === ownerPlayerId);
  await updateTrackerData(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    next.odyssey ?? (next.odyssey = structuredClone(getTrackerData(token).odyssey));
    (_a = next.odyssey).owner ?? (_a.owner = { playerId: "", playerName: "" });
    next.odyssey.owner.playerId = selectedPlayer?.id ?? "";
    next.odyssey.owner.playerName = selectedPlayer?.name ?? "";
    return next;
  });
  await ensureOverlayForToken(token.id);
  await syncState();
}
async function setOdysseySkill(skill, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey skills.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    next.odyssey.skills[skill] = clamp(Number(value) || 0, 0, 10);
    return next;
  });
  await syncState();
}
async function setOdysseySkillStrengthBonus(skill, enabled) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit Odyssey skills.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey).skillStrengthBonuses ?? (_a.skillStrengthBonuses = {});
    next.odyssey.skillStrengthBonuses[skill] = Boolean(enabled);
    return next;
  });
  await syncState();
}
async function setOdysseyAttribute(attribute, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit characteristics.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    next.odyssey.attributes[attribute] = clamp(Number(value) || 0, 0, 15);
    return next;
  });
  await syncState();
}
async function setWeaponDamage(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].damage = clamp(Number(value) || 0, -99, 99);
    return next;
  });
  await syncState();
}
async function setWeaponName(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can edit this token.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0 };
    }
    next.odyssey.weapons.melee[index].name = String(value || "").trim() || "Default";
    return next;
  });
  await syncState();
}
async function autosaveDraftField(draft2) {
  const token = getCharacterById(draft2.tokenId);
  if (!token) return;
  await updateTrackerData(token.id, (current2) => {
    var _a, _b;
    const next = structuredClone(current2);
    if (draft2.action === "set-odyssey-skill") {
      if (!isEditable()) return next;
      next.odyssey.skills[draft2.skill] = clamp(Number(draft2.value) || 0, 0, 10);
      return next;
    }
    if (draft2.action === "set-odyssey-attribute") {
      if (!isEditable()) return next;
      next.odyssey.attributes[draft2.attribute] = clamp(Number(draft2.value) || 0, 0, 15);
      return next;
    }
    if (draft2.action === "set-weapon-damage") {
      (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
      if (!next.odyssey.weapons.melee[draft2.weaponIndex]) {
        next.odyssey.weapons.melee[draft2.weaponIndex] = { name: "Default", damage: 0 };
      }
      next.odyssey.weapons.melee[draft2.weaponIndex].damage = clamp(Number(draft2.value) || 0, -99, 99);
      return next;
    }
    if (draft2.action === "set-weapon-name") {
      (_b = next.odyssey.weapons).melee ?? (_b.melee = []);
      if (!next.odyssey.weapons.melee[draft2.weaponIndex]) {
        next.odyssey.weapons.melee[draft2.weaponIndex] = { name: "Default", damage: 0 };
      }
      next.odyssey.weapons.melee[draft2.weaponIndex].name = String(draft2.value || "").trim() || "Default";
      return next;
    }
    if (draft2.action === "set-field") {
      if (!canEditTokenData(token)) return next;
      const part = next.body[draft2.partName];
      if (!part) return next;
      const numericValue = clamp(Number(draft2.value) || 0, 0, 99);
      if (draft2.field === "current") {
        part.current = clamp(numericValue, 0, part.max);
      } else if (draft2.field === "max") {
        part.max = numericValue;
        part.current = clamp(part.current, 0, part.max);
      } else if (draft2.field === "armor") {
        part.armor = numericValue;
      }
      return next;
    }
    return next;
  });
  if (draft2.action === "set-field") {
    await ensureOverlayForToken(token.id);
  }
}
function queueInputAutosave(draft2) {
  const key = [
    draft2.tokenId,
    draft2.action,
    draft2.field ?? "",
    draft2.skill ?? "",
    draft2.attribute ?? "",
    draft2.weaponIndex ?? "",
    draft2.partName ?? ""
  ].join("|");
  const existing = inputAutosaveTimers.get(key);
  if (existing) {
    clearTimeout(existing);
  }
  const timeoutId = setTimeout(() => {
    inputAutosaveTimers.delete(key);
    void autosaveDraftField(draft2).catch((error) => {
      console.warn("[Body HP] Autosave failed", error);
    });
  }, 250);
  inputAutosaveTimers.set(key, timeoutId);
}
function getActionFieldValue(selector) {
  const tokenPanel = ui.selectedTokenPanel;
  const field = tokenPanel.querySelector(selector);
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }
  return field.value;
}
async function performAttack() {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }
  const targetTokenId = getActionFieldValue('[data-attack-field="targetTokenId"]') || "";
  const target = targetTokenId ? getCharacterById(targetTokenId) : null;
  if (targetTokenId && !target) {
    setStatus("Choose a valid target token.", "error");
    return;
  }
  if (target?.visible === false) {
    setStatus("Hidden tokens cannot be targeted.", "error");
    return;
  }
  if (target && target.id === attacker.id) {
    setStatus("Attacker and target must be different tokens.", "error");
    return;
  }
  const attackerData = getTrackerData(attacker);
  const attackerOdyssey = getOdysseyData(attacker);
  const targetData = target ? getTrackerData(target) : null;
  const targetOdyssey = target ? getOdysseyData(target) : null;
  const skillName = getActionFieldValue('[data-attack-field="skill"]');
  const targetPart = getActionFieldValue('[data-attack-field="targetPart"]');
  const weaponDamage = Number(getActionFieldValue('[data-attack-field="weaponDamage"]')) || 0;
  const attackBonuses = Number(getActionFieldValue('[data-attack-field="attackBonuses"]')) || 0;
  const manualAttackPenalties = Number(getActionFieldValue('[data-attack-field="attackPenalties"]')) || 0;
  const automaticTargetPenalty = getAutomaticTargetPenalty(targetPart);
  const totalAttackPenalties = manualAttackPenalties + automaticTargetPenalty;
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  const manualArmor = clamp(Number(getActionFieldValue('[data-attack-field="manualArmor"]')) || 0, 0, 99);
  const manualParry = clamp(Number(getActionFieldValue('[data-attack-field="manualParry"]')) || 0, 0, 10);
  const parryMode = getActionFieldValue('[data-attack-field="parryMode"]') || "1";
  const parryDivisor = getParryDivisor(parryMode);
  saveAttackDraftValue(attacker.id, "skill", skillName);
  saveAttackDraftValue(attacker.id, "targetTokenId", targetTokenId);
  saveAttackDraftValue(attacker.id, "targetPart", targetPart);
  saveAttackDraftValue(attacker.id, "weaponDamage", getActionFieldValue('[data-attack-field="weaponDamage"]'));
  saveAttackDraftValue(attacker.id, "attackBonuses", getActionFieldValue('[data-attack-field="attackBonuses"]'));
  saveAttackDraftValue(attacker.id, "attackPenalties", getActionFieldValue('[data-attack-field="attackPenalties"]'));
  saveAttackDraftValue(attacker.id, "defenseBonuses", getActionFieldValue('[data-attack-field="defenseBonuses"]'));
  saveAttackDraftValue(attacker.id, "defensePenalties", getActionFieldValue('[data-attack-field="defensePenalties"]'));
  saveAttackDraftValue(attacker.id, "manualArmor", getActionFieldValue('[data-attack-field="manualArmor"]'));
  saveAttackDraftValue(attacker.id, "manualParry", getActionFieldValue('[data-attack-field="manualParry"]'));
  saveAttackDraftValue(attacker.id, "parryMode", parryMode);
  const targetArmor = target ? targetData?.body?.[targetPart]?.armor ?? 0 : manualArmor;
  const targetPartState = targetData?.body?.[targetPart] ?? { current: 0, max: 0, armor: 0, minor: 0, serious: 0 };
  const beforeHp = target ? targetPartState.current ?? 0 : null;
  const beforeMinor = target ? targetPartState.minor ?? 0 : null;
  const beforeSerious = target ? targetPartState.serious ?? 0 : null;
  const strengthBonus = getSkillStrengthBonusFlag(attackerOdyssey, skillName) ? Math.max((attackerOdyssey.attributes.Strength ?? 0) - 10, 0) : 0;
  const finalWeaponDamage = weaponDamage + strengthBonus;
  const baseTargetParry = skillName === MELEE_SKILL_NAME ? target ? targetOdyssey?.skills?.[PARRY_SKILL_NAME] ?? 0 : manualParry : 0;
  const targetParry = parryDivisor <= 0 ? 0 : Math.max(Math.floor(baseTargetParry / parryDivisor), 0);
  const result = resolveAttack({
    attackSkill: attackerOdyssey.skills[skillName] ?? 0,
    weaponDamage: finalWeaponDamage,
    defenseBonuses,
    defensePenalties,
    attackBonuses,
    attackPenalties: totalAttackPenalties,
    parry: targetParry,
    targetPart,
    targetArmor
  });
  const projectedPartState = target && result.hit && result.damage ? projectPartDamage(targetPartState, result.damage) : {
    ...target ? targetPartState : {},
    critApplied: 0
  };
  const afterHp = target ? projectedPartState.current ?? beforeHp : null;
  const afterMinor = target ? projectedPartState.minor ?? beforeMinor : null;
  const afterSerious = target ? projectedPartState.serious ?? beforeSerious : null;
  const resolvedTargetName = target ? getCharacterName(target) : "Manual Defense";
  await updateTrackerData(attacker.id, (current2) => {
    const next = structuredClone(current2);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `${getCharacterName(attacker)} -> ${resolvedTargetName}: ${result.summary}`,
      outcome: result.outcome,
      total: result.attackTotal,
      targetPart: result.targetPart,
      timestamp: (/* @__PURE__ */ new Date()).toISOString(),
      source: "owlbear-extension"
    };
    next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
    return next;
  });
  if (target) {
    await updateTrackerData(target.id, (current2) => {
      const next = structuredClone(current2);
      if (result.hit && next.body[result.targetPart]) {
        next.body[result.targetPart].current = projectedPartState.current;
        next.body[result.targetPart].minor = projectedPartState.minor;
        next.body[result.targetPart].serious = projectedPartState.serious;
      }
      next.lastRoll = {
        eventId: 0,
        actorName: getCharacterName(attacker),
        summary: result.summary,
        outcome: result.outcome,
        total: result.attackTotal,
        targetPart: result.targetPart,
        timestamp: (/* @__PURE__ */ new Date()).toISOString(),
        source: "owlbear-extension"
      };
      next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
      return next;
    });
  }
  await ensureOverlayForToken(attacker.id);
  if (target) {
    await ensureOverlayForToken(target.id);
  }
  await pushDebugEntry(
    `${getCharacterName(attacker)} attacks ${resolvedTargetName}`,
    formatAttackDebug({
      attackerName: getCharacterName(attacker),
      targetName: resolvedTargetName,
      targetPart,
      attackSkillName: skillName,
      attackSkillValue: attackerOdyssey.skills[skillName] ?? 0,
      weaponDamage: finalWeaponDamage,
      strengthBonus,
      attackBonuses,
      manualAttackPenalties,
      automaticTargetPenalty,
      totalAttackPenalties,
      defenseBonuses,
      defensePenalties,
      baseTargetParry,
      targetParry,
      parryMode,
      targetArmor,
      result,
      beforeHp,
      afterHp,
      beforeMinor,
      afterMinor,
      beforeSerious,
      afterSerious,
      critApplied: projectedPartState.critApplied ?? 0
    }),
    result.hit ? "success" : "info"
  );
  await syncState();
  setStatus(`${getCharacterName(attacker)} -> ${resolvedTargetName}: ${result.summary}`, result.hit ? "success" : "info");
}
async function performRollDice() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }
  const dice = Number(getActionFieldValue('[data-roll-field="dice"]')) || 20;
  const count = Number(getActionFieldValue('[data-roll-field="count"]')) || 1;
  const modifier = Number(getActionFieldValue('[data-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `Rolled ${diceLabel}: ${result.rolls.join(", ")}${modifier ? ` ${modifier >= 0 ? "+" : ""}${modifier}` : ""} = ${result.total}`,
      outcome: "roll",
      total: result.total,
      targetPart: "",
      timestamp: (/* @__PURE__ */ new Date()).toISOString(),
      source: "owlbear-extension"
    };
    next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
    return next;
  });
  await ensureOverlayForToken(token.id);
  await pushDebugEntry(`${getCharacterName(token)} rolls dice`, formatDiceDebug({
    tokenName: getCharacterName(token),
    result
  }), "success");
  await syncState();
  setStatus(`${diceLabel} rolled ${result.total}.`, "success");
}
async function addOdysseySkill() {
  if (!isEditable()) {
    setStatus("Only the GM can add skills.", "error");
    return;
  }
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  const name = getActionFieldValue('[data-skill-field="new-name"]').trim();
  const value = clamp(Number(getActionFieldValue('[data-skill-field="new-value"]')) || 0, 0, 10);
  const strengthBonusField = ui.selectedTokenPanel.querySelector(
    '[data-skill-field="new-strength-bonus"]'
  );
  const addStrengthBonus = strengthBonusField instanceof HTMLInputElement ? strengthBonusField.checked : false;
  if (!name) {
    setStatus("Enter a skill name first.", "error");
    return;
  }
  const category = name === MELEE_SKILL_NAME || name === PARRY_SKILL_NAME ? COMBAT_SKILL_CATEGORY : getActionFieldValue('[data-skill-field="new-category"]') === COMBAT_SKILL_CATEGORY ? COMBAT_SKILL_CATEGORY : APPLIED_SKILL_CATEGORY;
  await updateTrackerData(token.id, (current2) => {
    var _a, _b;
    const next = structuredClone(current2);
    next.odyssey.skills[name] = value;
    (_a = next.odyssey).skillCategories ?? (_a.skillCategories = {});
    (_b = next.odyssey).skillStrengthBonuses ?? (_b.skillStrengthBonuses = {});
    next.odyssey.skillCategories[name] = category;
    next.odyssey.skillStrengthBonuses[name] = name === MELEE_SKILL_NAME ? true : name === PARRY_SKILL_NAME ? false : Boolean(addStrengthBonus);
    return next;
  });
  await syncState();
  setStatus(`Skill "${name}" saved for ${getCharacterName(token)}.`, "success");
}
async function removeOdysseySkill(skillName) {
  if (!isEditable()) {
    setStatus("Only the GM can remove skills.", "error");
    return;
  }
  if (CORE_COMBAT_SKILLS.includes(skillName)) {
    setStatus("Core combat skills cannot be removed.", "error");
    return;
  }
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    delete next.odyssey.skills[skillName];
    delete next.odyssey.skillCategories?.[skillName];
    delete next.odyssey.skillStrengthBonuses?.[skillName];
    return next;
  });
  await syncState();
  setStatus(`Skill "${skillName}" removed.`, "success");
}
async function performRollChar() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }
  const attribute = getActionFieldValue('[data-roll-char-field="attribute"]') || "Strength";
  const modifier = Number(getActionFieldValue('[data-roll-char-field="modifier"]')) || 0;
  const odyssey = getOdysseyData(token);
  const result = rollCharacterCheck(odyssey.attributes[attribute] ?? 0, modifier);
  const attributeLabel = ATTRIBUTE_UI_FIELDS.find(([key]) => key === attribute)?.[1] ?? attribute;
  const summary = `Characteristic ${attributeLabel}: ${result.roll} vs ${result.finalAttribute} (${result.result})`;
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    next.lastRoll = {
      eventId: 0,
      actorName: getCharacterName(token),
      summary,
      outcome: "roll-char",
      total: result.roll,
      targetPart: "",
      timestamp: (/* @__PURE__ */ new Date()).toISOString(),
      source: "owlbear-extension"
    };
    next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
    return next;
  });
  await pushDebugEntry(
    `${getCharacterName(token)} rolls characteristic`,
    formatRollCharDebug({
      tokenName: getCharacterName(token),
      attributeLabel,
      result
    }),
    result.result.includes("Failed") ? "info" : "success"
  );
  await syncState();
  setStatus(summary, result.result.includes("Failed") ? "error" : "success");
}
async function performRollSkill() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }
  const skillName = getActionFieldValue('[data-roll-skill-field="skill"]');
  if (!skillName) {
    setStatus("Choose a skill first.", "error");
    return;
  }
  const modifier = Number(getActionFieldValue('[data-roll-skill-field="modifier"]')) || 0;
  const odyssey = getOdysseyData(token);
  const result = rollSkillCheck(odyssey.skills[skillName] ?? 0, modifier);
  const summary = `Skill ${skillName}: ${result.totalPrimary} vs ${result.totalSecondary} (${result.result})`;
  await updateTrackerData(token.id, (current2) => {
    const next = structuredClone(current2);
    next.lastRoll = {
      eventId: 0,
      actorName: getCharacterName(token),
      summary,
      outcome: "roll-skill",
      total: result.totalPrimary,
      targetPart: "",
      timestamp: (/* @__PURE__ */ new Date()).toISOString(),
      source: "owlbear-extension"
    };
    next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
    return next;
  });
  await pushDebugEntry(
    `${getCharacterName(token)} checks skill`,
    formatRollSkillDebug({
      tokenName: getCharacterName(token),
      skillName,
      result
    }),
    result.result === "Check Passed" ? "success" : "info"
  );
  await syncState();
  setStatus(summary, result.result === "Check Passed" ? "success" : "error");
}
async function performPrivateGmRoll() {
  if (!isEditable()) {
    setStatus("Only the GM can use private rolls.", "error");
    return;
  }
  const dice = Number(getActionFieldValue('[data-gm-roll-field="dice"]')) || 20;
  const count = Number(getActionFieldValue('[data-gm-roll-field="count"]')) || 1;
  const modifier = Number(getActionFieldValue('[data-gm-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier, count);
  const diceLabel = `${result.count}d${result.sides}`;
  pushPrivateGmEntry(
    `GM private ${diceLabel}`,
    formatDiceDebug({
      tokenName: "GM private roll",
      result
    })
  );
  render();
  setStatus(`Private GM roll: ${diceLabel} = ${result.total}.`, "success");
}
function bindUiEvents() {
  ui.refreshBtn.addEventListener("click", () => {
    void syncState(true).catch((error) => {
      setStatus(error?.message ?? "Refresh failed.", "error");
    });
  });
  ui.syncBtn.addEventListener("click", () => {
    if (!isEditable()) {
      setStatus("Only the GM can rebuild overlays.", "error");
      return;
    }
    void syncTrackedOverlays().then(() => syncState()).then(() => {
      setStatus("Tracked overlays rebuilt.", "success");
    }).catch((error) => {
      setStatus(error?.message ?? "Overlay rebuild failed.", "error");
    });
  });
  document.addEventListener(
    "toggle",
    (event) => {
      const target = event.target;
      if (!(target instanceof HTMLDetailsElement)) return;
      if (!target.dataset.sectionKey) return;
      collapsibleSectionState.set(target.dataset.sectionKey, target.open);
    },
    true
  );
  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const actionNode = target.closest("[data-action]");
    if (!(actionNode instanceof HTMLElement)) return;
    const action = actionNode.dataset.action;
    const tokenId = actionNode.dataset.tokenId;
    const partName = actionNode.dataset.part;
    const field = actionNode.dataset.field;
    const delta = Number(actionNode.dataset.delta ?? 0);
    const skill = actionNode.dataset.skill;
    if (action === "select-character" && tokenId) {
      void selectCharacter(tokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to select token.", "error");
      });
    }
    if (action === "focus-token" && activeTokenId) {
      void selectCharacter(activeTokenId).catch((error) => {
        setStatus(error?.message ?? "Unable to focus token.", "error");
      });
    }
    if (action === "pick-attack-target") {
      void startTargetPick().catch((error) => {
        setStatus(error?.message ?? "Unable to start target picking.", "error");
      });
      return;
    }
    if (action === "clear-debug-console") {
      void clearDebugConsole().catch((error) => {
        setStatus(error?.message ?? "Unable to clear debug console.", "error");
      });
      return;
    }
    if (action === "change-part" && partName && field) {
      void changeBodyField(partName, field, delta).catch((error) => {
        setStatus(error?.message ?? "Unable to update body value.", "error");
      });
    }
    if (action === "perform-attack") {
      void performAttack().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve attack.", "error");
      });
    }
    if (action === "perform-roll-dice") {
      void performRollDice().catch((error) => {
        setStatus(error?.message ?? "Unable to roll dice.", "error");
      });
      return;
    }
    if (action === "perform-roll-char") {
      void performRollChar().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve Roll_Char.", "error");
      });
      return;
    }
    if (action === "perform-roll-skill") {
      void performRollSkill().catch((error) => {
        setStatus(error?.message ?? "Unable to resolve Roll_Skill.", "error");
      });
      return;
    }
    if (action === "perform-gm-private-roll") {
      void performPrivateGmRoll().catch((error) => {
        setStatus(error?.message ?? "Unable to perform private GM roll.", "error");
      });
      return;
    }
    if (action === "add-skill") {
      void addOdysseySkill().catch((error) => {
        setStatus(error?.message ?? "Unable to add skill.", "error");
      });
      return;
    }
    if (action === "remove-skill" && skill) {
      void removeOdysseySkill(skill).catch((error) => {
        setStatus(error?.message ?? "Unable to remove skill.", "error");
      });
    }
  });
  document.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;
    if (target.dataset.attackField && activeTokenId) {
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
      if (target.dataset.attackField === "targetTokenId") {
        void syncTargetHighlight().catch((error) => {
          console.warn("[Body HP] Unable to sync target highlight", error);
        });
      }
    }
    if (target.dataset.action === "select-owner-player") {
      void setOwnerPlayer(target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save owner.", "error");
      });
      return;
    }
    if (target.dataset.action === "set-odyssey-skill") {
      const skill = target.dataset.skill;
      if (!skill) return;
      void setOdysseySkill(skill, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save skill.", "error");
      });
      return;
    }
    if (target.dataset.action === "set-skill-strength-bonus") {
      const skill = target.dataset.skill;
      if (!skill || !(target instanceof HTMLInputElement)) return;
      void setOdysseySkillStrengthBonus(skill, target.checked).catch((error) => {
        setStatus(error?.message ?? "Unable to save strength bonus flag.", "error");
      });
      return;
    }
    if (target.dataset.action === "set-odyssey-attribute") {
      const attribute = target.dataset.attribute;
      if (!attribute) return;
      void setOdysseyAttribute(attribute, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save attribute.", "error");
      });
      return;
    }
    if (target.dataset.action === "set-weapon-damage") {
      const index = Number(target.dataset.weaponIndex ?? 0);
      void setWeaponDamage(index, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save weapon damage.", "error");
      });
      return;
    }
    if (target.dataset.action === "set-weapon-name") {
      const index = Number(target.dataset.weaponIndex ?? 0);
      void setWeaponName(index, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save weapon name.", "error");
      });
      return;
    }
    if (target.dataset.action !== "set-field") return;
    const partName = target.dataset.part;
    const field = target.dataset.field;
    if (!partName || !field) return;
    void setBodyField(partName, field, target.value).catch((error) => {
      setStatus(error?.message ?? "Unable to save field.", "error");
    });
  });
  document.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!activeTokenId) return;
    if (target.dataset.attackField) {
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
    }
    if (target.dataset.action === "set-odyssey-skill") {
      const skill = target.dataset.skill;
      if (!skill) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-odyssey-skill",
        skill,
        value: target.value
      });
      return;
    }
    if (target.dataset.action === "set-odyssey-attribute") {
      const attribute = target.dataset.attribute;
      if (!attribute) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-odyssey-attribute",
        attribute,
        value: target.value
      });
      return;
    }
    if (target.dataset.action === "set-weapon-damage") {
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-weapon-damage",
        weaponIndex: Number(target.dataset.weaponIndex ?? 0),
        value: target.value
      });
      return;
    }
    if (target.dataset.action === "set-weapon-name") {
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-weapon-name",
        weaponIndex: Number(target.dataset.weaponIndex ?? 0),
        value: target.value
      });
      return;
    }
    if (target.dataset.action === "set-field") {
      const partName = target.dataset.part;
      const field = target.dataset.field;
      if (!partName || !field) return;
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-field",
        partName,
        field,
        value: target.value
      });
    }
  });
}
lib_default.onReady(async () => {
  try {
    bindUiEvents();
    await loadSharedDebugConsole();
    await syncState(true);
    startSelectionPolling();
    setStatus(
      "Ready. Select a character token on the map to edit it here.",
      "info"
    );
    lib_default.scene.items.onChange((items) => {
      sceneItems = items;
      render();
      void syncTargetHighlight().catch((error) => {
        console.warn("[Body HP] Unable to sync target highlight", error);
      });
    });
    lib_default.player.onChange((player) => {
      playerRole = player.role;
      playerId = player.id ?? playerId;
      playerName = player.name ?? playerName;
      playerColor = player.color ?? playerColor;
      selectionIds = player.selection ?? [];
      const selectedCharacterId = selectionIds.find(
        (selectionId) => sceneItems.some((item) => item.id === selectionId && isCharacterToken(item))
      );
      if (selectedCharacterId) {
        activeTokenId = selectedCharacterId;
      }
      void syncState().catch((error) => {
        console.warn("[Body HP] Player state sync failed", error);
        render();
      });
    });
    lib_default.party.onChange((players) => {
      partyPlayers = players ?? [];
      render();
    });
    lib_default.broadcast.onMessage(DEBUG_BROADCAST_CHANNEL, (event) => {
      const payload = event?.data;
      if (!payload || typeof payload !== "object") return;
      if (payload.type === "debug-clear") {
        debugEntries = [];
        renderDebugConsole();
        return;
      }
      if (payload.type !== "debug-entry") return;
      const nextEntries = mergeDebugEntries([payload.entry], debugEntries);
      debugEntries = nextEntries;
      renderDebugConsole();
      if (playerRole === "GM") {
        void lib_default.room.setMetadata({
          [DEBUG_LOG_KEY]: nextEntries
        }).catch((error) => {
          console.warn("[Body HP] Unable to persist broadcast debug entry", error);
        });
      }
    });
    lib_default.room.onMetadataChange((metadata) => {
      debugEntries = sanitizeDebugEntries(metadata?.[DEBUG_LOG_KEY]);
      renderDebugConsole();
    });
  } catch (error) {
    setStatus(error?.message ?? "Extension failed to initialize.", "error");
  }
});
