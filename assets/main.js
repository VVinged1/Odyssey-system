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
    EventEmitter2.prototype.setMaxListeners = function setMaxListeners(n2) {
      if (typeof n2 !== "number" || n2 < 0 || NumberIsNaN(n2)) {
        throw new RangeError('The value of "n" is out of range. It must be a non-negative number. Received ' + n2 + ".");
      }
      this._maxListeners = n2;
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
    function arrayClone(arr, n2) {
      var copy = new Array(n2);
      for (var i = 0; i < n2; ++i)
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

// rangeRules.js
var WEAPON_CATEGORIES = [["", "Unclassified"], ["pistol", "Pistol"], ["smg", "SMG"], ["shotgun", "Shotgun"], ["assault", "Assault rifle / machine gun"], ["sniper", "Sniper rifle"]];
var normalizeWeaponCategory = (value) => WEAPON_CATEGORIES.some(([key]) => key === value) ? value : "";
var weaponCategoryLabel = (value) => WEAPON_CATEGORIES.find(([key]) => key === value)?.[1] || "Unclassified";
function distanceMeters(cells, scale) {
  const value = Number(scale?.parsed?.multiplier), unit = String(scale?.parsed?.unit || "").trim().toLowerCase();
  const factor = { m: 1, meter: 1, meters: 1, metre: 1, metres: 1, "\u043C": 1, cm: 0.01, "\u0441\u043C": 0.01, km: 1e3, "\u043A\u043C": 1e3, ft: 0.3048, foot: 0.3048, feet: 0.3048, "\u0444\u0443\u0442": 0.3048, "\u0444\u0443\u0442\u044B": 0.3048, in: 0.0254, inch: 0.0254, inches: 0.0254, yd: 0.9144, yard: 0.9144, yards: 0.9144, mi: 1609.344, mile: 1609.344, miles: 1609.344 }[unit];
  return Number.isFinite(cells) && cells >= 0 && Number.isFinite(value) && value > 0 && factor ? cells * value * factor : null;
}
var RANGE_BANDS = [["clinch", "Clinch \xB7 0\u20132 m"], ["short", "Short \xB7 >2\u201315 m"], ["medium", "Medium \xB7 >15\u201350 m"], ["long", "Long \xB7 >50 m"]];
function distanceBand(meters) {
  return !Number.isFinite(meters) || meters < 0 ? null : meters <= 2 ? "clinch" : meters <= 15 ? "short" : meters <= 50 ? "medium" : "long";
}
function normalizeRangeReduction(raw) {
  return Object.fromEntries(RANGE_BANDS.map(([key]) => [key, Math.max(0, Math.min(999, Math.trunc(Number(raw?.[key]) || 0)))]).filter(([, v]) => v));
}
var MATRIX = { pistol: [-15, 0, -5, -10], smg: [-25, 15, 0, -10], shotgun: [-25, 30, -15, -50], assault: [-30, 0, 0, 0], sniper: [-50, -30, 0, 10] };
function rangeAdjustment(weapon, meters, odyssey = {}, skill = "") {
  const category = normalizeWeaponCategory(weapon?.weaponCategory), band = distanceBand(meters);
  if (!category) return { category, band, raw: 0, reduction: 0, modifier: 0, required: false };
  if (!band) return { category, band, raw: 0, reduction: 0, modifier: 0, required: true };
  let raw = MATRIX[category][RANGE_BANDS.findIndex(([key]) => key === band)];
  if (band === "long" && ["pistol", "smg", "sniper"].includes(category)) raw = Math.sign(raw) * Math.min(50, Math.ceil((meters - 50) / 25) * 10);
  const reduction = [weapon.rangeReduction, odyssey.skillRangeReduction?.[skill], odyssey.rangeReduction].reduce((sum, row) => sum + (normalizeRangeReduction(row)[band] || 0), 0);
  return { category, band, raw, reduction, modifier: raw < 0 ? Math.min(0, raw + reduction) : raw, required: false };
}
function rangeAdjustmentLabel(result) {
  if (!result.category) return "Range: unclassified";
  if (result.required) return "Range: distance in metres required";
  const band = RANGE_BANDS.find(([key]) => key === result.band)[1].split(" \xB7 ")[0], signed = (n2) => (n2 >= 0 ? "+" : "") + n2;
  return band + " \xB7 " + signed(result.modifier) + (result.raw < 0 && result.reduction ? " (base " + result.raw + ", relief " + Math.min(result.reduction, -result.raw) + ")" : "");
}

// targetDistance.js
function distanceLabel(cells, scale) {
  if (!Number.isFinite(cells) || cells < 0) throw new Error("Distance is unavailable.");
  const digits = Math.max(0, Math.min(5, Number(scale?.parsed?.digits) || 0)), multiplier = Number(scale?.parsed?.multiplier) || 1;
  const format = (n2, d, min = 0) => Number(n2.toFixed(d)).toLocaleString("en-US", { minimumFractionDigits: min, maximumFractionDigits: d, useGrouping: false });
  const unit = String(scale?.parsed?.unit || "");
  return format(cells, 2) + " cells" + (unit ? " \xB7 " + format(cells * multiplier, digits, digits) + unit : "");
}
function createTargetDistance({ OBR: OBR2, buildRuler: buildRuler2, onRead, context = "main" }) {
  const key = "com.codex.body-hp/target-distance", id = crypto.randomUUID();
  let request = 0, generation = 0, timer = null, latest2 = null, added = false, chain = Promise.resolve(), scalePromise = null;
  const cache = /* @__PURE__ */ new Map();
  const initialize = OBR2.scene.local.getItems().then((items) => {
    const ids = items.filter((i) => i.metadata?.[key] === context).map((i) => i.id);
    if (ids.length) return OBR2.scene.local.deleteItems(ids);
  }).catch(() => {
  });
  const center = async (target) => target.id ? (await OBR2.scene.items.getItemBounds([target.id])).center : target;
  async function read(source, target) {
    const cacheKey = JSON.stringify([source.id, source.position, source.scale, source.rotation, source.grid, source.image?.width, source.image?.height, target.id, target.position || target, target.scale, target.rotation, target.grid, target.image?.width, target.image?.height]);
    if (!cache.has(cacheKey)) {
      if (cache.size >= 32) cache.clear();
      const promise = (async () => {
        const [from, to, scale] = await Promise.all([center(source), center(target), scalePromise ?? (scalePromise = OBR2.scene.grid.getScale())]);
        const cells = await OBR2.scene.grid.getDistance(from, to);
        return { sourceId: source.id, targetId: target.id || "", from, to, cells, meters: distanceMeters(cells, scale), label: distanceLabel(cells, scale) };
      })();
      cache.set(cacheKey, promise);
      promise.catch(() => {
        if (cache.get(cacheKey) === promise) cache.delete(cacheKey);
      });
    }
    return cache.get(cacheKey);
  }
  const publish = async (source, target, preview, epoch) => {
    const current2 = ++request;
    try {
      const reading = await read(source, target);
      if (epoch !== generation || current2 !== request) return preview ? null : reading;
      onRead(reading);
      if (preview) {
        chain = chain.catch(() => {
        }).then(async () => {
          await initialize;
          if (epoch !== generation || current2 !== request) return;
          if (!added) {
            const ruler = buildRuler2().id(id).name("Odyssey target distance").startPosition(reading.from).endPosition(reading.to).measurement(reading.label).variant("DASHED").layer("RULER").locked(true).disableHit(true).metadata({ [key]: context }).build();
            await OBR2.scene.local.addItems([ruler]);
            added = true;
          } else await OBR2.scene.local.updateItems([id], (items) => {
            if (epoch === generation && current2 === request && items[0]) {
              items[0].startPosition = reading.from;
              items[0].endPosition = reading.to;
              items[0].measurement = reading.label;
            }
          }, true);
        });
        await chain;
      }
      return reading;
    } catch {
      if (epoch === generation && current2 === request) onRead(null);
      return null;
    }
  };
  return {
    preview(source, target) {
      latest2 = { source, target };
      if (timer) return;
      const epoch = generation;
      timer = setTimeout(() => {
        timer = null;
        const next = latest2;
        if (next) void publish(next.source, next.target, true, epoch);
      }, 80);
    },
    measure(source, target) {
      return publish(source, target, false, generation);
    },
    invalidate() {
      generation++;
      request++;
      cache.clear();
      scalePromise = null;
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
      if (latest2) this.preview(latest2.source, latest2.target);
    },
    async clear() {
      generation++;
      request++;
      latest2 = null;
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
      chain = chain.catch(() => {
      }).then(async () => {
        await initialize;
        if (added) {
          await OBR2.scene.local.deleteItems([id]);
          added = false;
        }
      });
      await chain.catch(() => {
      });
    }
  };
}

// ui/combatLayout.js
function arrangeCombatPanels(root = document) {
  for (const panel of root.querySelectorAll("[data-combat-layout]:not([data-combat-arranged])")) {
    panel.dataset.combatArranged = "true";
    const kind = panel.dataset.combatLayout;
    const grid = document.createElement("div");
    grid.className = "od-combat-grid";
    const cards = ["Weapon & skill", kind === "manual" ? "Manual defense" : "Target", "Action"].map((title) => {
      const card = document.createElement("section");
      card.className = "od-combat-card";
      const heading = document.createElement("div");
      heading.className = "od-combat-title";
      heading.textContent = title;
      card.append(heading);
      grid.append(card);
      return card;
    });
    for (const field of [...panel.querySelectorAll("label.field-stack")]) {
      const control = field.querySelector("[data-attack-field],[data-manual-attack-field],input,select,button,[id],.combat-value");
      const key = control?.dataset.attackField || control?.dataset.manualAttackField || control?.id || control?.dataset.action || "";
      const isWeapon = ["skill", "weaponName", "weaponDamage", "ammoId", "rounds", "attackSourceName", "attackSkill", "weaponAccuracy"].includes(key) || field.querySelector("[data-weapon-accuracy-display]");
      const isTarget = ["distanceMeters", "targetPart", "targetName", "targetTokenId", "pick-attack-target", "pickTargetBtn", "clearTargetBtn", "manualArmor", "manualParry"].includes(key) || field.querySelector(".combat-value:not([data-weapon-accuracy-display])");
      cards[isWeapon ? 0 : isTarget ? 1 : 2].append(field);
    }
    for (const hidden of [...panel.querySelectorAll('input[type="hidden"]')]) cards[1].append(hidden);
    for (const note of [...panel.querySelectorAll(".combat-notes")]) cards[2].append(note);
    const notes = [...panel.querySelectorAll(":scope > .muted")];
    if (notes.length) {
      const details2 = document.createElement("details");
      details2.className = "combat-notes";
      const summary = document.createElement("summary");
      summary.textContent = "Attack rules";
      details2.append(summary, ...notes);
      cards[2].append(details2);
    }
    const footer = document.createElement("div");
    footer.className = "od-combat-footer";
    const attack = panel.querySelector('[data-action="perform-attack"],[data-action="perform-manual-attack"],#environmentAttackBtn');
    const pick = cards[1].querySelector('[data-action="pick-attack-target"],#pickTargetBtn');
    const wide = (button, label) => {
      button.classList.remove("od-icon-button");
      button.classList.add("od-wide-action");
      button.textContent = label;
      button.title = label;
      button.setAttribute("aria-label", label);
      footer.append(button);
    };
    if (attack) wide(attack, kind === "manual" ? "No Target Attack" : "Attack");
    if (pick) {
      const field = pick.closest("label");
      wide(pick, /cancel/i.test(pick.textContent || pick.title) ? "Cancel Target Pick" : "Pick Target On Map");
      field?.remove();
    }
    for (const obsolete of [...panel.querySelectorAll(":scope > .combat-action,:scope > .row")]) obsolete.remove();
    for (const obsolete of [...panel.querySelectorAll(":scope > .combat-form-grid,:scope > .combat-modifiers")]) obsolete.remove();
    panel.append(grid, footer);
    panel.classList.add("od-combat-panel");
  }
}

// ui/compactButtons.js
var paths = {
  plus: "M12 5v14M5 12h14",
  close: "m6 6 12 12M18 6 6 18",
  edit: "m4 16-1 5 5-1L20 8l-4-4Z M14 6l4 4",
  save: "M5 3h13l3 3v15H3V3Z M7 3v6h10V3 M7 21v-8h10v8",
  refresh: "M20 7v5h-5M4 17v-5h5 M5 8a8 8 0 0 1 13-3l2 3 M19 16a8 8 0 0 1-13 3l-2-3",
  target: "M12 2v4m0 12v4M2 12h4m12 0h4 M19 12a7 7 0 1 1-14 0 7 7 0 0 1 14 0 M12 10v4m-2-2h4",
  download: "M12 3v12m-5-5 5 5 5-5 M4 16v5h16v-5",
  upload: "M12 16V4m-5 5 5-5 5 5 M4 16v5h16v-5",
  cloud: "M6 18a4 4 0 0 1-1-8 7 7 0 0 1 13-1 5 5 0 0 1 0 9Z M12 10v8m-3-3 3 3 3-3",
  shield: "m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6Z",
  dice: "M4 4h16v16H4Z M8 8h.01 M16 8h.01 M12 12h.01 M8 16h.01 M16 16h.01",
  attack: "m4 3 16 17m0-17L4 20 M3 15l6 6m6-18 6 6",
  heal: "M9 3h6v6h6v6h-6v6H9v-6H3V9h6Z",
  book: "M3 4h7l2 2 2-2h7v15h-7l-2 2-2-2H3Z M12 6v15",
  play: "m8 4 12 8-12 8Z",
  box: "m3 7 9-4 9 4v13H3Z M3 7l9 4 9-4 M12 11v9",
  window: "M3 3h18v18H3Z M3 8h18 M8 8v13",
  copy: "M8 8h13v13H8Z M16 8V3H3v13h5"
};
function iconFor(text2) {
  if (/cancel|clear|remove|delete|unequip|detach|close/i.test(text2)) return "close";
  if (/edit|rename|set rounds|set magazine/i.test(text2)) return "edit";
  if (/duplicate/i.test(text2)) return "copy";
  if (/export|empty to|load from/i.test(text2)) return "download";
  if (/import/i.test(text2)) return "upload";
  if (/account/i.test(text2)) return "cloud";
  if (/save/i.test(text2)) return "save";
  if (/add|new|^\+/i.test(text2)) return "plus";
  if (/reload|refresh|rebuild|fill|sync/i.test(text2)) return "refresh";
  if (/select|pick|target|focus/i.test(text2)) return "target";
  if (/roll|d100|dice/i.test(text2)) return "dice";
  if (/attack|resolve/i.test(text2)) return "attack";
  if (/heal|repair/i.test(text2)) return "heal";
  if (/equip/i.test(text2)) return "shield";
  if (/template/i.test(text2)) return "book";
  if (/hud|window|open/i.test(text2)) return "window";
  if (/use|apply/i.test(text2)) return "play";
  return "box";
}
function installCompactButtons() {
  const enhance = () => {
    for (const button of document.querySelectorAll("button")) {
      if (button.matches(".cp-attr,.cp-skill-card,.list-item,.od-wide-action") || button.querySelector("[data-compact-icon]")) continue;
      const text2 = button.textContent.trim();
      if (!text2 || text2.length <= 2) continue;
      button.title = text2;
      button.setAttribute("aria-label", text2);
      button.classList.add("od-icon-button");
      button.innerHTML = `<svg data-compact-icon viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${paths[iconFor(text2)]}"></path></svg>`;
    }
  };
  let scheduled = false;
  new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;
    queueMicrotask(() => {
      scheduled = false;
      enhance();
    });
  }).observe(document.body, { childList: true, subtree: true, characterData: true });
  enhance();
}

// characterRules.js
var ARMOR_TRAINING = { light_armor: "Light Armor", medium_armor: "Medium Armor", heavy_armor: "Heavy Armor", superheavy_armor: "Superheavy Armor" };
var ATTRIBUTES = ["Strength", "Agility", "Reaction", "Endurance", "Perception", "Intelligence", "Charisma", "Willpower", "Magic"];
var n = (v) => Math.max(-999, Math.min(999, Math.trunc(Number(v) || 0)));
function normalizeBonuses(raw) {
  const clean = (rows2, attributes = false) => Object.fromEntries(Object.entries(rows2 || {}).filter(([k, v]) => k.length <= 80 && !["__proto__", "constructor", "prototype"].includes(k) && (!attributes || ATTRIBUTES.includes(k)) && n(v)).slice(0, 80).map(([k, v]) => [k, n(v)]));
  const rangeReduction = normalizeRangeReduction(raw?.rangeReduction);
  return { attributes: clean(raw?.attributes, true), skills: clean(raw?.skills), ...Object.keys(rangeReduction).length ? { rangeReduction } : {} };
}
function normalizeEffects(raw) {
  const seen = /* @__PURE__ */ new Set();
  return (Array.isArray(raw) ? raw : []).slice(0, 100).filter((e) => e && typeof e.id === "string" && !seen.has(e.id) && seen.add(e.id)).map((e) => ({ id: e.id.slice(0, 120), name: String(e.name || "Effect").slice(0, 120), bonuses: normalizeBonuses(e.bonuses) }));
}
function equipmentFits(item, partName, part) {
  const slots = Array.isArray(item.slots) ? item.slots : [item.slot || "any"];
  return slots.includes("any") || slots.includes(partName) || slots.includes(part?.slot);
}
function itemActive(item) {
  return item.quantity > 0 && (item.kind === "implant" || !item.criticalMax || item.criticalDamage < item.criticalMax);
}
function armorSummary(o) {
  const total = (o.inventory?.items || []).filter((i) => ["armor", "shield"].includes(i.kind) && i.equippedPart && itemActive(i)).reduce((sum, i) => sum + i.armor, 0);
  const key = total <= 0 ? null : total <= 70 ? "light_armor" : total <= 140 ? "medium_armor" : total <= 260 ? "heavy_armor" : "superheavy_armor";
  return { total, key, name: key ? ARMOR_TRAINING[key] : "No armor" };
}
function effectiveOdyssey(base) {
  const o = structuredClone(base), rows2 = [];
  o.baseAttributes = structuredClone(base.attributes);
  o.baseSkills = structuredClone(base.skills);
  o.grantedSkills = {};
  for (const i of o.inventory?.items || []) {
    if (!itemActive(i)) continue;
    if (i.kind === "resource" || ["armor", "shield", "implant"].includes(i.kind) && i.equippedPart) rows2.push({ id: i.id, name: i.name, bonuses: i.bonuses, passive: true });
    if (i.kind === "implant" && i.equippedPart) for (const a of i.abilities || []) {
      const key = `implant:${i.id}:${a.id}`;
      o.skills[key] = a.level;
      o.skillCategories[key] = "abilities";
      o.skillLabels[key] = a.name;
      o.skillPerks[key] = a.description ? [a.description] : [];
      o.grantedSkills[key] = i.name;
    }
  }
  rows2.push(...base.effects || []);
  for (const row of rows2) for (const key of Object.keys(row.bonuses?.skills || {})) if (!Object.hasOwn(o.skills, key)) {
    o.skills[key] = 0;
    o.baseSkills[key] = 0;
    o.skillLabels[key] = key;
    o.skillCategories[key] = "applied";
    o.skillStrengthBonuses[key] = false;
    o.skillPerks[key] = [];
    o.grantedSkills[key] = row.name;
  }
  const skillMods = {};
  for (const row of rows2) for (const [k, v] of Object.entries(row.bonuses?.skills || {})) skillMods[k] = (skillMods[k] || 0) + v;
  for (const [k, v] of Object.entries(skillMods)) if (Object.hasOwn(o.skills, k)) o.skills[k] = Math.max(0, Math.min(Object.hasOwn(ARMOR_TRAINING, k) ? 1 : 10, o.skills[k] + v));
  const armor = armorSummary(o), trained = !armor.key || (o.skills[armor.key] || 0) >= 1;
  if (!trained) {
    const penalties = armor.key === "light_armor" ? { Agility: -1, Endurance: -1 } : armor.key === "medium_armor" ? { Agility: -2, Endurance: -2, Strength: -2 } : { Agility: -4, Endurance: -5, Strength: -4 };
    rows2.push({ id: "armor-penalty", name: `${armor.name} \xB7 untrained (${armor.total})`, bonuses: { attributes: penalties, skills: {} }, passive: true });
  }
  o.modifiers = { attributes: {}, skills: skillMods };
  for (const row of rows2) for (const [k, v] of Object.entries(row.bonuses?.attributes || {})) o.modifiers.attributes[k] = (o.modifiers.attributes[k] || 0) + v;
  for (const k of ATTRIBUTES) o.attributes[k] = Math.max(0, (base.attributes[k] || 0) + (o.modifiers.attributes[k] || 0));
  o.rangeReduction = Object.fromEntries(RANGE_BANDS.map(([key]) => [key, rows2.reduce((sum, row) => sum + (normalizeRangeReduction(row.bonuses?.rangeReduction)[key] || 0), 0)]));
  o.effectRows = rows2.filter((r) => Object.keys(r.bonuses?.attributes || {}).length || Object.keys(r.bonuses?.skills || {}).length || Object.keys(r.bonuses?.rangeReduction || {}).length);
  o.armorSummary = { ...armor, trained };
  return o;
}
function reconcileEquipment(tracker) {
  const inv = tracker.odyssey.inventory;
  if (!inv) return;
  inv.baseArmor ?? (inv.baseArmor = {});
  inv.baseMax ?? (inv.baseMax = {});
  for (const [partName, part] of Object.entries(tracker.body)) {
    const equipped = inv.items.filter((i) => ["armor", "shield", "implant"].includes(i.kind) && i.equippedPart === partName);
    const implantHp = equipped.filter((i) => i.kind === "implant" && itemActive(i)).reduce((sum, i) => sum + (Number(i.criticalMax) || 0), 0);
    const armor = equipped.filter((i) => ["armor", "shield"].includes(i.kind) && i.quantity > 0);
    const pool = { max: armor.reduce((sum, i) => sum + (Number(i.criticalMax) || 0), 0), remaining: armor.reduce((sum, i) => sum + Math.max(0, (Number(i.criticalMax) || 0) - (Number(i.criticalDamage) || 0)), 0) };
    inv.armorHp ?? (inv.armorHp = {});
    const previous = inv.armorHp[partName]?.remaining || 0;
    const hpBonus = implantHp + pool.max;
    if (hpBonus && !Object.hasOwn(inv.baseMax, partName)) inv.baseMax[partName] = part.max;
    if (Object.hasOwn(inv.baseMax, partName)) {
      part.max = Math.min(999, inv.baseMax[partName] + hpBonus);
      part.current = Math.max(0, Math.min(part.current + pool.remaining - previous, part.max));
      if (!hpBonus) delete inv.baseMax[partName];
    }
    if (pool.max) inv.armorHp[partName] = pool;
    else delete inv.armorHp[partName];
    if (!equipped.length) {
      if (Object.hasOwn(inv.baseArmor, partName)) {
        part.armor = inv.baseArmor[partName];
        delete inv.baseArmor[partName];
      }
      continue;
    }
    if (!Object.hasOwn(inv.baseArmor, partName)) {
      const legacy = equipped.find((i) => i.legacyEquipment);
      inv.baseArmor[partName] = legacy ? legacy.previousArmor : part.armor;
    }
    part.armor = inv.baseArmor[partName] + equipped.filter(itemActive).reduce((sum, i) => sum + i.armor, 0);
    for (const i of equipped) i.legacyEquipment = false;
  }
}
function absorbEquipmentCritical(tracker, partName, critical) {
  let remaining = Math.max(0, Number(critical) || 0), absorbed = 0;
  const items = tracker.odyssey.inventory?.items || [];
  for (const i of items.filter((i2) => ["armor", "shield"].includes(i2.kind) && i2.equippedPart === partName && i2.criticalMax > 0 && itemActive(i2)).sort((a, b) => (a.kind === "implant") - (b.kind === "implant"))) {
    const amount = Math.min(remaining, i.criticalMax - i.criticalDamage);
    i.criticalDamage += amount;
    remaining -= amount;
    absorbed += amount;
    if (!remaining) break;
  }
  reconcileEquipment(tracker);
  return { remaining, absorbed };
}
function armorHealth(tracker, partName) {
  const gear = tracker.odyssey.inventory.items.filter((i) => ["armor", "shield"].includes(i.kind) && i.equippedPart === partName && i.quantity > 0);
  return { max: gear.reduce((sum, i) => sum + i.criticalMax, 0), remaining: gear.reduce((sum, i) => sum + i.criticalMax - i.criticalDamage, 0) };
}
function setCombinedPartMax(tracker, partName, value) {
  const part = tracker.body[partName];
  if (!part) return;
  const inv = tracker.odyssey.inventory, goal = Math.max(0, Math.min(999, Number(value) || 0));
  if (Object.hasOwn(inv.baseMax || {}, partName)) inv.baseMax[partName] = Math.max(0, goal - (part.max - inv.baseMax[partName]));
  part.max = goal;
  part.current = Math.min(part.current, part.max);
  reconcileEquipment(tracker);
}
function setCombinedPartHp(tracker, partName, value) {
  const part = tracker.body[partName];
  if (!part) return;
  const goal = Math.max(0, Math.min(Number(value) || 0, part.max)), delta = goal - part.current;
  if (delta < 0) absorbEquipmentCritical(tracker, partName, -delta);
  if (delta > 0) {
    const pool = armorHealth(tracker, partName), bodyCurrent = Math.max(0, part.current - pool.remaining), bodyMax = Math.max(0, part.max - pool.max);
    let repair = Math.max(0, delta - Math.max(0, bodyMax - bodyCurrent));
    for (const i of tracker.odyssey.inventory.items.filter((i2) => ["armor", "shield"].includes(i2.kind) && i2.equippedPart === partName && i2.quantity > 0)) {
      const heal = Math.min(repair, i.criticalDamage);
      i.criticalDamage -= heal;
      repair -= heal;
    }
    reconcileEquipment(tracker);
  }
  part.current = goal;
}
function useItemOnTracker(target, item, partName, effectId) {
  var _a;
  if (item.subtype === "grenade") throw new Error("Grenades are a placeholder; no action is performed.");
  if (item.subtype === "stim") {
    (_a = target.odyssey).effects ?? (_a.effects = []);
    if (target.odyssey.effects.length >= 100) throw new Error("Too many active effects.");
    target.odyssey.effects.push({ id: effectId, name: item.name, bonuses: normalizeBonuses(item.bonuses) });
    return;
  }
  const part = target.body[partName];
  if (!part) throw new Error("Choose a body part.");
  if (item.subtype === "medkit") {
    const pool = armorHealth(target, partName), missing = Math.max(0, part.max - pool.max - Math.max(0, part.current - pool.remaining));
    if (!item.healCritical || !missing) throw new Error("No critical wounds to heal on this part.");
    part.current += Math.min(missing, item.healCritical);
    return;
  }
  if (item.subtype === "repairkit") {
    const gear = target.odyssey.inventory.items.filter((i) => i.equippedPart === partName && ["armor", "shield"].includes(i.kind) && i.quantity > 0 && i.criticalDamage > 0);
    if (!gear.length) throw new Error("No damaged armor is equipped on this part.");
    if (!item.healCritical) throw new Error("Repair amount must be greater than zero.");
    let remaining = item.healCritical;
    for (const i of gear) {
      const heal = Math.min(remaining, i.criticalDamage);
      i.criticalDamage -= heal;
      remaining -= heal;
    }
    reconcileEquipment(target);
    return;
  }
  throw new Error("This item cannot be used.");
}

// inventory.js
var LIMIT = 2e3;
var int = (v, max = 999999) => Math.min(max, Math.max(0, Math.floor(Number(v) || 0)));
var text = (v, max = 120) => String(v ?? "").trim().slice(0, max);
var kinds = /* @__PURE__ */ new Set(["item", "consumable", "medical", "tool", "resource", "armor", "shield", "implant"]);
function newId() {
  return globalThis.crypto.randomUUID();
}
function rows(raw, map) {
  const seen = /* @__PURE__ */ new Set();
  return (Array.isArray(raw) ? raw : []).slice(0, LIMIT).filter((r) => r && typeof r === "object").map(map).filter((r) => r.id && !seen.has(r.id) && seen.add(r.id));
}
function normalizeInventory(raw, weapons = []) {
  const items = rows(raw?.items, (r, i) => ({
    id: text(r.id || `legacy-item-${i}`),
    name: text(r.name) || "Item",
    kind: kinds.has(r.kind) ? r.kind : "item",
    quantity: int(r.quantity ?? 1),
    description: text(r.description, 2e3),
    armor: int(r.armor, 999),
    subtype: ["grenade", "medkit", "repairkit", "stim"].includes(r.subtype) ? r.subtype : r.kind === "medical" ? "medkit" : "grenade",
    healCritical: int(r.healCritical ?? 1, 999),
    slots: [...new Set((Array.isArray(r.slots) ? r.slots : [r.slot || "any"]).map((s) => text(s, 40)).filter(Boolean))].slice(0, 100),
    slot: text(r.slot || "any", 40),
    criticalMax: int(r.criticalMax, 999),
    criticalDamage: int(r.criticalDamage, int(r.criticalMax, 999)),
    playerRemovable: r.playerRemovable === true,
    bonuses: normalizeBonuses(r.bonuses),
    abilities: (Array.isArray(r.abilities) ? r.abilities : []).slice(0, 20).map((a, j) => ({ id: text(a.id || "ability-" + j), name: text(a.name) || "Ability", level: int(a.level, 10), description: text(a.description, 2e3) })),
    legacyEquipment: r.legacyEquipment ?? (raw?.version < 3 || !raw?.version),
    equippedPart: text(r.equippedPart, 40),
    previousArmor: int(r.previousArmor, 999)
  }));
  for (const item of [...items]) if (item.equippedPart && ["armor", "shield", "implant"].includes(item.kind) && item.quantity > 1) {
    let id = item.id.slice(0, 90) + "-reserve", suffix = 1;
    while (items.some((i) => i.id === id)) id = item.id.slice(0, 90) + "-reserve-" + suffix++;
    items.push({ ...structuredClone(item), id, quantity: item.quantity - 1, equippedPart: "", previousArmor: 0, legacyEquipment: false });
    item.quantity = 1;
  }
  const magazines = rows(raw?.magazines, (r, i) => ({
    id: text(r.id || `legacy-magazine-${i}`),
    name: text(r.name) || "Magazine",
    capacity: int(r.capacity, 999),
    loadedAmmoId: text(r.loadedAmmoId),
    loaded: int(r.loaded, int(r.capacity, 999))
  }));
  const assigned = /* @__PURE__ */ new Set();
  for (const [i, weapon] of weapons.entries()) {
    if (weapon.acceptsMagazines === false) {
      weapon.magazineId = "";
      continue;
    }
    const requested = text(weapon.magazineId);
    let id = requested || ((!Object.hasOwn(weapon, "magazineId") || weapon.loaded > 0) && weapon.capacity > 0 ? `legacy-mag-${weapon.id || i}` : "");
    if (!id) {
      weapon.magazineId = "";
      weapon.loaded = 0;
      weapon.loadedAmmoId = "";
      continue;
    }
    if (assigned.has(id)) id = `recovered-mag-${weapon.id || i}-${i}`;
    let mag = magazines.find((m) => m.id === id);
    if (!mag) {
      mag = { id, name: `${weapon.name} magazine`, capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId };
      magazines.push(mag);
    }
    weapon.magazineId = id;
    assigned.add(id);
    Object.assign(mag, { capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId });
  }
  return { version: 5, items, magazines, armorHp: Object.fromEntries(Object.entries(raw?.armorHp || {}).filter(([k]) => !["__proto__", "constructor", "prototype"].includes(k)).map(([k, v]) => [text(k, 40), { max: int(v?.max, 99999), remaining: int(v?.remaining, 99999) }])), baseMax: Object.fromEntries(Object.entries(raw?.baseMax || {}).filter(([k]) => !["__proto__", "constructor", "prototype"].includes(k)).map(([k, v]) => [text(k, 40), int(v, 99999)])), baseArmor: Object.fromEntries(Object.entries(raw?.baseArmor || {}).filter(([k]) => !["__proto__", "constructor", "prototype"].includes(k)).map(([k, v]) => [text(k, 40), int(v, 999)])) };
}
function magazineWeapon(odyssey, id) {
  return odyssey.weapons.ranged.find((w) => w.magazineId === id);
}
function magazineCompatible(odyssey, weapon, magazine) {
  if (weapon.acceptsMagazines === false) return false;
  if (!Array.isArray(weapon.acceptedMagazineIds)) return magazine.capacity === weapon.capacity && (!magazine.loaded || weapon.ammoIds.includes(magazine.loadedAmmoId));
  return weapon.acceptedMagazineIds.some((id) => id === magazine.id || odyssey.inventory.magazines.some((reference) => reference.id === id && reference.name === magazine.name && reference.capacity === magazine.capacity));
}
function swapMagazine(odyssey, weaponIndex, magazineId) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  if (!weapon) throw new Error("Weapon no longer exists.");
  if (weapon.acceptsMagazines === false) throw new Error("This weapon has a built-in magazine.");
  const current2 = odyssey.inventory.magazines.find((m) => m.id === weapon.magazineId);
  if (current2) Object.assign(current2, { capacity: weapon.capacity, loaded: weapon.loaded, loadedAmmoId: weapon.loadedAmmoId });
  const next = magazineId ? odyssey.inventory.magazines.find((m) => m.id === magazineId) : null;
  if (magazineId && !next) throw new Error("Magazine no longer exists.");
  if (next && magazineWeapon(odyssey, next.id) && weapon.magazineId !== next.id) throw new Error("Magazine is already mounted.");
  const selectedMagazines = Array.isArray(weapon.acceptedMagazineIds);
  if (next && selectedMagazines && !magazineCompatible(odyssey, weapon, next)) throw new Error("This magazine is incompatible with the weapon.");
  if (next && !selectedMagazines && next.capacity !== weapon.capacity) throw new Error("Magazine capacity must match this weapon.");
  if (next?.loaded && !selectedMagazines && !weapon.ammoIds.includes(next.loadedAmmoId)) throw new Error("Magazine ammunition is incompatible.");
  if (selectedMagazines) {
    weapon.capacity = next?.capacity || 0;
    weapon.ammoIds = next?.loadedAmmoId ? [next.loadedAmmoId] : [];
  }
  weapon.magazineId = next?.id || "";
  weapon.loaded = next?.loaded || 0;
  weapon.loadedAmmoId = next?.loadedAmmoId || "";
}
function fillMagazine(odyssey, id, ammoId) {
  const mag = odyssey.inventory.magazines.find((m) => m.id === id);
  const ammo = odyssey.ammunition.find((a) => a.id === ammoId);
  if (!mag || !ammo) throw new Error("Choose a magazine and ammunition.");
  if (magazineWeapon(odyssey, id)) throw new Error("Detach the magazine before filling it, or use weapon reload.");
  if (mag.loaded && mag.loadedAmmoId !== ammoId) throw new Error("Empty the magazine before changing ammunition.");
  const count = Math.min(mag.capacity - mag.loaded, ammo.quantity);
  if (count < 1) throw new Error("Magazine is full or reserve is empty.");
  mag.loadedAmmoId = ammoId;
  mag.loaded += count;
  ammo.quantity -= count;
}
function setMagazineRounds(odyssey, id, rounds, ammoId) {
  const mag = odyssey.inventory.magazines.find((m) => m.id === id);
  if (!mag) throw new Error("Magazine no longer exists.");
  if (!Number.isInteger(rounds) || rounds < 0 || rounds > mag.capacity) {
    throw new Error(`Rounds must be a whole number from 0 to ${mag.capacity}.`);
  }
  const selectedAmmoId = String(ammoId ?? mag.loadedAmmoId ?? "");
  const ammo = odyssey.ammunition.find((a) => a.id === selectedAmmoId);
  const weapon = magazineWeapon(odyssey, id);
  if ((rounds > 0 || selectedAmmoId) && !ammo) throw new Error("Choose a valid ammunition type.");
  if (weapon && !Array.isArray(weapon.acceptedMagazineIds) && selectedAmmoId && !weapon.ammoIds.includes(selectedAmmoId)) {
    throw new Error("This ammunition is incompatible with the mounted weapon.");
  }
  mag.loaded = rounds;
  mag.loadedAmmoId = selectedAmmoId;
  if (weapon) {
    if (Array.isArray(weapon.acceptedMagazineIds)) weapon.ammoIds = selectedAmmoId ? [selectedAmmoId] : [];
    weapon.loaded = rounds;
    weapon.loadedAmmoId = selectedAmmoId;
  }
}
function emptyMagazine(odyssey, id) {
  const mag = odyssey.inventory.magazines.find((m) => m.id === id);
  if (!mag) throw new Error("Magazine no longer exists.");
  if (magazineWeapon(odyssey, id)) throw new Error("Detach the magazine first.");
  const ammo = odyssey.ammunition.find((a) => a.id === mag.loadedAmmoId);
  if (mag.loaded && !ammo) throw new Error("The loaded ammunition is missing.");
  if (ammo && ammo.quantity + mag.loaded > 999999) throw new Error("Ammo reserve is full.");
  if (ammo) ammo.quantity += mag.loaded;
  mag.loaded = 0;
  mag.loadedAmmoId = "";
}
function equipItem(tracker, id, partName, gm = true) {
  var _a;
  const item = tracker.odyssey.inventory.items.find((i) => i.id === id);
  if (!item || !["armor", "shield", "implant"].includes(item.kind)) throw new Error("Choose equipment.");
  if (item.kind === "implant" && !gm && !item.playerRemovable) throw new Error("Only the GM can equip this implant.");
  if (item.quantity < 1 || item.equippedPart) throw new Error("Item is unavailable or already equipped.");
  const part = tracker.body[partName];
  if (!part) throw new Error("Body part no longer exists.");
  if (!equipmentFits(item, partName, part)) throw new Error("This item does not fit the selected slot.");
  if (item.kind !== "implant" && tracker.odyssey.inventory.items.some((i) => ["armor", "shield"].includes(i.kind) && i.equippedPart === partName)) throw new Error("Armor is already equipped on this part.");
  if (item.quantity > 1) {
    if (tracker.odyssey.inventory.items.length >= 1e3) throw new Error("Inventory is full.");
    tracker.odyssey.inventory.items.push({ ...structuredClone(item), id: newId(), quantity: item.quantity - 1, equippedPart: "", previousArmor: 0, legacyEquipment: false });
    item.quantity = 1;
  }
  (_a = tracker.odyssey.inventory).baseArmor ?? (_a.baseArmor = {});
  if (!Object.hasOwn(tracker.odyssey.inventory.baseArmor, partName)) tracker.odyssey.inventory.baseArmor[partName] = part.armor;
  item.equippedPart = partName;
  item.legacyEquipment = false;
  reconcileEquipment(tracker);
}
function unequipItem(tracker, id, gm = true) {
  const item = tracker.odyssey.inventory.items.find((i) => i.id === id);
  if (!item) throw new Error("Item no longer exists.");
  if (item.kind === "implant" && !gm && !item.playerRemovable) throw new Error("Only the GM can remove this implant.");
  const part = tracker.body[item.equippedPart], inv = tracker.odyssey.inventory;
  if (part && Object.hasOwn(inv.baseArmor || {}, item.equippedPart)) {
    const expected = inv.baseArmor[item.equippedPart] + inv.items.filter((i) => i.equippedPart === item.equippedPart && ["armor", "shield", "implant"].includes(i.kind) && i.quantity > 0 && (i.kind === "implant" || !i.criticalMax || i.criticalDamage < i.criticalMax)).reduce((sum, i) => sum + i.armor, 0);
    if (part.armor !== expected) inv.baseArmor[item.equippedPart] = part.armor;
  }
  item.equippedPart = "";
  reconcileEquipment(tracker);
}
function consumeItem(tracker, id) {
  const item = tracker.odyssey.inventory.items.find((i) => i.id === id);
  if (!item || !["consumable", "medical", "resource"].includes(item.kind) || item.quantity < 1) throw new Error("No usable item available.");
  item.quantity -= 1;
}
function findInventoryEntry(tracker, type, id) {
  const o = tracker.odyssey;
  if (type === "ammo") return o.ammunition.find((a) => a.id === id);
  if (type === "magazine") return o.inventory.magazines.find((m) => m.id === id);
  if (type?.startsWith("weapon-")) return o.weapons[type.slice(7)]?.find((w, index) => (w.id || String(index)) === id);
  return o.inventory.items.find((i) => i.id === id);
}
function deleteInventoryEntry(tracker, type, id) {
  const o = tracker.odyssey;
  if (type === "ammo") {
    if (o.inventory.magazines.some((m) => m.loadedAmmoId === id && m.loaded) || o.weapons.ranged.some((w) => w.acceptsMagazines === false && w.loadedAmmoId === id && w.loaded)) throw new Error("Empty all magazines using this ammo first.");
    o.ammunition = o.ammunition.filter((a) => a.id !== id);
    for (const w of o.weapons.ranged) w.ammoIds = w.ammoIds.filter((a) => a !== id);
    return;
  }
  if (type === "magazine") {
    const removed = o.inventory.magazines.find((m) => m.id === id), replacement = removed && o.inventory.magazines.find((m) => m.id !== id && m.name === removed.name && m.capacity === removed.capacity);
    const index = o.weapons.ranged.findIndex((w) => w.magazineId === id);
    if (index >= 0) swapMagazine(o, index, "");
    emptyMagazine(o, id);
    o.inventory.magazines = o.inventory.magazines.filter((m) => m.id !== id);
    for (const w of o.weapons.ranged) if (w.acceptedMagazineIds) w.acceptedMagazineIds = [...new Set(w.acceptedMagazineIds.flatMap((m) => m === id ? replacement ? [replacement.id] : [] : [m]))];
    return;
  }
  if (type?.startsWith("weapon-")) {
    const kind = type.slice(7), index = o.weapons[kind].findIndex((w, j) => (w.id || String(j)) === id);
    if (index < 0) throw new Error("Weapon no longer exists.");
    if (kind === "ranged") {
      const w = o.weapons.ranged[index];
      if (w.acceptsMagazines === false && w.loaded) {
        const a = o.ammunition.find((a2) => a2.id === w.loadedAmmoId);
        if (!a || a.quantity + w.loaded > 999999) throw new Error("Cannot return loaded ammunition to reserve.");
        a.quantity += w.loaded;
      } else if (w.magazineId) swapMagazine(o, index, "");
    }
    o.weapons[kind].splice(index, 1);
    return;
  }
  const item = o.inventory.items.find((i) => i.id === id);
  if (!item) throw new Error("Item no longer exists.");
  if (item.equippedPart) unequipItem(tracker, id, true);
  o.inventory.items = o.inventory.items.filter((i) => i.id !== id);
}
function saveInventoryEntry(tracker, origin, values, id) {
  const o = tracker.odyssey, existing = origin ? findInventoryEntry(tracker, origin.type, origin.id) : null;
  if (origin && !existing) throw new Error("Item no longer exists.");
  const kind = values.kind, weaponKind = values.weaponKind === "ranged" ? "ranged" : "melee";
  const dest = kind === "weapon" ? "weapon-" + weaponKind : kind === "ammo" ? "ammo" : kind === "magazine" ? "magazine" : "item";
  if (origin && origin.type !== dest) deleteInventoryEntry(tracker, origin.type, origin.id);
  const same = origin?.type === dest ? existing : null;
  if (kind === "ammo") {
    const row2 = { id, name: values.name, damage: Number(values.damage) || 0, penetration: Number(values.penetration) || 0, quantity: Number(values.quantity) || 0 };
    if (same) Object.assign(same, row2);
    else {
      if (o.ammunition.length >= 200) throw new Error("Too many ammo types.");
      o.ammunition.push(row2);
    }
    return;
  }
  if (kind === "magazine") {
    const capacity = Number(values.capacity), loaded = Number(values.loaded);
    if (!Number.isInteger(capacity) || capacity < 1 || capacity > 999 || !Number.isInteger(loaded) || loaded < 0 || loaded > capacity) throw new Error("Invalid magazine capacity or rounds.");
    const mounted = same && magazineWeapon(o, same.id);
    if (mounted && capacity !== mounted.capacity) throw new Error("Detach the magazine before changing its capacity.");
    if (same) {
      same.capacity = capacity;
      same.name = values.name;
      setMagazineRounds(o, same.id, loaded, values.loadedAmmoId);
    } else {
      const quantity = Number(values.quantity ?? 1);
      if (!Number.isInteger(quantity) || quantity < 1 || quantity > 1e3 || o.inventory.magazines.length + quantity > 1e3) throw new Error("Invalid magazine quantity or inventory is full.");
      for (let i = 0; i < quantity; i++) {
        const magazineId = i ? id + "-copy-" + i : id;
        o.inventory.magazines.push({ id: magazineId, name: values.name, capacity, loaded: 0, loadedAmmoId: "" });
        setMagazineRounds(o, magazineId, loaded, values.loadedAmmoId);
      }
    }
    return;
  }
  if (kind === "weapon") {
    if ([...o.weapons.melee, ...o.weapons.ranged].some((w) => w !== same && w.name === values.name)) throw new Error("A weapon with this name already exists.");
    const row2 = { id, name: values.name, accuracy: Number(values.accuracy) || 0, damage: Number(values.damage) || 0 };
    if (weaponKind === "ranged") {
      row2.weaponCategory = normalizeWeaponCategory(values.weaponCategory);
      row2.rangeReduction = normalizeRangeReduction(values.rangeReduction);
      const acceptsMagazines = values.acceptsMagazines !== false;
      if (same && same.acceptsMagazines !== false !== acceptsMagazines && same.loaded > 0) throw new Error("Empty or detach the current magazine before changing magazine type.");
      if (acceptsMagazines) {
        const acceptedMagazineIds = [...new Set(values.acceptedMagazineIds || [])];
        if (acceptedMagazineIds.some((id2) => !o.inventory.magazines.some((m) => m.id === id2))) throw new Error("Selected magazine no longer exists.");
        if (same?.magazineId && !magazineCompatible(o, { ...same, acceptedMagazineIds }, o.inventory.magazines.find((m) => m.id === same.magazineId))) throw new Error("Detach the current magazine before removing its compatibility.");
        Object.assign(row2, { acceptsMagazines: true, acceptedMagazineIds, capacity: same?.magazineId ? same.capacity : 0, ammoIds: same?.magazineId ? same.ammoIds || [] : [], loaded: same?.loaded || 0, loadedAmmoId: same?.loadedAmmoId || "", magazineId: same?.magazineId || "" });
      } else {
        const capacity = Number(values.capacity), ammoIds = values.ammoIds || [], loaded = Number(values.loaded ?? same?.loaded ?? 0), loadedAmmoId = String(values.loadedAmmoId ?? same?.loadedAmmoId ?? "");
        if (!Number.isInteger(capacity) || capacity < 1 || capacity > 999 || !Number.isInteger(loaded) || loaded < 0 || loaded > capacity) throw new Error("Invalid built-in magazine capacity or rounds.");
        if (loaded && (!ammoIds.includes(loadedAmmoId) || !o.ammunition.some((a) => a.id === loadedAmmoId))) throw new Error("Choose compatible ammunition for loaded rounds.");
        Object.assign(row2, { acceptsMagazines: false, acceptedMagazineIds: [], capacity, ammoIds, loaded, loadedAmmoId: loaded ? loadedAmmoId : "", magazineId: "" });
      }
    }
    if (same) Object.assign(same, row2);
    else {
      if (o.weapons[weaponKind].length >= 20) throw new Error("Weapon limit is 20 per type.");
      o.weapons[weaponKind].push(row2);
    }
    return;
  }
  if (same?.equippedPart && !["armor", "shield", "implant"].includes(kind)) unequipItem(tracker, id, true);
  const row = { ...values, id, quantity: Number(values.quantity) || 0, armor: Number(values.armor) || 0, healCritical: Number(values.healCritical) || 0, criticalMax: Number(values.criticalMax) || 0, criticalDamage: Number(values.criticalDamage) || 0, playerRemovable: values.playerRemovable === true, legacyEquipment: false, equippedPart: same?.equippedPart || "", previousArmor: same?.previousArmor || 0 };
  if (!["armor", "shield", "implant"].includes(kind)) {
    row.armor = 0;
    row.criticalMax = 0;
    row.criticalDamage = 0;
  }
  if (kind === "implant") row.criticalDamage = 0;
  if (row.equippedPart && row.quantity > 1 && o.inventory.items.length >= 1e3) throw new Error("Inventory is full; cannot separate equipped and spare items.");
  if (row.criticalDamage > row.criticalMax) throw new Error("Critical damage cannot exceed durability.");
  if (same?.equippedPart && !equipmentFits(row, same.equippedPart, tracker.body[same.equippedPart])) throw new Error("Unequip before changing the equipment slot.");
  if (same) Object.assign(same, row);
  else {
    if (o.inventory.items.length >= 1e3) throw new Error("Inventory limit is 1000 items.");
    o.inventory.items.push(row);
  }
  reconcileEquipment(tracker);
}

// ui/characterView.js
var esc = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
var bonusText = (b) => [...Object.entries(b?.attributes || {}), ...Object.entries(b?.skills || {}), ...RANGE_BANDS.filter(([key]) => b?.rangeReduction?.[key]).map(([key, label]) => [label.split(" \xB7 ")[0] + " relief", b.rangeReduction[key]])].map(([k, v]) => `${k} ${v > 0 ? "+" : ""}${v}`).join(" \xB7 ");
function attributeCards(o, fields, editable) {
  return `<div class="cp-screen"><div class="cp-row cp-roll-options"><label>Roll modifier <input type="number" value="0" data-profile-modifier="attribute" aria-label="Characteristic roll modifier"></label></div><div class="cp-attrs">${fields.map(([k, label]) => {
    const delta = o.modifiers?.attributes?.[k] || 0;
    return `<div class="cp-attr-wrap"><button type="button" class="cp-attr" data-action="profile-roll-char" data-attribute="${esc(k)}" aria-label="Roll ${esc(label)}"><span class="cp-attr-val">${o.attributes[k] ?? 0}</span><span class="cp-attr-code">${esc(label)}</span>${delta ? `<span class="cp-mod ${delta > 0 ? "good" : "bad"}">${delta > 0 ? "+" : ""}${delta}</span>` : ""}</button>${editable ? `<label class="cp-edit-value">Base <input type="number" min="0" max="20" value="${o.baseAttributes?.[k] ?? o.attributes[k] ?? 0}" data-action="set-odyssey-attribute" data-attribute="${esc(k)}" aria-label="Edit ${esc(label)}"></label>` : ""}</div>`;
  }).join("")}</div></div>`;
}
function skillCards(o, groups, editable) {
  const passives = Object.entries(ARMOR_TRAINING).filter(([k]) => Object.hasOwn(o.skills, k) && (editable || o.skills[k] === 1));
  return `<div class="cp-screen"><div class="cp-row cp-roll-options"><label>Roll modifier <input type="number" value="0" data-profile-modifier="skill" aria-label="Skill roll modifier"></label><button type="button" class="secondary" data-action="profile-random-skill">d100 vs d100</button>${editable ? '<button type="button" class="secondary" data-od-action="add-skill">+ Skill</button>' : ""}</div>${groups.filter(([, entries]) => entries.length).map(([label, entries]) => `<div class="cp-section-title">${esc(label)}</div><div class="cp-list">${entries.map(([k, value]) => {
    const name = o.skillLabels[k] || k, perks = o.skillPerks[k] || [], mod = o.modifiers?.skills?.[k] || 0;
    return `<div class="cp-skill-row"><button class="cp-card cp-skill-card" type="button" data-action="profile-roll-skill" data-skill="${esc(k)}" aria-label="Roll ${esc(name)}"><span><b>${esc(name)}</b>${o.skillStrengthBonuses[k] ? '<span class="cp-pill">STR</span>' : ""}<span class="cp-perks">${RANGE_BANDS.filter(([band]) => o.skillRangeReduction?.[k]?.[band]).map(([band, label2]) => `<span class="cp-pill">${esc(label2.split(" \xB7 ")[0])} relief ${o.skillRangeReduction[k][band]}</span>`).join("")}${perks.map((p) => `<span class="cp-pill">${esc(p)}</span>`).join("")}</span></span><span class="cp-skill-level"><span class="cp-pips">${Array.from({ length: 10 }, (_, i) => `<i class="cp-pip ${i < value ? "on" : ""}"></i>`).join("")}</span><strong>${value}/10</strong>${mod ? `<span class="cp-mod ${mod > 0 ? "good" : "bad"}">${mod > 0 ? "+" : ""}${mod}</span>` : ""}</span></button>${editable && !o.grantedSkills?.[k] ? `<button type="button" class="cp-pencil" data-od-action="edit-skill" data-skill="${esc(k)}" aria-label="Edit ${esc(name)}">\u270E</button>` : ""}</div>`;
  }).join("")}</div>`).join("")}${passives.length ? `<div class="cp-passives cp-row"><span class="cp-muted">Armor training</span>${passives.map(([k, label]) => `<span class="cp-pill">${esc(label)} ${o.skills[k]}</span>`).join("")}${editable ? '<button type="button" class="cp-pencil" data-od-action="edit-training" aria-label="Edit armor training">\u270E</button>' : ""}</div>` : ""}</div>`;
}
function effectChips(o, gm) {
  const rows2 = o.effectRows || [];
  return `<div class="cp-row cp-effects">${rows2.map((e) => `<span class="cp-effect" title="${esc(bonusText(e.bonuses))}"><b>${esc(e.name)}</b> ${esc(bonusText(e.bonuses))}${gm && !e.passive ? `<button type="button" data-od-action="remove-effect" data-id="${esc(e.id)}" aria-label="Remove ${esc(e.name)}">\xD7</button>` : ""}</span>`).join("")}${o.armorSummary?.total ? `<span class="cp-pill">${esc(o.armorSummary.name)} \xB7 ${o.armorSummary.total} \xB7 ${o.armorSummary.trained ? "trained" : "untrained"}</span>` : ""}</div>`;
}
function inventoryView(tracker, gm) {
  const o = tracker.odyssey, inv = o.inventory;
  const edit = (type, id) => gm ? `<button type="button" class="cp-pencil" data-od-action="edit-entry" data-entry-type="${type}" data-id="${esc(id)}" aria-label="Edit item">\u270E</button><button type="button" class="danger cp-pencil" data-od-action="delete-entry" data-entry-type="${type}" data-id="${esc(id)}" aria-label="Remove item">\xD7</button>` : "";
  const renderItem = (i) => {
    const equipment = ["armor", "shield", "implant"].includes(i.kind), broken = i.kind !== "implant" && i.criticalMax > 0 && i.criticalDamage >= i.criticalMax;
    const usable = ["consumable", "medical"].includes(i.kind), subtype = i.kind === "medical" ? "medkit" : i.subtype;
    let controls = "";
    if (equipment) {
      const label = i.equippedPart ? '<span class="cp-muted">' + esc(tracker.body[i.equippedPart]?.label || i.equippedPart) + "</span>" : "";
      if (!gm && i.kind === "implant" && !i.playerRemovable) controls = label + '<span class="cp-pill">GM installation / removal only</span>';
      else if (i.equippedPart) controls = label + '<button type="button" data-od-action="unequip" data-id="' + esc(i.id) + '">Unequip</button>';
      else {
        const options = Object.entries(tracker.body).filter(([key, p]) => (gm || !p.hidden) && equipmentFits(i, key, p) && (i.kind === "implant" || !inv.items.some((other) => ["armor", "shield"].includes(other.kind) && other.equippedPart === key)));
        controls = '<label>Slot <select data-equip-part="' + esc(i.id) + '">' + options.map(([key, p]) => '<option value="' + esc(key) + '">' + esc(p.label || key) + "</option>").join("") + '</select></label><button type="button" data-od-action="equip" data-id="' + esc(i.id) + '" ' + (!i.quantity || !options.length ? "disabled" : "") + ">Equip</button>";
      }
    } else if (usable) controls = '<button type="button" data-od-action="use-item" data-id="' + esc(i.id) + '" ' + (!i.quantity || subtype === "grenade" ? "disabled" : "") + ">" + (subtype === "grenade" ? "Grenade \xB7 placeholder" : "Use\u2026") + "</button>";
    return '<article class="cp-card"><div class="cp-row"><b>' + esc(i.name) + '</b><span class="cp-pill">' + esc(usable ? subtype : i.kind) + '</span><span class="cp-pill">\xD7' + i.quantity + "</span>" + (i.armor ? '<span class="cp-pill">Armor ' + i.armor + "</span>" : "") + (i.kind === "implant" && i.criticalMax ? '<span class="cp-pill">Body HP +' + i.criticalMax + "</span>" : i.criticalMax ? '<span class="cp-pill ' + (broken ? "bad" : "") + '">Durability ' + (i.criticalMax - i.criticalDamage) + "/" + i.criticalMax + "</span>" : "") + edit("item", i.id) + "</div>" + (i.description ? '<p class="cp-muted cp-description">' + esc(i.description) + "</p>" : "") + (bonusText(i.bonuses) ? '<p class="cp-muted">' + esc(bonusText(i.bonuses)) + "</p>" : "") + '<div class="cp-row">' + controls + "</div></article>";
  };
  const equippedItems = inv.items.filter((i) => i.equippedPart).map(renderItem).join("");
  const items = inv.items.filter((i) => !i.equippedPart).map(renderItem).join("");
  const ammo = o.ammunition.map((a) => `<article class="cp-card cp-row"><b>${esc(a.name)}</b><span class="cp-pill">Ammo \xD7${a.quantity}</span><span class="cp-muted">DMG ${a.damage} \xB7 AP ${a.penetration}</span>${edit("ammo", a.id)}</article>`).join("");
  const mags = inv.magazines.map((m) => {
    const w = o.weapons.ranged.find((w2) => w2.magazineId === m.id), a = o.ammunition.find((a2) => a2.id === m.loadedAmmoId);
    return `<article class="cp-card"><div class="cp-row"><b>${esc(m.name)}</b><span class="cp-pill">Magazine ${m.loaded}/${m.capacity}</span><span class="cp-muted">${esc(a?.name || "Empty")}${w ? " \xB7 " + esc(w.name) : ""}</span>${gm ? `<button type="button" class="secondary" data-od-action="set-mag-rounds" data-id="${esc(m.id)}">GM \xB7 Set rounds\u2026</button>` : ""}${edit("magazine", m.id)}</div>${gm && !w ? `<div class="cp-row"><select aria-label="Magazine ammunition" data-mag-ammo="${esc(m.id)}">${o.ammunition.map((a2) => `<option value="${esc(a2.id)}">${esc(a2.name)} (${a2.quantity})</option>`).join("")}</select><button type="button" data-od-action="fill-mag" data-id="${esc(m.id)}">Fill</button><button type="button" data-od-action="empty-mag" data-id="${esc(m.id)}">Empty to reserve</button></div>` : ""}</article>`;
  }).join("");
  const weapons = ["melee", "ranged"].flatMap((kind) => o.weapons[kind].map((w, index) => {
    let controls = "";
    if (gm && kind === "ranged") {
      if (w.acceptsMagazines !== false) {
        const compatible = inv.magazines.filter((m) => m.id === w.magazineId || !o.weapons.ranged.some((x) => x.magazineId === m.id) && magazineCompatible(o, w, m));
        controls = '<div class="cp-row"><label>Magazine<select data-weapon-mag="' + index + '"><option value="">Detached</option>' + compatible.map((m) => '<option value="' + esc(m.id) + '" ' + (m.id === w.magazineId ? "selected" : "") + ">" + esc(m.name) + " \xB7 " + m.loaded + "/" + m.capacity + "</option>").join("") + '</select></label><button type="button" data-od-action="mount-mag" data-index="' + index + '">Apply magazine</button></div>';
      } else controls = '<div class="cp-row"><label>Ammo<select data-reload-kind="' + index + '">' + o.ammunition.filter((a) => w.ammoIds.includes(a.id)).map((a) => '<option value="' + esc(a.id) + '" ' + (a.id === w.loadedAmmoId ? "selected" : "") + ">" + esc(a.name) + " (" + a.quantity + ")</option>").join("") + '</select></label><button type="button" data-od-action="reload-weapon" data-index="' + index + '">Reload built-in magazine</button></div>';
    }
    return '<article class="cp-card"><div class="cp-row"><b>' + esc(w.name) + '</b><span class="cp-pill">' + esc(kind === "ranged" ? weaponCategoryLabel(w.weaponCategory) : kind) + '</span><span class="cp-muted">ACC ' + w.accuracy + (kind === "melee" ? " \xB7 DMG " + w.damage : " \xB7 " + (w.acceptsMagazines === false ? "Built-in " : "Magazine ") + w.loaded + "/" + w.capacity) + "</span>" + (Object.keys(w.rangeReduction || {}).length ? '<span class="cp-muted">' + esc(bonusText({ rangeReduction: w.rangeReduction })) + "</span>" : "") + edit("weapon-" + kind, w.id || String(index)) + "</div>" + controls + "</article>";
  }));
  const group = (label, html) => html ? `<div class="cp-section-title">${label}</div><div class="cp-list">${html}</div>` : "";
  return `<div class="cp-screen">${gm ? '<div class="cp-row"><button type="button" data-od-action="add-item">+ Item</button><button type="button" data-od-action="open-item-pool">Campaign item pool</button></div>' : ""}${group("Equipped", equippedItems)}${group("Items", items)}${group("Weapons", weapons.join(""))}${group("Ammunition", ammo)}${group("Magazines", mags)}${!equippedItems && !items && !ammo && !mags && !weapons.length ? '<p class="cp-muted">Inventory is empty.</p>' : ""}</div>`;
}
function templatesView(library, gm) {
  if (!gm) return "";
  return `<div class="cp-screen cp-templates"><div class="cp-section-title">Character Templates</div><p class="cp-muted">Templates are saved in this browser. Export JSON for backup, or save a library scene to your Owlbear account. Account save opens the Assets folder picker; it creates a snapshot.</p><div class="cp-row"><button type="button" data-od-action="save-template">Save character as template</button><button type="button" data-od-action="export-library">Export JSON</button><button type="button" data-od-action="import-library">Import JSON</button><button type="button" data-od-action="account-save">Save to account</button><button type="button" data-od-action="account-load">Load from account</button></div><div class="cp-list">${library.templates.map((t) => `<article class="cp-card cp-row"><b>${esc(t.name)}</b>${gm ? `<button type="button" data-od-action="apply-template" data-id="${esc(t.id)}">Apply to token\u2026</button>` : ""}<button type="button" class="secondary" data-od-action="rename-template" data-id="${esc(t.id)}">Rename</button><button type="button" class="secondary" data-od-action="duplicate-template" data-id="${esc(t.id)}">Duplicate</button><button type="button" class="danger" data-od-action="delete-template" data-id="${esc(t.id)}">Remove</button></article>`).join("") || '<p class="cp-muted">No templates saved.</p>'}</div></div>`;
}

// ui/compactHud.js
function compactHud({ token, tracker, floating, target, targetData, skills, weapons, draft: draft2, parts, ammunition }) {
  return `<section class="od-hud cp-screen"><div class="cp-row"><div><span class="cp-section-title">Character</span><h3>${esc(token.name || "Character")}</h3></div><span class="cp-pill">${Object.entries(tracker.body).filter(([k]) => !["Shield", "Special"].includes(k)).reduce((n2, [, p]) => n2 + p.current, 0)} HP</span><button type="button" data-action="${floating ? "close-character-hud" : "open-character-hud"}" class="secondary">${floating ? "Close HUD" : "Floating HUD"}</button></div><div class="od-hud-grid"><div class="cp-card"><div class="cp-section-title">Weapon & skill</div><label>Weapon<select data-hud-attack-field="weaponName">${weapons}</select></label><label>Skill<select data-hud-attack-field="skill">${skills}</select></label>${ammunition.replaceAll("data-attack-field=", "data-hud-attack-field=")}</div><div class="cp-card"><div class="cp-section-title">Target</div><div class="cp-row"><b>${esc(target?.name || "No target selected")}</b><span data-target-distance class="od-target-distance"></span><span data-range-modifier="attack" class="od-range-readout cp-muted"></span><button type="button" data-action="pick-attack-target">Pick on map</button></div><label>Body zone<select data-hud-attack-field="targetPart">${parts.map(([key, label]) => `<option value="${esc(key)}" ${key === draft2.targetPart ? "selected" : ""}>${esc(label)}</option>`).join("")}</select></label><p class="cp-muted">${targetData ? "Target selected \xB7 V1 armor, Parry and damage rules" : "Select a visible target to attack."}</p></div><div class="cp-card"><div class="cp-section-title">Action</div><label>Attack bonus<input type="number" value="${Number(draft2.attackBonuses) || 0}" data-hud-attack-field="attackBonuses"></label><label>Attack penalty<input type="number" value="${Number(draft2.attackPenalties) || 0}" data-hud-attack-field="attackPenalties"></label><button type="button" data-action="perform-attack" ${!target ? "disabled" : ""}>Attack</button></div></div><p class="cp-muted" role="status">${esc(tracker.lastRoll?.summary || "No rolls yet.")}</p></section>`;
}

// itemPool.js
var ITEM_POOL_PREFIX = "com.codex.body-hp/item-pool/";
var ITEM_POOL_FORMAT = "odyssey-item-pool";
var TYPES = /* @__PURE__ */ new Set(["item", "ammo", "magazine", "weapon-melee", "weapon-ranged"]);
var stable = (value) => Array.isArray(value) ? value.map(stable) : value && typeof value === "object" ? Object.fromEntries(Object.keys(value).sort().map((key) => [key, stable(value[key])])) : value;
var blueprintText = (blueprint) => JSON.stringify(stable(blueprint));
var ammoDefinition = (a) => ({ name: a.name, damage: a.damage, penetration: a.penetration });
var sameAmmo = (a, b) => a.name === b.name && a.damage === b.damage && a.penetration === b.penetration;
var sameMagazine = (a, b) => a.name === b.name && a.capacity === b.capacity;
function uniqueDefinitions(rows2, project) {
  return [...new Map(rows2.map((row) => {
    const value = project(row);
    return [blueprintText(value), value];
  })).entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, value]) => value);
}
function captureBlueprint(tracker, type, id) {
  const o = tracker.odyssey, entry = findInventoryEntry(tracker, type, id);
  if (!entry) throw new Error("Item no longer exists.");
  let values;
  const dependencies = { ammunition: [], magazines: [] };
  if (type === "ammo") values = { kind: "ammo", ...ammoDefinition(entry), quantity: 1 };
  else if (type === "magazine") values = { kind: "magazine", name: entry.name, quantity: 1, capacity: entry.capacity, loaded: 0, loadedAmmoId: "" };
  else if (type.startsWith("weapon-")) {
    const kind = type.slice(7);
    values = { kind: "weapon", weaponKind: kind, name: entry.name, accuracy: entry.accuracy, damage: entry.damage, ...kind === "ranged" && entry.weaponCategory ? { weaponCategory: entry.weaponCategory } : {}, ...kind === "ranged" && Object.values(entry.rangeReduction || {}).some(Boolean) ? { rangeReduction: structuredClone(entry.rangeReduction) } : {} };
    if (kind === "ranged") {
      values.acceptsMagazines = entry.acceptsMagazines !== false;
      values.loaded = 0;
      values.loadedAmmoId = "";
      values.magazineId = "";
      if (values.acceptsMagazines) {
        let magazines = o.inventory.magazines.filter((m) => m.id === entry.magazineId || magazineCompatible(o, entry, m));
        if (!magazines.length && entry.capacity) magazines = [{ name: entry.name + " magazine", capacity: entry.capacity }];
        dependencies.magazines = uniqueDefinitions(magazines, (m) => ({ name: m.name, capacity: m.capacity })).map((m, index) => ({ id: "mag-" + index, ...m }));
        values.acceptedMagazineIds = dependencies.magazines.map((m) => m.id);
      } else {
        dependencies.ammunition = uniqueDefinitions(o.ammunition.filter((a) => entry.ammoIds.includes(a.id)), ammoDefinition).map((a, index) => ({ id: "ammo-" + index, ...a }));
        values.capacity = entry.capacity;
        values.ammoIds = dependencies.ammunition.map((a) => a.id);
      }
    }
  } else {
    values = {};
    for (const key of ["kind", "name", "description", "armor", "criticalMax", "playerRemovable", "slots", "subtype", "healCritical", "abilities", "bonuses"]) if (entry[key] !== void 0) values[key] = structuredClone(entry[key]);
    values.quantity = 1;
    values.criticalDamage = 0;
    values.slots ?? (values.slots = [entry.slot || "any"]);
    if (values.bonuses) {
      const skills = {};
      for (const [key, value] of Object.entries(values.bonuses.skills || {})) {
        const label = o.skillLabels[key] || key;
        skills[label] = (skills[label] || 0) + value;
      }
      values.bonuses.skills = skills;
    }
    if (values.abilities) values.abilities = values.abilities.map((ability, index) => ({ ...ability, id: "ability-" + index }));
  }
  return { format: "odyssey-item-blueprint", version: 1, type, values, dependencies };
}
function validateBlueprint(raw) {
  if (raw?.format !== "odyssey-item-blueprint" || raw.version !== 1 || !TYPES.has(raw.type) || !raw.values || typeof raw.values.name !== "string" || !raw.values.name.trim() || raw.values.name.length > 120) throw new Error("Invalid item pool definition.");
  if (!["item", "consumable", "medical", "tool", "resource", "armor", "shield", "implant", "ammo", "magazine", "weapon"].includes(raw.values.kind)) throw new Error("Invalid item type.");
  const expected = raw.type.startsWith("weapon-") ? "weapon" : raw.type === "item" ? null : raw.type;
  if (raw.type === "item" && ["weapon", "ammo", "magazine"].includes(raw.values.kind)) throw new Error("Item type does not match its definition.");
  if (expected && raw.values.kind !== expected) throw new Error("Item type does not match its definition.");
  for (const key of ["ammunition", "magazines"]) if (!Array.isArray(raw.dependencies?.[key]) || raw.dependencies[key].length > 200 || raw.dependencies[key].some((d) => !d || typeof d.id !== "string" || typeof d.name !== "string")) throw new Error("Invalid item dependencies.");
  if (blueprintText(raw).length > 2e5) throw new Error("Item definition is too large.");
  return structuredClone(raw);
}
async function blueprintId(blueprint) {
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(blueprintText(blueprint)));
  return [...new Uint8Array(hash)].map((n2) => n2.toString(16).padStart(2, "0")).join("");
}
function poolEntries(metadata) {
  const entries = [];
  let errors2 = 0;
  for (const [key, value] of Object.entries(metadata || {})) if (key.startsWith(ITEM_POOL_PREFIX) && value && !value.deleted) {
    try {
      entries.push({ id: key.slice(ITEM_POOL_PREFIX.length), blueprint: validateBlueprint(value.blueprint) });
    } catch {
      errors2++;
    }
  }
  return { entries: entries.sort((a, b) => a.blueprint.values.name.localeCompare(b.blueprint.values.name)), errors: errors2 };
}
function parsePoolJSON(raw) {
  if (raw?.format !== ITEM_POOL_FORMAT || raw.version !== 1 || !Array.isArray(raw.entries) || raw.entries.length > 2e3) throw new Error("Choose an Odyssey item pool JSON (version 1).");
  return raw.entries.map(validateBlueprint);
}
function exportPool(entries) {
  return { format: ITEM_POOL_FORMAT, version: 1, entries: entries.map((e) => e.blueprint) };
}
function copyBlueprint(tracker, raw, quantity = 1, id = newId()) {
  const blueprint = validateBlueprint(raw), o = tracker.odyssey, values = structuredClone(blueprint.values), ammoIds = /* @__PURE__ */ new Map(), magazineIds = /* @__PURE__ */ new Map();
  if (!Number.isInteger(quantity) || quantity < 1 || quantity > 999999) throw new Error("Invalid item quantity.");
  for (const [index, definition] of blueprint.dependencies.ammunition.entries()) {
    let ammo = o.ammunition.find((a) => sameAmmo(a, definition));
    if (!ammo) {
      if (o.ammunition.length >= 200) throw new Error("Ammunition type limit reached.");
      ammo = { ...definition, id: id + "-ammo-" + index, quantity: 0 };
      o.ammunition.push(ammo);
    }
    ammoIds.set(definition.id, ammo.id);
  }
  for (const [index, definition] of blueprint.dependencies.magazines.entries()) {
    let matches = o.inventory.magazines.filter((m) => sameMagazine(m, definition));
    if (!matches.length) {
      const magazine = { id: id + "-mag-" + index, name: definition.name, capacity: definition.capacity, loaded: 0, loadedAmmoId: "" };
      if (o.inventory.magazines.length >= 1e3) throw new Error("Magazine inventory is full.");
      o.inventory.magazines.push(magazine);
      matches = [magazine];
    }
    magazineIds.set(definition.id, matches.map((m) => m.id));
  }
  values.equippedPart = "";
  values.previousArmor = 0;
  values.legacyEquipment = false;
  values.criticalDamage = 0;
  if (blueprint.type.startsWith("weapon-")) values.weaponKind = blueprint.type.slice(7);
  if (blueprint.type === "magazine" || blueprint.type.startsWith("weapon-")) {
    values.loaded = 0;
    values.loadedAmmoId = "";
    values.magazineId = "";
  }
  if (values.ammoIds) values.ammoIds = values.ammoIds.map((key) => ammoIds.get(key)).filter(Boolean);
  if (values.acceptedMagazineIds) values.acceptedMagazineIds = values.acceptedMagazineIds.flatMap((key) => magazineIds.get(key) || []);
  if (values.bonuses?.skills) {
    const skills = {};
    for (const [label, value] of Object.entries(values.bonuses.skills)) {
      const key = Object.keys(o.skills).find((key2) => key2.toLowerCase() === label.toLowerCase() || (o.skillLabels[key2] || key2).toLowerCase() === label.toLowerCase()) || label;
      skills[key] = (skills[key] || 0) + value;
    }
    values.bonuses.skills = skills;
  }
  if (blueprint.type === "ammo") {
    const existing = o.ammunition.find((a) => sameAmmo(a, values));
    if (existing) {
      if (existing.quantity + quantity > 999999) throw new Error("Ammo reserve is full.");
      existing.quantity += quantity;
      return existing.id;
    }
  }
  if (blueprint.type.startsWith("weapon-")) {
    if (quantity > 20) throw new Error("Weapon limit is 20 per type.");
    for (let i = 0; i < quantity; i++) {
      const base = values.name;
      let name = base, suffix = 2;
      while ([...o.weapons.melee, ...o.weapons.ranged].some((w) => w.name === name)) name = base + " (" + suffix++ + ")";
      saveInventoryEntry(tracker, null, { ...values, name, quantity: 1 }, i ? id + "-copy-" + i : id);
    }
  } else saveInventoryEntry(tracker, null, { ...values, quantity }, id);
  return id;
}

// poolCategories.js
var ARMOR_POOL_TIERS = ["Light Armor", "Medium Armor", "Heavy Armor", "Superheavy Armor", "Outside listed ranges"];
var SLOTS = { Torso: ["Torso", true], torso: ["Torso", true], Head: ["Head", false], head: ["Head", false], "L.Arm": ["Left Arm", false], "left-arm": ["Left Arm", false], "R.Arm": ["Right Arm", false], "right-arm": ["Right Arm", false], "L.Leg": ["Left Leg", false], "left-leg": ["Left Leg", false], "R.Leg": ["Right Leg", false], "right-leg": ["Right Leg", false], Shield: ["Shield", false], Special: ["Special", false], other: ["Other parts", false] };
function armorPoolTier(armor, torso = false) {
  const limits = torso ? [20, 40, 60, 100] : [10, 20, 40, 60];
  const index = limits.findIndex((max) => Number(armor || 0) <= max);
  return ARMOR_POOL_TIERS[index < 0 ? 4 : index];
}
function poolCategoryPaths(blueprint) {
  const v = blueprint.values;
  if (v.kind === "armor" || v.kind === "shield" || v.kind === "implant") {
    let slots = Array.isArray(v.slots) ? v.slots : [v.slot || "any"];
    if (slots.includes("any")) slots = ["Torso", "Head", "L.Arm", "R.Arm", "L.Leg", "R.Leg", "other"];
    const parts = [...new Map(slots.map((key) => {
      const value = SLOTS[key] || [key, false];
      return [value[0], value];
    })).values()];
    return parts.map(([label, torso]) => v.kind === "implant" ? ["Implants", label] : ["Armor", armorPoolTier(v.armor, torso), label]);
  }
  if (v.kind === "consumable" || v.kind === "medical") return [["Usables", { grenade: "Grenades", medkit: "Medkits", repairkit: "Repair kits", stim: "Stims" }[v.subtype] || "Other usables"]];
  if (v.kind === "weapon") return [v.weaponKind === "ranged" ? ["Weapons", "Ranged", weaponCategoryLabel(v.weaponCategory)] : ["Weapons", "Melee"]];
  return [[{ ammo: "Ammunition", magazine: "Magazines", resource: "Resources", tool: "Tools" }[v.kind] || "Other items"]];
}
function groupPoolEntries(entries) {
  const root = { children: /* @__PURE__ */ new Map(), entries: /* @__PURE__ */ new Map() };
  for (const entry of entries) for (const path of poolCategoryPaths(entry.blueprint)) {
    let node = root;
    for (const label of path) {
      if (!node.children.has(label)) node.children.set(label, { label, children: /* @__PURE__ */ new Map(), entries: /* @__PURE__ */ new Map(), all: /* @__PURE__ */ new Set() });
      node = node.children.get(label);
      node.all.add(entry.id);
    }
    node.entries.set(entry.id, entry);
  }
  const order = ["Usables", "Resources", "Ammunition", "Magazines", "Weapons", "Armor", "Implants", "Tools", "Other items", ...ARMOR_POOL_TIERS, "Torso", "Head", "Left Arm", "Right Arm", "Left Leg", "Right Leg", "Shield", "Special", "Other parts"];
  const result = (node) => [...node.children.values()].sort((a, b) => {
    const ai = order.indexOf(a.label), bi = order.indexOf(b.label);
    return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi) || a.label.localeCompare(b.label);
  }).map((child) => ({ label: child.label, count: child.all.size, entries: [...child.entries.values()], children: result(child) }));
  return result(root);
}

// ui/itemPoolView.js
function showItemPool({ getCatalog, character, onGive, onDelete, onExport, onImport }) {
  const dialog = document.createElement("dialog");
  dialog.className = "od-dialog od-pool-dialog";
  dialog.innerHTML = '<div class="cp-row"><h2>Campaign item pool</h2><button type="button" data-pool-close aria-label="Close item pool">\xD7</button></div><p>Give reusable items to ' + esc(character) + '. The pool is shared across scenes in this Owlbear room. Copied equipment is unequipped; magazines are empty.</p><div class="cp-row"><input type="search" data-pool-search placeholder="Search items" aria-label="Search item pool"><button type="button" data-pool-export>Export item pool</button><button type="button" data-pool-import>Import item pool</button></div><div class="cp-row"><button type="button" data-pool-expand title="Expand categories" aria-label="Expand categories">\u25BE</button><button type="button" data-pool-collapse title="Collapse categories" aria-label="Collapse categories">\u25B8</button></div><p data-pool-status role="status"></p><div class="cp-list" data-pool-list></div>';
  document.body.append(dialog);
  let busy = false;
  const expanded = /* @__PURE__ */ new Map();
  const message = dialog.querySelector("[data-pool-status]");
  const refresh = () => {
    const catalog = getCatalog(), search = dialog.querySelector("[data-pool-search]").value.trim().toLowerCase();
    const list = dialog.querySelector("[data-pool-list]");
    for (const details2 of list.querySelectorAll("details[data-pool-category]")) expanded.set(details2.dataset.poolCategory, details2.open);
    const row = (e) => {
      const b = e.blueprint, v = b.values;
      return '<article class="cp-card cp-row"><b>' + esc(v.name) + '</b><span class="cp-pill">' + esc(v.kind) + '</span><span class="cp-muted">' + esc([v.armor ? "Armor " + v.armor : "", v.criticalMax ? "HP " + v.criticalMax : "", v.capacity ? "Capacity " + v.capacity : "", v.damage ? "DMG " + v.damage : ""].filter(Boolean).join(" \xB7 ")) + '</span><button type="button" data-pool-give="' + esc(e.id) + '" ' + (busy ? "disabled" : "") + '>Give item</button><button type="button" data-pool-delete="' + esc(e.id) + '" class="danger" aria-label="Remove definition from pool" ' + (busy ? "disabled" : "") + ">\xD7</button></article>";
    };
    if (search) {
      const words = search.split(/\s+/);
      const matches = catalog.entries.filter((e) => {
        const text2 = [e.blueprint.values.name, e.blueprint.values.kind, e.blueprint.values.description || "", ...poolCategoryPaths(e.blueprint).flat()].join(" ").toLowerCase();
        return words.every((word) => text2.includes(word));
      });
      list.innerHTML = matches.map(row).join("") || '<p class="cp-muted">No matching items.</p>';
    } else {
      const group = (node, path = [], level = 0) => {
        const key = JSON.stringify([...path, node.label]), open = expanded.get(key) ?? level === 0;
        return '<details class="od-pool-category" data-pool-category="' + esc(key) + '" ' + (open ? "open" : "") + "><summary>" + esc(node.label) + ' <span class="cp-pill">' + node.count + '</span></summary><div class="od-pool-category-body">' + node.entries.map(row).join("") + node.children.map((child) => group(child, [...path, node.label], level + 1)).join("") + "</div></details>";
      };
      list.innerHTML = groupPoolEntries(catalog.entries).map((node) => group(node)).join("") || '<p class="cp-muted">The campaign item pool is empty.</p>';
    }
    if (catalog.errors) message.textContent = "Some catalog entries could not be read. Export a backup before importing changes.";
  };
  const run = async (fn) => {
    if (busy) return;
    busy = true;
    refresh();
    try {
      const result = await fn();
      if (result !== false) message.textContent = "Saved.";
    } catch (e) {
      message.textContent = e.message;
    } finally {
      busy = false;
      refresh();
    }
  };
  dialog.querySelector("[data-pool-search]").oninput = refresh;
  dialog.addEventListener("toggle", (event) => {
    if (event.target instanceof HTMLDetailsElement && event.target.isConnected && event.target.dataset.poolCategory) expanded.set(event.target.dataset.poolCategory, event.target.open);
  }, true);
  dialog.addEventListener("click", (event) => {
    const node = event.target.closest("button");
    if (!node) return;
    if (node.hasAttribute("data-pool-expand") || node.hasAttribute("data-pool-collapse")) {
      const open = node.hasAttribute("data-pool-expand");
      for (const details2 of dialog.querySelectorAll("details[data-pool-category]")) {
        details2.open = open;
        expanded.set(details2.dataset.poolCategory, open);
      }
    }
    if (node.dataset.poolGive) void run(() => onGive(node.dataset.poolGive));
    if (node.dataset.poolDelete) void run(() => onDelete(node.dataset.poolDelete));
    if (node.hasAttribute("data-pool-export")) void run(onExport);
    if (node.hasAttribute("data-pool-import")) void run(onImport);
  });
  let finish;
  const closed = new Promise((resolve) => finish = resolve);
  const close = () => {
    if (busy) return;
    dialog.close();
    dialog.remove();
    finish();
  };
  dialog.querySelector("[data-pool-close]").onclick = close;
  dialog.addEventListener("cancel", (event) => {
    event.preventDefault();
    close();
  });
  refresh();
  dialog.showModal();
  return { closed, refresh };
}

// ui/itemEditor.js
var ITEM_TYPES = [["consumable", "Usable"], ["resource", "Resource"], ["ammo", "Ammunition"], ["magazine", "Magazine"], ["weapon", "Weapon"], ["armor", "Armor"], ["shield", "Shield"], ["implant", "Implant"]];
function itemEditor({ existing = {}, type = "resource", skills = [], parts = [], ammunition = [], magazines = [] }) {
  const dialog = document.createElement("dialog");
  dialog.className = "od-dialog od-item-editor";
  const field = (name, label, value = "", inputType = "text", max = 999) => `<label>${esc(label)}<input name="${name}" type="${inputType}" value="${esc(value)}" ${inputType === "number" ? `min="${["damage", "accuracy", "penetration"].includes(name) ? -999 : 0}" max="${max}" step="1"` : 'maxlength="120"'}></label>`;
  const select = (name, label, options, value) => `<label>${esc(label)}<select name="${name}">${options.map(([k, l]) => `<option value="${esc(k)}" ${k === value ? "selected" : ""}>${esc(l)}</option>`).join("")}</select></label>`;
  const checkbox = (name, label, checked) => `<label class="od-check"><input name="${name}" type="checkbox" ${checked ? "checked" : ""}>${esc(label)}</label>`;
  const values = { ...existing, kind: type, acceptsMagazines: existing.acceptsMagazines !== false };
  values.slots ?? (values.slots = [existing.slot || "any"]);
  values.acceptedMagazineIds ?? (values.acceptedMagazineIds = existing.id ? magazines.filter((m) => m.id === existing.magazineId || m.capacity === existing.capacity && (!m.loaded || (existing.ammoIds || []).includes(m.loadedAmmoId))).map((m) => m.id) : []);
  const magazineChoices = [...new Map(magazines.map((m) => [JSON.stringify([m.name, m.capacity]), m])).values()];
  const skillChoices = [...skills];
  for (const key of Object.keys(existing.bonuses?.skills || {})) if (!skillChoices.some(([k]) => k === key)) skillChoices.push([key, key]);
  dialog.innerHTML = `<form><h2>${existing.id ? "Edit item" : "New item"}</h2>${select("kind", "Type", ITEM_TYPES, type)}${field("name", "Name", existing.name || "")}${field("quantity", "Quantity / ammo reserve", existing.quantity ?? 1, "number", 999999)}<div data-type-fields></div><label>Description<textarea name="description" maxlength="2000">${esc(existing.description || "")}</textarea></label><div class="cp-row"><button type="button" data-cancel>Cancel</button><button type="submit">Save</button></div></form>`;
  const form = dialog.querySelector("form"), host = dialog.querySelector("[data-type-fields]");
  const rangeEditor = (prefix, value = {}) => `<fieldset class="od-range-reduction"><legend>Range penalty reduction \xB7 per band</legend>${RANGE_BANDS.map(([key, label]) => field(prefix + key, label, value[key] || 0, "number")).join("")}</fieldset>`;
  function bonusEditor(bonuses = {}, allowSkills = true) {
    return `<fieldset class="od-bonuses"><legend>Passive bonuses</legend>${rangeEditor("bonusRange:", bonuses.rangeReduction)}${ATTRIBUTES.map((k) => field("attr:" + k, k, bonuses.attributes?.[k] || 0, "number")).join("")}${allowSkills ? skillChoices.map(([k, label]) => field("skill:" + k, `${label} \xB7 skill level`, bonuses.skills?.[k] || 0, "number")).join("") + `<div class="od-new-bonus-skill">${field("newBonusSkillName", "New skill name")}${field("newBonusSkillValue", "Level bonus", 1, "number", 10)}<button type="button" data-add-bonus-skill>Add skill bonus</button></div>` : ""}</fieldset>`;
  }
  function abilityRow(a = {}) {
    return `<div class="od-ability-row">${field("abilityName", "Ability", a.name || "")}${field("abilityLevel", "Level", a.level || 0, "number", 10)}${field("abilityDescription", "Description", a.description || "")}<button type="button" data-remove-ability aria-label="Remove ability">\xD7</button></div>`;
  }
  function collect() {
    const data = Object.fromEntries(new FormData(form));
    Object.assign(values, data);
    if (form.elements.acceptsMagazines) values.acceptsMagazines = form.elements.acceptsMagazines.checked;
    if (form.querySelector('[name="magazineCompatible"]')) values.acceptedMagazineIds = [...form.querySelectorAll('[name="magazineCompatible"]:checked')].map((e) => e.value);
    if (form.querySelector("[data-equipment-slots]")) values.slots = [...form.querySelectorAll('[name="equipmentSlot"]:checked')].map((e) => e.value);
    values.playerRemovable = !!form.elements.playerRemovable?.checked;
    const bonuses = { attributes: {}, skills: {} };
    for (const input of form.querySelectorAll('[name^="attr:"],[name^="skill:"]')) {
      const [kind, ...key] = input.name.split(":");
      const value = Number(input.value) || 0;
      if (value) bonuses[kind === "attr" ? "attributes" : "skills"][key.join(":")] = value;
    }
    if (form.querySelector(".od-bonuses")) {
      bonuses.rangeReduction = normalizeRangeReduction(Object.fromEntries(RANGE_BANDS.map(([key]) => [key, data["bonusRange:" + key]])));
      values.bonuses = bonuses;
    }
    if (form.elements["weaponRange:clinch"]) values.rangeReduction = normalizeRangeReduction(Object.fromEntries(RANGE_BANDS.map(([key]) => [key, data["weaponRange:" + key]])));
    if (form.querySelector("[data-abilities]")) values.abilities = [...form.querySelectorAll(".od-ability-row")].map((row, index) => ({ id: existing.abilities?.[index]?.id || crypto.randomUUID(), name: row.querySelector('[name="abilityName"]').value, level: Number(row.querySelector('[name="abilityLevel"]').value) || 0, description: row.querySelector('[name="abilityDescription"]').value })).filter((a) => a.name.trim());
    values.ammoIds = [...form.querySelectorAll('[name="ammoCompatible"]:checked')].map((e) => e.value);
  }
  function render2() {
    const kind = form.elements.kind.value;
    values.kind = kind;
    let html = "";
    if (kind === "consumable") {
      html += select("subtype", "Use type", [["grenade", "Grenade \xB7 placeholder"], ["medkit", "Medkit"], ["repairkit", "Repair kit"], ["stim", "Stim"]], values.subtype || "grenade");
      if (["medkit", "repairkit"].includes(values.subtype)) html += field("healCritical", "Critical wounds restored", values.healCritical ?? 1, "number");
      if (values.subtype === "stim") html += bonusEditor(values.bonuses, false);
      if (!values.subtype || values.subtype === "grenade") html += '<p class="cp-muted">Grenades have no action yet.</p>';
    }
    if (kind === "resource") html += bonusEditor(values.bonuses);
    if (kind === "ammo") html += field("damage", "Damage / round", values.damage ?? 0, "number") + field("penetration", "Armor penetration", values.penetration ?? 0, "number");
    if (kind === "magazine") html += field("capacity", "Capacity", values.capacity ?? 6, "number") + field("loaded", "Rounds \xB7 GM correction", values.loaded ?? 0, "number") + select("loadedAmmoId", "Ammunition", [["", "Empty"], ...ammunition.map((a) => [a.id, a.name])], values.loadedAmmoId || "");
    if (kind === "weapon") {
      html += select("weaponKind", "Weapon type", [["melee", "Melee"], ["ranged", "Ranged"]], values.weaponKind || "melee") + field("accuracy", "Accuracy", values.accuracy ?? 0, "number");
      if (values.weaponKind === "ranged") {
        html += select("weaponCategory", "Weapon category", WEAPON_CATEGORIES, values.weaponCategory || "") + rangeEditor("weaponRange:", values.rangeReduction);
        html += checkbox("acceptsMagazines", "Accepts detachable magazines", values.acceptsMagazines);
        if (values.acceptsMagazines) html += "<fieldset><legend>Compatible magazines</legend>" + magazineChoices.map((m) => checkbox("magazineCompatible", m.name + " \xB7 " + m.capacity + " rounds", values.acceptedMagazineIds.some((id) => magazines.some((reference) => reference.id === id && reference.name === m.name && reference.capacity === m.capacity))).replace('name="magazineCompatible"', 'name="magazineCompatible" value="' + esc(m.id) + '"')).join("") + (magazines.length ? "" : '<p class="cp-muted">Create a magazine in Items first.</p>') + "</fieldset>";
        else html += field("capacity", "Built-in magazine capacity", values.capacity || 6, "number") + field("loaded", "Loaded rounds \xB7 GM correction", values.loaded || 0, "number") + select("loadedAmmoId", "Loaded ammunition", [["", "Empty"], ...ammunition.map((a) => [a.id, a.name])], values.loadedAmmoId || "") + "<fieldset><legend>Compatible ammunition</legend>" + ammunition.map((a) => checkbox("ammoCompatible", a.name, (values.ammoIds || []).includes(a.id)).replace('name="ammoCompatible"', 'name="ammoCompatible" value="' + esc(a.id) + '"')).join("") + "</fieldset>";
      } else html += field("damage", "Damage", values.damage ?? 0, "number");
    }
    if (["armor", "shield", "implant"].includes(kind)) {
      html += field("armor", "Armor", values.armor ?? 0, "number") + field("criticalMax", kind === "implant" ? "Body part HP bonus" : "Critical durability", values.criticalMax ?? 1, "number") + (kind === "implant" ? "" : field("criticalDamage", "Critical damage taken", values.criticalDamage ?? 0, "number")) + '<fieldset data-equipment-slots class="od-equipment-slots"><legend>Equipment slots \xB7 choose allowed parts</legend>' + [["any", "Any part"], ...parts, ...values.slots.filter((key) => key !== "any" && !parts.some(([k]) => k === key)).map((key) => [key, key])].map(([key, label]) => checkbox("equipmentSlot", label, values.slots.includes(key)).replace('name="equipmentSlot"', 'name="equipmentSlot" value="' + esc(key) + '"')).join("") + "</fieldset>";
      if (kind === "implant") html += '<p class="cp-muted">Adds maximum body-part HP. Equipping does not heal wounds.</p>';
      if (kind === "implant") html += checkbox("playerRemovable", "Player may equip / remove this implant", values.playerRemovable) + bonusEditor(values.bonuses) + `<fieldset><legend>Granted abilities \xB7 available only while equipped</legend><div data-abilities>${(values.abilities || []).map(abilityRow).join("")}</div><button type="button" data-add-ability>+ Ability</button></fieldset>`;
    }
    host.innerHTML = html;
    if (form.elements.weaponCategory && !existing.id) form.elements.weaponCategory.required = true;
    for (const input of host.querySelectorAll('[name^="attr:"],[name^="skill:"]')) input.min = "-999";
    form.elements.quantity.closest("label").hidden = kind === "weapon" || kind === "magazine" && !!existing.id;
  }
  form.addEventListener("change", (e) => {
    if (["kind", "subtype", "weaponKind", "acceptsMagazines"].includes(e.target.name)) {
      collect();
      render2();
    }
  });
  form.addEventListener("click", (e) => {
    if (e.target.closest("[data-add-bonus-skill]")) {
      const input = form.elements.newBonusSkillName, name = input.value.trim(), bonus = Number(form.elements.newBonusSkillValue.value) || 0;
      if (!name || name.length > 40 || ["__proto__", "constructor", "prototype"].includes(name)) {
        input.setCustomValidity("Enter a skill name (1\u201340 characters).");
        input.reportValidity();
        input.oninput = () => input.setCustomValidity("");
        return;
      }
      collect();
      let choice = skillChoices.find(([key, label]) => key.toLowerCase() === name.toLowerCase() || label.toLowerCase() === name.toLowerCase());
      if (!choice) {
        choice = [name, name];
        skillChoices.push(choice);
      }
      values.bonuses.skills[choice[0]] = bonus;
      render2();
      return;
    }
    if (e.target.closest("[data-add-ability]")) {
      host.querySelector("[data-abilities]").insertAdjacentHTML("beforeend", abilityRow());
    }
    e.target.closest("[data-remove-ability]")?.closest(".od-ability-row").remove();
  });
  document.body.append(dialog);
  render2();
  return new Promise((resolve) => {
    const finish = (v) => {
      dialog.close();
      dialog.remove();
      resolve(v);
    };
    form.querySelector("[data-cancel]").onclick = () => finish(null);
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      finish(null);
    });
    form.addEventListener("submit", (e) => {
      e.preventDefault();
      collect();
      if (["armor", "shield", "implant"].includes(values.kind) && !values.slots.length) {
        host.querySelector("[data-equipment-slots]").insertAdjacentHTML("beforeend", '<p class="cp-muted">Choose at least one allowed slot.</p>');
        return;
      }
      if (!values.name.trim()) {
        form.elements.name.setCustomValidity("Enter a name.");
        form.elements.name.reportValidity();
        form.elements.name.oninput = () => form.elements.name.setCustomValidity("");
        return;
      }
      finish(values);
    });
    dialog.showModal();
  });
}
function skillEditor(o, key) {
  const existing = key && Object.hasOwn(o.skills, key), core = ["Melee", "Parry"].includes(key);
  const dialog = document.createElement("dialog");
  dialog.className = "od-dialog";
  dialog.innerHTML = `<form><h2>${existing ? "Edit skill" : "New skill"}</h2><label>Name<input name="name" required maxlength="40" value="${esc(existing ? o.skillLabels[key] || key : "")}"></label><label>Base level<input name="level" type="number" min="0" max="10" value="${existing ? o.skills[key] : 0}"></label><label>Category<select name="category">${["combat", "abilities", "applied"].map((c) => `<option ${o.skillCategories[key] === c ? "selected" : ""}>${c}</option>`).join("")}</select></label><label class="od-check"><input name="strength" type="checkbox" ${o.skillStrengthBonuses[key] ? "checked" : ""}>Strength bonus for attack damage</label><fieldset><legend>Perks</legend><p class="cp-muted">Combined perk relief for attacks using this skill. Enter each distance band separately.</p>${RANGE_BANDS.map(([band, label]) => `<label>${esc(label)} \xB7 penalty reduction<input name="perkRange:${band}" type="number" min="0" max="999" value="${o.skillRangeReduction?.[key]?.[band] || 0}"></label>`).join("")}<div data-perks></div><button type="button" data-add-perk>+ Perk</button></fieldset><div class="cp-row"><button type="button" data-cancel>Cancel</button>${existing && !core ? '<button type="button" class="danger" data-delete>Delete skill</button>' : ""}<button type="submit">Save</button></div></form>`;
  const list = dialog.querySelector("[data-perks]");
  const add = (value) => list.insertAdjacentHTML("beforeend", `<div class="cp-row"><input name="perk" maxlength="40" value="${esc(value)}" aria-label="Perk"><button type="button" data-remove-perk aria-label="Remove perk">\xD7</button></div>`);
  for (const p of o.skillPerks[key] || []) add(p);
  document.body.append(dialog);
  return new Promise((resolve) => {
    const finish = (v) => {
      dialog.close();
      dialog.remove();
      resolve(v);
    };
    dialog.querySelector("[data-cancel]").onclick = () => finish(null);
    dialog.querySelector("[data-delete]")?.addEventListener("click", () => finish({ delete: true }));
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      finish(null);
    });
    dialog.addEventListener("click", (e) => {
      if (e.target.closest("[data-add-perk]") && list.children.length < 12) add("");
      e.target.closest("[data-remove-perk]")?.parentElement.remove();
    });
    dialog.querySelector("form").onsubmit = (e) => {
      e.preventDefault();
      const f = e.target;
      finish({ name: f.elements.name.value.trim(), level: Number(f.elements.level.value), category: f.elements.category.value, strength: f.elements.strength.checked, rangeReduction: normalizeRangeReduction(Object.fromEntries(RANGE_BANDS.map(([band]) => [band, f.elements["perkRange:" + band].value]))), perks: [...list.querySelectorAll("input")].map((e2) => e2.value.trim()).filter(Boolean) });
    };
    dialog.showModal();
  });
}

// ammunition.js
var integer = (value, max = 999999) => Math.min(max, Math.max(0, Math.floor(Number(value) || 0)));
var modifier = (value) => Math.min(999, Math.max(-999, Math.trunc(Number(value) || 0)));
function sanitizeAmmunition(raw) {
  const seen = /* @__PURE__ */ new Set();
  return (Array.isArray(raw) ? raw : []).filter((item) => {
    if (!item || typeof item.id !== "string" || !item.id || seen.has(item.id)) return false;
    seen.add(item.id);
    return true;
  }).map((item) => ({
    id: item.id,
    name: String(item.name || "Ammo").trim() || "Ammo",
    damage: integer(item.damage, 999),
    penetration: modifier(item.penetration),
    quantity: integer(item.quantity)
  }));
}
function sanitizeMagazine(weapon) {
  const capacity = integer(weapon.capacity, 999);
  return {
    weaponCategory: normalizeWeaponCategory(weapon.weaponCategory),
    rangeReduction: normalizeRangeReduction(weapon.rangeReduction),
    acceptsMagazines: weapon.acceptsMagazines !== false,
    ...Array.isArray(weapon.acceptedMagazineIds) ? { acceptedMagazineIds: [...new Set(weapon.acceptedMagazineIds.filter((id) => typeof id === "string"))].slice(0, 200) } : {},
    ammoIds: [...new Set(Array.isArray(weapon.ammoIds) ? weapon.ammoIds.filter((id) => typeof id === "string") : [])],
    capacity,
    loadedAmmoId: String(weapon.loadedAmmoId || ""),
    loaded: integer(weapon.loaded, capacity)
  };
}
function reloadMagazine(odyssey, weaponIndex, ammoId) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  const ammo = odyssey.ammunition.find((item) => item.id === ammoId);
  if (!weapon || !ammo || !weapon.ammoIds.includes(ammoId)) throw new Error("Choose compatible ammunition.");
  if (weapon.capacity < 1) throw new Error("Set magazine capacity first.");
  if (ammo.quantity < 1) throw new Error("Ammo reserve is empty.");
  if (weapon.loadedAmmoId !== ammoId && weapon.loaded > 0) {
    const previous = odyssey.ammunition.find((item) => item.id === weapon.loadedAmmoId);
    if (!previous) throw new Error("Loaded ammunition is missing from the inventory.");
    if (previous.quantity + weapon.loaded > 999999) throw new Error("Ammo reserve is full.");
    previous.quantity += weapon.loaded;
    weapon.loaded = 0;
  }
  const count = Math.min(weapon.capacity - weapon.loaded, ammo.quantity);
  if (count <= 0) throw new Error("Magazine is full or ammo reserve is empty.");
  weapon.loadedAmmoId = ammoId;
  weapon.loaded += count;
  ammo.quantity -= count;
}
function getAmmoShot(odyssey, weaponIndex, ammoId, count) {
  const weapon = odyssey.weapons.ranged[weaponIndex];
  const ammo = odyssey.ammunition.find((item) => item.id === ammoId);
  if (!weapon || !ammo || !weapon.ammoIds.includes(ammoId) || weapon.loadedAmmoId !== ammoId) {
    throw new Error("The selected ammunition must be loaded by the GM first.");
  }
  if (!Number.isInteger(count) || count < 1 || count > weapon.loaded || count > weapon.capacity) {
    throw new Error("Not enough rounds in the magazine.");
  }
  return { ammoId, count, damage: ammo.damage * count, penetration: ammo.penetration, name: ammo.name };
}
function consumeAmmo(odyssey, weaponIndex, shot) {
  getAmmoShot(odyssey, weaponIndex, shot.ammoId, shot.count);
  odyssey.weapons.ranged[weaponIndex].loaded -= shot.count;
}
function effectiveArmor(armor, penetration) {
  if (penetration < 0 && armor <= 0) return armor;
  return Math.max(0, armor - penetration);
}

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

// node_modules/@owlbear-rodeo/sdk/lib/builders/ImageBuilder.js
var ImageBuilder = class extends GenericItemBuilder {
  constructor(player, image, grid) {
    super(player);
    this._image = image;
    this._grid = grid;
    this._item.name = "Image";
    this._text = {
      richText: [
        {
          type: "paragraph",
          children: [{ text: "" }]
        }
      ],
      plainText: "",
      style: {
        padding: 8,
        fontFamily: "Roboto",
        fontSize: 24,
        fontWeight: 400,
        textAlign: "CENTER",
        textAlignVertical: "BOTTOM",
        fillColor: "white",
        fillOpacity: 1,
        strokeColor: "white",
        strokeOpacity: 1,
        strokeWidth: 0,
        lineHeight: 1.5
      },
      type: "PLAIN",
      width: "AUTO",
      height: "AUTO"
    };
    this._textItemType = "LABEL";
  }
  text(text2) {
    this._text = text2;
    return this.self();
  }
  textItemType(textItemType) {
    this._textItemType = textItemType;
    return this.self();
  }
  textWidth(width) {
    this._text.width = width;
    return this.self();
  }
  textHeight(height) {
    this._text.height = height;
    return this.self();
  }
  richText(richText) {
    this._text.richText = richText;
    return this.self();
  }
  plainText(plainText) {
    this._text.plainText = plainText;
    return this.self();
  }
  textType(textType) {
    this._text.type = textType;
    return this.self();
  }
  textPadding(padding) {
    this._text.style.padding = padding;
    return this.self();
  }
  fontFamily(fontFamily) {
    this._text.style.fontFamily = fontFamily;
    return this.self();
  }
  fontSize(fontSize) {
    this._text.style.fontSize = fontSize;
    return this.self();
  }
  fontWeight(fontWeight) {
    this._text.style.fontWeight = fontWeight;
    return this.self();
  }
  textAlign(textAlign) {
    this._text.style.textAlign = textAlign;
    return this.self();
  }
  textAlignVertical(textAlignVertical) {
    this._text.style.textAlignVertical = textAlignVertical;
    return this.self();
  }
  textFillColor(fillColor) {
    this._text.style.fillColor = fillColor;
    return this.self();
  }
  textFillOpacity(fillOpacity) {
    this._text.style.fillOpacity = fillOpacity;
    return this.self();
  }
  textStrokeColor(strokeColor) {
    this._text.style.strokeColor = strokeColor;
    return this.self();
  }
  textStrokeOpacity(strokeOpacity) {
    this._text.style.strokeOpacity = strokeOpacity;
    return this.self();
  }
  textStrokeWidth(strokeWidth) {
    this._text.style.strokeWidth = strokeWidth;
    return this.self();
  }
  textLineHeight(lineHeight) {
    this._text.style.lineHeight = lineHeight;
    return this.self();
  }
  build() {
    return Object.assign(Object.assign({}, this._item), { type: "IMAGE", image: this._image, grid: this._grid, text: this._text, textItemType: this._textItemType });
  }
};

// node_modules/@owlbear-rodeo/sdk/lib/builders/RulerBuilder.js
var RulerBuilder = class extends GenericItemBuilder {
  constructor(player) {
    super(player);
    this._startPosition = { x: 0, y: 0 };
    this._endPosition = { x: 0, y: 0 };
    this._measurement = "";
    this._item.layer = "RULER";
    this._item.name = "Ruler";
    this._style = {
      variant: "FILLED"
    };
  }
  startPosition(startPosition) {
    this._startPosition = startPosition;
    return this.self();
  }
  endPosition(endPosition) {
    this._endPosition = endPosition;
    return this.self();
  }
  measurement(measurement) {
    this._measurement = measurement;
    return this.self();
  }
  variant(variant) {
    this._style.variant = variant;
    return this.self();
  }
  build() {
    return Object.assign(Object.assign({}, this._item), { type: "RULER", startPosition: this._startPosition, endPosition: this._endPosition, measurement: this._measurement, style: this._style });
  }
};

// node_modules/@owlbear-rodeo/sdk/lib/builders/TextBuilder.js
var TextBuilder = class extends GenericItemBuilder {
  constructor(player) {
    super(player);
    this._text = {
      richText: [
        {
          type: "paragraph",
          children: [{ text: "" }]
        }
      ],
      plainText: "",
      style: {
        padding: 0,
        fontFamily: "Roboto",
        fontSize: 16,
        fontWeight: 400,
        textAlign: "LEFT",
        textAlignVertical: "TOP",
        fillColor: "white",
        fillOpacity: 1,
        strokeColor: "white",
        strokeOpacity: 1,
        strokeWidth: 0,
        lineHeight: 1.5
      },
      type: "RICH",
      width: "AUTO",
      height: "AUTO"
    };
    this._item.layer = "TEXT";
    this._item.name = "Text";
  }
  text(text2) {
    this._text = text2;
    return this.self();
  }
  width(width) {
    this._text.width = width;
    return this.self();
  }
  height(height) {
    this._text.height = height;
    return this.self();
  }
  richText(richText) {
    this._text.richText = richText;
    return this.self();
  }
  plainText(plainText) {
    this._text.plainText = plainText;
    return this.self();
  }
  textType(textType) {
    this._text.type = textType;
    return this.self();
  }
  padding(padding) {
    this._text.style.padding = padding;
    return this.self();
  }
  fontFamily(fontFamily) {
    this._text.style.fontFamily = fontFamily;
    return this.self();
  }
  fontSize(fontSize) {
    this._text.style.fontSize = fontSize;
    return this.self();
  }
  fontWeight(fontWeight) {
    this._text.style.fontWeight = fontWeight;
    return this.self();
  }
  textAlign(textAlign) {
    this._text.style.textAlign = textAlign;
    return this.self();
  }
  textAlignVertical(textAlignVertical) {
    this._text.style.textAlignVertical = textAlignVertical;
    return this.self();
  }
  fillColor(fillColor) {
    this._text.style.fillColor = fillColor;
    return this.self();
  }
  fillOpacity(fillOpacity) {
    this._text.style.fillOpacity = fillOpacity;
    return this.self();
  }
  strokeColor(strokeColor) {
    this._text.style.strokeColor = strokeColor;
    return this.self();
  }
  strokeOpacity(strokeOpacity) {
    this._text.style.strokeOpacity = strokeOpacity;
    return this.self();
  }
  strokeWidth(strokeWidth) {
    this._text.style.strokeWidth = strokeWidth;
    return this.self();
  }
  lineHeight(lineHeight) {
    this._text.style.lineHeight = lineHeight;
    return this.self();
  }
  build() {
    return Object.assign(Object.assign({}, this._item), { type: "TEXT", text: this._text });
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

// node_modules/@owlbear-rodeo/sdk/lib/builders/SceneUploadBuilder.js
var SceneUploadBuilder = class {
  constructor() {
    this._upload = {
      name: "New Scene",
      fog: { filled: false, style: { color: "#222222", strokeWidth: 5 } },
      grid: {
        dpi: 150,
        scale: "5ft",
        style: {
          lineColor: "LIGHT",
          lineOpacity: 0.4,
          lineType: "DASHED",
          lineWidth: 2
        },
        measurement: "CHEBYSHEV",
        type: "SQUARE"
      },
      items: []
    };
  }
  name(name) {
    this._upload.name = name;
    return this;
  }
  fogFilled(filled) {
    this._upload.fog.filled = filled;
    return this;
  }
  fogColor(color) {
    this._upload.fog.style.color = color;
    return this;
  }
  fogStrokeWidth(width) {
    this._upload.fog.style.strokeWidth = width;
    return this;
  }
  gridScale(scale) {
    this._upload.grid.scale = scale;
    return this;
  }
  gridColor(color) {
    this._upload.grid.style.lineColor = color;
    return this;
  }
  gridOpacity(opacity) {
    this._upload.grid.style.lineOpacity = opacity;
    return this;
  }
  gridLineType(lineType) {
    this._upload.grid.style.lineType = lineType;
    return this;
  }
  gridMeasurement(measurement) {
    this._upload.grid.measurement = measurement;
    return this;
  }
  gridType(type) {
    this._upload.grid.type = type;
    return this;
  }
  items(items) {
    this._upload.items = items;
    return this;
  }
  baseMap(baseMap) {
    this._upload.baseMap = baseMap;
    return this;
  }
  thumbnail(thumbnail) {
    this._upload.thumbnail = thumbnail;
    return this;
  }
  build() {
    return this._upload;
  }
};

// node_modules/js-base64/base64.mjs
var _TD = typeof TextDecoder === "function" ? new TextDecoder("utf-8", { ignoreBOM: true }) : void 0;
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
var btoaPolyfill = (bin) => {
  let u32, c0, c1, c2, asc = "";
  const pad = bin.length % 3;
  for (let i = 0; i < bin.length; ) {
    if ((c0 = bin.charCodeAt(i++)) > 255 || (c1 = bin.charCodeAt(i++)) > 255 || (c2 = bin.charCodeAt(i++)) > 255)
      throw new TypeError("invalid character found");
    u32 = c0 << 16 | c1 << 8 | c2;
    asc += b64chs[u32 >> 18 & 63] + b64chs[u32 >> 12 & 63] + b64chs[u32 >> 6 & 63] + b64chs[u32 & 63];
  }
  return pad ? asc.slice(0, pad - 3) + "===".substring(pad) : asc;
};
var _btoa = typeof btoa === "function" ? (bin) => btoa(bin) : btoaPolyfill;
var _fromUint8Array = typeof Uint8Array.prototype.toBase64 === "function" ? Function.prototype.call.bind(Uint8Array.prototype.toBase64) : (u8a) => {
  const maxargs = 4096;
  let strs = [];
  for (let i = 0, l = u8a.length; i < l; i += maxargs) {
    strs.push(_fromCC.apply(null, u8a.subarray(i, i + maxargs)));
  }
  return _btoa(strs.join(""));
};
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
var _atob = typeof atob === "function" ? (asc) => atob(_tidyB64(asc)) : atobPolyfill;
var _toUint8Array = typeof Uint8Array.fromBase64 === "function" ? (a) => Uint8Array.fromBase64(a) : (a) => _U8Afrom(_atob(a).split("").map((c) => c.charCodeAt(0)));
var _decode = _TD ? (a) => _TD.decode(_toUint8Array(a)) : (a) => btou(_atob(a));
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
function buildImage(image, grid) {
  return new ImageBuilder(playerApi, image, grid);
}
function buildRuler() {
  return new RulerBuilder(playerApi);
}
function buildText() {
  return new TextBuilder(playerApi);
}
function buildPath() {
  return new PathBuilder(playerApi);
}
function buildSceneUpload() {
  return new SceneUploadBuilder();
}
var lib_default = OBR;

// profiles.js
var LIBRARY_FORMAT = "odyssey-template-library";
var TEMPLATE_FORMAT = "odyssey-character-template";
var LIBRARY_ITEM_KEY = "com.codex.body-hp/template-library";
var LIBRARY_STORAGE_KEY = "odyssey.templates.v1";
var profileFields = ["skills", "skillCategories", "skillStrengthBonuses", "skillLabels", "skillPerks", "skillRangeReduction", "attributes", "weapons", "ammunition", "inventory"];
function profileData(tracker) {
  return {
    enabled: true,
    minor: tracker.minor || 0,
    serious: tracker.serious || 0,
    body: structuredClone(tracker.body),
    odyssey: Object.fromEntries(profileFields.map((k) => [k, structuredClone(tracker.odyssey[k])]))
  };
}
function makeTemplate(name, tracker) {
  const data = profileData(tracker);
  data.minor = 0;
  data.serious = 0;
  for (const item of data.odyssey.inventory.items) item.criticalDamage = 0;
  reconcileEquipment(data);
  for (const part of Object.values(data.body)) {
    part.current = part.max;
    part.minor = 0;
    part.serious = 0;
  }
  return { format: TEMPLATE_FORMAT, version: 1, id: newId(), name: String(name || "Character").trim().slice(0, 120), updatedAt: (/* @__PURE__ */ new Date()).toISOString(), data };
}
function normalizeLibrary(raw, sanitize) {
  if (raw?.format !== LIBRARY_FORMAT || raw.version !== 1 || !Array.isArray(raw.templates)) throw new Error("Choose an Odyssey template library JSON (version 1).");
  if (raw.templates.length > 100) throw new Error("A library can contain at most 100 templates.");
  const seen = /* @__PURE__ */ new Set();
  const templates = raw.templates.map((t) => {
    if (t?.format !== TEMPLATE_FORMAT || t.version !== 1 || !t.data?.body || !t.data?.odyssey) throw new Error("Invalid character template.");
    const id = String(t.id || newId()).slice(0, 120);
    if (seen.has(id)) throw new Error("Duplicate template IDs in library.");
    seen.add(id);
    return { format: TEMPLATE_FORMAT, version: 1, id, name: String(t.name || "Character").trim().slice(0, 120), updatedAt: String(t.updatedAt || ""), data: profileData(sanitize(t.data)) };
  });
  return { format: LIBRARY_FORMAT, version: 1, templates };
}
function emptyLibrary() {
  return { format: LIBRARY_FORMAT, version: 1, templates: [] };
}
function mergeLibraries(local, imported) {
  const entries = new Map(local.templates.map((t) => [t.id, t]));
  for (const template of imported.templates) entries.set(template.id, template);
  if (entries.size > 100) throw new Error("Library limit is 100 templates. Export or remove templates first.");
  return { ...emptyLibrary(), templates: [...entries.values()] };
}
function instantiateTemplate(current2, template, sanitize) {
  const validated = normalizeLibrary({ ...emptyLibrary(), templates: [template] }, sanitize).templates[0];
  const data = structuredClone(validated.data);
  const ammoIds = new Map(data.odyssey.ammunition.map((a) => [a.id, newId()]));
  const magIds = new Map(data.odyssey.inventory.magazines.map((m) => [m.id, newId()]));
  for (const a of data.odyssey.ammunition) a.id = ammoIds.get(a.id);
  for (const mag of data.odyssey.inventory.magazines) {
    mag.id = magIds.get(mag.id);
    mag.loadedAmmoId = ammoIds.get(mag.loadedAmmoId) || "";
  }
  for (const w of data.odyssey.weapons.ranged) {
    if (w.acceptedMagazineIds) w.acceptedMagazineIds = w.acceptedMagazineIds.map((id) => magIds.get(id)).filter(Boolean);
    w.id = newId();
    w.ammoIds = w.ammoIds.map((id) => ammoIds.get(id)).filter(Boolean);
    w.loadedAmmoId = ammoIds.get(w.loadedAmmoId) || "";
    w.magazineId = magIds.get(w.magazineId) || "";
  }
  for (const w of data.odyssey.weapons.melee) w.id = newId();
  for (const item of data.odyssey.inventory.items) item.id = newId();
  const next = sanitize({ ...current2, ...data, odyssey: { ...current2.odyssey, ...data.odyssey, owner: current2.odyssey.owner, attackDraft: { targetTokenId: "", targetTokenName: "" } } });
  return next;
}
function readLibrary(storage, sanitize) {
  const raw = storage.getItem(LIBRARY_STORAGE_KEY);
  return raw ? normalizeLibrary(JSON.parse(raw), sanitize) : emptyLibrary();
}
function writeLibrary(storage, library) {
  const json = JSON.stringify(library);
  if (new TextEncoder().encode(json).length > 4 * 1024 * 1024) throw new Error("Library exceeds 4 MB. Export a smaller library.");
  storage.setItem(LIBRARY_STORAGE_KEY, json);
}
function libraryFromScenes(scenes, sanitize) {
  let result = emptyLibrary();
  let found = false;
  for (const scene of scenes || []) for (const item of scene.items || []) {
    const payload = item.metadata?.[LIBRARY_ITEM_KEY];
    if (payload) {
      result = mergeLibraries(result, normalizeLibrary(payload, sanitize));
      found = true;
    }
  }
  if (!found) throw new Error("Selected scenes contain no Odyssey template library.");
  return result;
}

// ui/characterTools.js
function formDialog(title, fields, { danger = false, message = "", onChange } = {}) {
  const dialog = document.createElement("dialog");
  dialog.className = "od-dialog";
  dialog.innerHTML = `<form><h2>${esc(title)}</h2>${message ? `<p>${esc(message)}</p>` : ""}${fields.map((f) => `<label>${esc(f.label)}${f.options ? `<select name="${esc(f.name)}">${f.options.map(([v, l]) => `<option value="${esc(v)}" ${v === f.value ? "selected" : ""}>${esc(l)}</option>`).join("")}</select>` : f.type === "textarea" ? `<textarea name="${esc(f.name)}" maxlength="2000">${esc(f.value)}</textarea>` : `<input name="${esc(f.name)}" type="${esc(f.type || "text")}" value="${esc(f.value ?? "")}" ${f.required ? "required" : ""} ${f.type === "number" ? `min="${f.min ?? 0}" max="${f.max ?? 999999}" step="1"` : 'maxlength="120"'}>`}</label>`).join("")}<div class="cp-row"><button type="button" data-cancel>Cancel</button><button type="submit" class="${danger ? "danger" : ""}">${danger ? "Apply" : "Save"}</button></div></form>`;
  document.body.append(dialog);
  if (onChange) dialog.querySelector("form").addEventListener("change", (event) => onChange(event.target.form, event));
  return new Promise((resolve) => {
    let done = false;
    const finish = (value) => {
      if (done) return;
      done = true;
      dialog.close();
      dialog.remove();
      resolve(value);
    };
    dialog.querySelector("[data-cancel]").onclick = () => finish(null);
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      finish(null);
    });
    dialog.querySelector("form").addEventListener("submit", (e) => {
      e.preventDefault();
      finish(Object.fromEntries(new FormData(e.target)));
    });
    dialog.showModal();
  });
}
function installCharacterTools(host) {
  const { OBR: OBR2, sanitize, getToken, getRole, canControl, update, refreshOverlay, render: render2, status, download } = host;
  let library, storageError = "";
  try {
    library = readLibrary(window.localStorage, sanitize);
  } catch (e) {
    library = emptyLibrary();
    storageError = e.message;
  }
  let poolMetadata = {}, catalog = { entries: [], errors: 0 }, poolDialog = null, scanTimer = null, scanRunning = false, scanAgain = false, scanFingerprint = "";
  const knownBlueprintIds = /* @__PURE__ */ new Map();
  const refreshPool = (metadata) => {
    poolMetadata = metadata || {};
    catalog = poolEntries(poolMetadata);
    poolDialog?.refresh();
    render2();
  };
  const requirePoolGM = async () => {
    if (getRole() !== "GM" || await OBR2.player.getRole() !== "GM") throw new Error("Only the GM can edit the campaign item pool.");
  };
  const putBlueprints = async (blueprints, automatic = false) => {
    await requirePoolGM();
    const metadata = await OBR2.room.getMetadata(), patch = {};
    let count = Object.entries(metadata).filter(([key, value]) => key.startsWith(ITEM_POOL_PREFIX) && value && !value.deleted).length;
    for (const blueprint of blueprints) {
      validateBlueprint(blueprint);
      const text2 = blueprintText(blueprint);
      let id = knownBlueprintIds.get(text2);
      if (!id) {
        id = await blueprintId(blueprint);
        knownBlueprintIds.set(text2, id);
      }
      const key = ITEM_POOL_PREFIX + id;
      if (metadata[key] && (automatic || !metadata[key].deleted)) continue;
      if (Object.hasOwn(patch, key)) continue;
      if ((!metadata[key] || metadata[key].deleted) && ++count > 2e3) throw new Error("Campaign item pool is full. Export a backup and remove unused definitions.");
      patch[key] = { blueprint };
    }
    await requirePoolGM();
    if (Object.keys(patch).length) await OBR2.room.setMetadata(patch);
    refreshPool({ ...metadata, ...patch });
  };
  const scanPool = async () => {
    if (getRole() !== "GM") return;
    if (scanRunning) {
      scanAgain = true;
      return;
    }
    scanRunning = true;
    try {
      const candidates = /* @__PURE__ */ new Map();
      for (const token of host.getTokens()) {
        const t = host.tracker(token), o = t.odyssey;
        const entries = [...o.inventory.items.map((i) => ["item", i.id]), ...o.inventory.magazines.map((i) => ["magazine", i.id]), ...o.ammunition.map((i) => ["ammo", i.id]), ...["melee", "ranged"].flatMap((kind) => o.weapons[kind].map((i) => ["weapon-" + kind, i.id]))];
        for (const [type, id] of entries) {
          const blueprint = captureBlueprint(t, type, id);
          candidates.set(blueprintText(blueprint), blueprint);
        }
      }
      const fingerprint = [...candidates.keys()].sort().join("\n");
      if (fingerprint === scanFingerprint) return;
      scanFingerprint = fingerprint;
      await putBlueprints([...candidates.values()], true);
    } catch (e) {
      scanFingerprint = "";
      status("Item pool: " + e.message, "error");
    } finally {
      scanRunning = false;
      if (scanAgain) {
        scanAgain = false;
        schedulePoolScan();
      }
    }
  };
  const schedulePoolScan = () => {
    if (getRole() !== "GM" || scanTimer) return;
    scanTimer = setTimeout(() => {
      scanTimer = null;
      void scanPool();
    }, 400);
  };
  void OBR2.room.getMetadata().then(refreshPool).catch((e) => status("Item pool: " + e.message, "error"));
  OBR2.room.onMetadataChange(refreshPool);
  let busy = false;
  const saveLibrary = (next) => {
    if (getRole() !== "GM") throw new Error("Only the GM can edit templates.");
    writeLibrary(window.localStorage, next);
    library = next;
    storageError = "";
    render2();
  };
  const check = (id, gm = false) => {
    const token = getToken(id);
    if (!token || !canControl(token) || gm && getRole() !== "GM") throw new Error(gm ? "Only the GM can perform this action." : "You cannot control this character.");
    return token;
  };
  const mutate = async (id, fn, gm = false) => {
    check(id, gm);
    await update(id, (current2) => {
      check(id, gm);
      const next = structuredClone(current2);
      fn(next);
      return next;
    });
    await refreshOverlay(id);
    render2();
  };
  const readSelect = (root, attribute, value) => Array.from(root.querySelectorAll(`[${attribute}]`)).find((e) => e.getAttribute(attribute) === String(value))?.value || "";
  async function handle(action, node) {
    const token = check();
    const tokenId = token.id, panel = node.closest("#selectedTokenPanel"), id = node.dataset.id;
    const gmActions = ["add-item", "edit-item", "delete-item", "add-mag", "fill-mag", "set-mag-rounds", "empty-mag", "delete-mag", "mount-mag", "apply-template"];
    if (action === "open-item-pool") check(tokenId, true);
    if (gmActions.includes(action) || action.includes("template") || action.includes("library") || action.startsWith("account-")) check(tokenId, true);
    if (action === "add-item" || action === "edit-entry") {
      check(tokenId, true);
      const t = host.tracker(token), o = t.odyssey, type = node.dataset.entryType || "item";
      const existing = action === "edit-entry" ? findInventoryEntry(t, type, id) : null;
      if (action === "edit-entry" && !existing) throw new Error("Item no longer exists.");
      const kind = type.startsWith("weapon-") ? "weapon" : type === "item" ? ["item", "tool"].includes(existing?.kind) ? "resource" : existing?.kind === "medical" ? "consumable" : existing?.kind || "resource" : type;
      const values = await itemEditor({ existing: { ...existing, ...type.startsWith("weapon-") ? { weaponKind: type.slice(7) } : {} }, type: kind, skills: Object.entries(o.skills).map(([k]) => [k, o.skillLabels[k] || k]), parts: Object.entries(t.body).map(([k, p]) => [k, p.label || k]), ammunition: o.ammunition, magazines: o.inventory.magazines });
      if (!values) return;
      const entryId = existing?.id || newId(), entryType = values.kind === "weapon" ? "weapon-" + values.weaponKind : ["ammo", "magazine"].includes(values.kind) ? values.kind : "item";
      let blueprint;
      await mutate(tokenId, (t2) => {
        saveInventoryEntry(t2, existing ? { type, id } : null, values, entryId);
        blueprint = captureBlueprint(sanitize(t2), entryType, entryId);
      }, true);
      try {
        await putBlueprints([blueprint]);
      } catch (e) {
        throw new Error("Item saved on character; campaign pool was not updated: " + e.message);
      }
    } else if (action === "open-item-pool") {
      await requirePoolGM();
      refreshPool(await OBR2.room.getMetadata());
      await scanPool();
      poolDialog = showItemPool({
        getCatalog: () => catalog,
        character: token.name || "Character",
        onGive: async (entryId) => {
          const entry = catalog.entries.find((e) => e.id === entryId);
          if (!entry) throw new Error("Item definition no longer exists.");
          const value = await formDialog("Give " + entry.blueprint.values.name, [{ name: "quantity", label: "Quantity", type: "number", min: 1, max: entry.blueprint.type === "ammo" ? 999999 : entry.blueprint.type.startsWith("weapon-") ? 20 : 1e3, value: 1 }]);
          if (!value) return false;
          const copyId = newId();
          await mutate(tokenId, (t) => copyBlueprint(t, entry.blueprint, Number(value.quantity), copyId), true);
          return true;
        },
        onDelete: async (entryId) => {
          if (!await formDialog("Remove from campaign pool?", [], { danger: true, message: "Existing items on characters are kept." })) return false;
          await requirePoolGM();
          await OBR2.room.setMetadata({ [ITEM_POOL_PREFIX + entryId]: { deleted: true } });
          refreshPool(await OBR2.room.getMetadata());
          return true;
        },
        onExport: async () => {
          await requirePoolGM();
          download("odyssey-item-pool.json", JSON.stringify(exportPool(catalog.entries), null, 2));
        },
        onImport: () => new Promise((resolve, reject) => {
          const input = document.createElement("input");
          input.type = "file";
          input.accept = ".json,application/json";
          input.addEventListener("cancel", () => resolve(false));
          input.onchange = async () => {
            try {
              const file = input.files?.[0];
              if (!file) {
                resolve(false);
                return;
              }
              if (file.size > 4 * 1024 * 1024) throw new Error("Item pool JSON must be under 4 MB.");
              await requirePoolGM();
              const blueprints = parsePoolJSON(JSON.parse(await file.text())).map((raw) => {
                const t = sanitize({}), id2 = newId();
                copyBlueprint(t, raw, 1, id2);
                return captureBlueprint(sanitize(t), raw.type, id2);
              });
              await putBlueprints(blueprints);
              resolve(true);
            } catch (e) {
              reject(e);
            }
          };
          input.click();
        })
      });
      await poolDialog.closed;
      poolDialog = null;
    } else if (action === "delete-entry") {
      check(tokenId, true);
      if (!await formDialog("Remove item?", [], { danger: true, message: "Remove this item from this token?" })) return;
      await mutate(tokenId, (t) => deleteInventoryEntry(t, node.dataset.entryType, id), true);
    } else if (action === "add-skill" || action === "edit-skill") {
      check(tokenId, true);
      const key = node.dataset.skill, o = host.tracker(token).odyssey;
      const value = await skillEditor(o, key);
      if (!value) return;
      const skillKey = key || newId();
      await mutate(tokenId, (t) => {
        const d = t.odyssey;
        if (key && !Object.hasOwn(d.skills, key)) throw new Error("Skill no longer exists.");
        if (value.delete) {
          if (["Melee", "Parry"].includes(key)) throw new Error("Core combat skill cannot be removed.");
          for (const field of ["skills", "skillLabels", "skillCategories", "skillStrengthBonuses", "skillPerks", "skillRangeReduction"]) delete d[field][key];
          return;
        }
        d.skills[skillKey] = value.level;
        d.skillLabels[skillKey] = value.name;
        d.skillCategories[skillKey] = value.category;
        d.skillStrengthBonuses[skillKey] = value.strength;
        d.skillPerks[skillKey] = value.perks;
        d.skillRangeReduction ?? (d.skillRangeReduction = {});
        d.skillRangeReduction[skillKey] = value.rangeReduction;
      }, true);
    } else if (action === "edit-training") {
      check(tokenId, true);
      const o = host.tracker(token).odyssey;
      const value = await formDialog("Armor training", Object.entries(ARMOR_TRAINING).map(([key, label]) => ({ name: key, label, type: "number", min: 0, max: 1, value: o.skills[key] || 0 })));
      if (!value) return;
      await mutate(tokenId, (t) => {
        for (const key of Object.keys(ARMOR_TRAINING)) t.odyssey.skills[key] = Number(value[key]) >= 1 ? 1 : 0;
      }, true);
    } else if (action === "remove-effect") {
      await mutate(tokenId, (t) => {
        t.odyssey.effects = t.odyssey.effects.filter((e) => e.id !== id);
      }, true);
    } else if (action === "use-item") {
      const item = host.tracker(token).odyssey.inventory.items.find((i) => i.id === id);
      if (!item) throw new Error("Item no longer exists.");
      const targets = host.getTokens().filter((t) => getRole() === "GM" || t.visible !== false);
      const partsFor = (target2) => Object.entries(host.tracker(target2).body).filter(([, p]) => getRole() === "GM" || !p.hidden).map(([k, p]) => [k, p.label || k]);
      const fields = [{ name: "target", label: "Target", value: "", options: [["", "Self"], ...targets.filter((t) => t.id !== tokenId).map((t) => [t.id, t.name || t.id])] }];
      if (item.subtype !== "stim") fields.push({ name: "part", label: "Body part", options: partsFor(token) });
      const value = await formDialog("Use " + item.name, fields, { onChange: (form, event) => {
        if (event.target.name !== "target" || !form.elements.part) return;
        const target2 = host.getToken(form.elements.target.value || tokenId);
        form.elements.part.innerHTML = target2 ? partsFor(target2).map(([key, label]) => '<option value="' + esc(key) + '">' + esc(label) + "</option>").join("") : "";
      } });
      if (!value) return;
      check(tokenId);
      const target = host.getToken(value.target || tokenId);
      if (!target) throw new Error("Target no longer exists.");
      await host.useItem(tokenId, id, target.id, value.part || "");
    } else if (action === "reload-weapon") {
      const index = Number(node.dataset.index), ammoId = readSelect(panel, "data-reload-kind", index);
      await mutate(tokenId, (t) => reloadMagazine(t.odyssey, index, ammoId), true);
    } else if (action === "delete-item") {
      const yes = await formDialog("Remove item?", [], { danger: true, message: "This removes the item from this token." });
      if (!yes) return;
      await mutate(tokenId, (t) => {
        unequipItem(t, id);
        t.odyssey.inventory.items = t.odyssey.inventory.items.filter((i) => i.id !== id);
      }, true);
    } else if (action === "consume") await mutate(tokenId, (t) => consumeItem(t, id));
    else if (action === "equip") {
      const part = readSelect(panel, "data-equip-part", id);
      await mutate(tokenId, (t) => equipItem(t, id, part, getRole() === "GM"));
    } else if (action === "unequip") await mutate(tokenId, (t) => unequipItem(t, id, getRole() === "GM"));
    else if (action === "add-mag") {
      const value = await formDialog("New magazine", [{ name: "name", label: "Name", required: true, value: "Magazine" }, { name: "capacity", label: "Capacity", type: "number", min: 1, max: 999, value: 6 }]);
      if (!value) return;
      const magazine = { id: newId(), name: value.name, capacity: Number(value.capacity), loaded: 0, loadedAmmoId: "" };
      await mutate(tokenId, (t) => {
        if (t.odyssey.inventory.magazines.length >= 200) throw new Error("Magazine limit is 200.");
        t.odyssey.inventory.magazines.push(magazine);
      }, true);
    } else if (action === "set-mag-rounds") {
      const tracker = host.tracker(check(tokenId, true)), o = tracker.odyssey;
      const mag = o.inventory.magazines.find((m) => m.id === id);
      if (!mag) throw new Error("Magazine no longer exists.");
      const weapon = o.weapons.ranged.find((w) => w.magazineId === id);
      const ammo = o.ammunition.filter((a) => !weapon || Array.isArray(weapon.acceptedMagazineIds) || weapon.ammoIds.includes(a.id));
      const value = await formDialog("GM \xB7 Set magazine rounds", [
        { name: "rounds", label: "Rounds (0\u2013" + mag.capacity + ")", type: "number", min: 0, max: mag.capacity, value: mag.loaded, required: true },
        { name: "ammoId", label: "Ammunition type", value: mag.loadedAmmoId, options: [["", "Empty / no ammunition"], ...ammo.map((a) => [a.id, a.name])] }
      ], { message: "Direct GM correction. Ammunition reserve is not changed; this is not a reload." });
      if (!value) return;
      await mutate(tokenId, (t) => setMagazineRounds(t.odyssey, id, Number(value.rounds), value.ammoId), true);
    } else if (action === "fill-mag") {
      const ammoId = readSelect(panel, "data-mag-ammo", id);
      await mutate(tokenId, (t) => fillMagazine(t.odyssey, id, ammoId), true);
    } else if (action === "empty-mag") await mutate(tokenId, (t) => emptyMagazine(t.odyssey, id), true);
    else if (action === "delete-mag") await mutate(tokenId, (t) => {
      emptyMagazine(t.odyssey, id);
      t.odyssey.inventory.magazines = t.odyssey.inventory.magazines.filter((m) => m.id !== id);
    }, true);
    else if (action === "mount-mag") {
      const index = Number(node.dataset.index), magId = readSelect(panel, "data-weapon-mag", index);
      await mutate(tokenId, (t) => swapMagazine(t.odyssey, index, magId), true);
    } else if (action === "save-template") {
      const value = await formDialog("Save character template", [{ name: "name", label: "Template name", required: true, value: token.name || "Character" }], { message: "Saves a healthy character. Owner, target, history and bridge events are excluded." });
      if (!value) return;
      if (storageError) throw new Error("Export the unreadable local library before importing a replacement. " + storageError);
      const current2 = check(tokenId);
      const template = makeTemplate(value.name, host.tracker(current2));
      saveLibrary(mergeLibraries(library, { ...emptyLibrary(), templates: [template] }));
    } else if (action === "export-library") {
      const json = storageError ? window.localStorage.getItem(LIBRARY_STORAGE_KEY) : JSON.stringify(library, null, 2);
      download("odyssey-templates.json", json || "{}");
      return;
    } else if (action === "import-library") {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json,application/json";
      input.onchange = async () => {
        try {
          const file = input.files?.[0];
          if (!file) return;
          if (file.size > 4 * 1024 * 1024) throw new Error("Library must be under 4 MB.");
          const incoming = normalizeLibrary(JSON.parse(await file.text()), sanitize);
          if (storageError) {
            const yes = await formDialog("Replace unreadable library?", [], { danger: true, message: "Export your existing JSON first. This replaces the unreadable local copy." });
            if (!yes) return;
            saveLibrary(incoming);
          } else saveLibrary(mergeLibraries(library, incoming));
          status("Template library imported.", "success");
        } catch (e) {
          status(e.message, "error");
        }
      };
      input.click();
      return;
    } else if (action === "apply-template") {
      const template = library.templates.find((t) => t.id === id);
      if (!template) throw new Error("Template no longer exists.");
      const yes = await formDialog(`Apply ${template.name}?`, [], { danger: true, message: "Replaces body, attributes, skills, weapons and inventory on this token. Destination owner and history are preserved. Saved attack target is cleared." });
      if (!yes) return;
      const source = check(tokenId, true), next = instantiateTemplate(host.tracker(source), template, sanitize);
      await mutate(tokenId, (t) => {
        t.enabled = next.enabled;
        t.minor = next.minor;
        t.serious = next.serious;
        t.body = next.body;
        t.odyssey = { ...next.odyssey, owner: t.odyssey.owner };
      }, true);
      host.resetDraft?.(tokenId);
    } else if (action === "rename-template" || action === "duplicate-template") {
      const t = library.templates.find((t2) => t2.id === id);
      if (!t) throw new Error("Template no longer exists.");
      const value = await formDialog(action === "rename-template" ? "Rename template" : "Duplicate template", [{ name: "name", label: "Name", required: true, value: t.name + (action === "duplicate-template" ? " copy" : "") }]);
      if (!value) return;
      const copy = { ...structuredClone(t), name: value.name.trim(), updatedAt: (/* @__PURE__ */ new Date()).toISOString() };
      const next = structuredClone(library);
      if (action === "duplicate-template") {
        copy.id = newId();
        next.templates.push(copy);
      } else next.templates = next.templates.map((x) => x.id === id ? copy : x);
      saveLibrary(normalizeLibrary(next, sanitize));
    } else if (action === "delete-template") {
      const yes = await formDialog("Remove template?", [], { danger: true, message: "Existing characters are unaffected." });
      if (!yes) return;
      saveLibrary({ ...library, templates: library.templates.filter((t) => t.id !== id) });
    } else if (action === "account-save") {
      if (!library.templates.length) throw new Error("Save at least one template first.");
      const items = library.templates.map((t) => buildText().name(t.name).plainText(t.name).position({ x: 0, y: 0 }).metadata({ [LIBRARY_ITEM_KEY]: { ...emptyLibrary(), templates: [t] } }).build());
      await OBR2.assets.uploadScenes([buildSceneUpload().name("Odyssey Templates").items(items).build()], true);
      status("Assets folder picker closed. Verify the library snapshot in your Owlbear account.", "info");
      return;
    } else if (action === "account-load") {
      const scenes = await OBR2.assets.downloadScenes(true, "Odyssey Templates");
      if (!scenes?.length) return;
      const incoming = libraryFromScenes(scenes, sanitize);
      if (storageError) {
        const yes = await formDialog("Replace unreadable library?", [], { danger: true, message: "Export your existing JSON before replacing it with the account snapshot." });
        if (!yes) return;
        saveLibrary(incoming);
      } else saveLibrary(mergeLibraries(library, incoming));
    }
    status("Saved.", "success");
  }
  document.addEventListener("click", (event) => {
    const node = event.target.closest?.("[data-od-action]");
    if (!node || node.disabled || busy) return;
    busy = true;
    document.querySelectorAll("[data-od-action]").forEach((e) => {
      if (e instanceof HTMLButtonElement) e.disabled = true;
    });
    handle(node.dataset.odAction, node).catch((e) => status(e.message, "error")).finally(() => {
      busy = false;
      render2();
      if (node.isConnected) node.disabled = false;
    });
  });
  const reload = () => {
    try {
      library = readLibrary(window.localStorage, sanitize);
      storageError = "";
      render2();
    } catch (e) {
      storageError = e.message;
      render2();
    }
  };
  window.addEventListener("storage", (event) => {
    if (event.key === LIBRARY_STORAGE_KEY) reload();
  });
  const disabledWhileBusy = (html) => busy ? html.replaceAll("<button ", "<button disabled ").replaceAll("<select ", "<select disabled ").replaceAll("<input ", "<input disabled ") : html;
  return { syncPool: schedulePoolScan, isBusy: () => busy, renderInventory: (tracker) => disabledWhileBusy(inventoryView(tracker, getRole() === "GM")), renderTemplates: () => getRole() !== "GM" ? "" : disabledWhileBusy(`${storageError ? `<p class="cp-banner err">Local library could not be read: ${esc(storageError)}. Export JSON before replacing it.</p>` : ""}${templatesView(library, getRole() === "GM")}`) };
}

// shared.js
var EXTENSION_ID = "com.codex.body-hp";
var META_KEY = `${EXTENSION_ID}/data`;
var OVERLAY_KEY = `${EXTENSION_ID}/overlayFor`;
var SHIELD_PART_NAME = "Shield";
var SPECIAL_PART_NAME = "Special";
var BODY_TOTAL_ORDER = ["Head", "L.Arm", "R.Arm", "Torso", "L.Leg", "R.Leg"];
var BODY_ORDER = [...BODY_TOTAL_ORDER, SHIELD_PART_NAME, SPECIAL_PART_NAME];
var BODY_PART_SLOTS = ["head", "left-arm", "right-arm", "left-leg", "right-leg", "other"];
var BODY_PART_SLOT_LABELS = {
  head: "Head",
  "left-arm": "Left Arm",
  "right-arm": "Right Arm",
  "left-leg": "Left Leg",
  "right-leg": "Right Leg",
  other: "Other"
};
var BODY_SLOT_BY_PART = {
  Head: "head",
  "L.Arm": "left-arm",
  "R.Arm": "right-arm",
  "L.Leg": "left-leg",
  "R.Leg": "right-leg",
  Torso: "torso"
};
var BODY_SLOT_LAYOUT = [
  { slot: "head", angle: -90 },
  { slot: "right-arm", angle: -18 },
  { slot: "right-leg", angle: 54 },
  { slot: "left-leg", angle: 126 },
  { slot: "left-arm", angle: 198 }
];
var ROLL_HISTORY_LIMIT = 12;
var COMBAT_SKILL_CATEGORY = "combat";
var APPLIED_SKILL_CATEGORY = "applied";
var ABILITIES_SKILL_CATEGORY = "abilities";
var MELEE_SKILL_NAME = "Melee";
var PARRY_SKILL_NAME = "Parry";
var LEGACY_MELEE_SKILL_NAMES = /* @__PURE__ */ new Set(["Hand", "Cold", "\u0420\u0443\u043A\u043E\u043F\u0430\u0448\u043D\u044B\u0439"]);
var LEGACY_REMOVED_SKILLS = /* @__PURE__ */ new Set(["Hand", "Cold", "Throwing", "Rifle", "Turrets"]);
var VISUAL_VERSION = 18;
var SPECIAL_RING_COLOR = "#57D8FF";
var HP_COLOR_STOPS = [
  { ratio: 1, color: "#73FF5A" },
  { ratio: 0.75, color: "#FFF243" },
  { ratio: 0.5, color: "#FFAF22" },
  { ratio: 0.25, color: "#AC0004" },
  { ratio: 0, color: "#000000" }
];
var RING_COLORS = {
  base: "#000000",
  border: "#050505"
};
var OVERLAY_STROKE_WIDTH = 0.75;
var overlayEnsureQueue = /* @__PURE__ */ new Map();
var OVERLAY_UPDATE_DELAY_MS = 75;
var cachedGridDpi = null;
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
  Torso: { current: 3, max: 3, armor: 6, minor: 0, serious: 0, slot: "torso", attackPenalty: 0, hidden: false, label: "Torso" },
  [SHIELD_PART_NAME]: { current: 0, max: 0, armor: 0, minor: 0, serious: 0 },
  [SPECIAL_PART_NAME]: { current: 0, max: 0, armor: 0, minor: 0, serious: 0 }
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
    attackDraft: {
      targetTokenId: "",
      targetTokenName: ""
    },
    skills: structuredClone(DEFAULT_ODYSSEY_SKILLS),
    skillCategories: structuredClone(DEFAULT_ODYSSEY_SKILL_CATEGORIES),
    skillStrengthBonuses: structuredClone(DEFAULT_ODYSSEY_SKILL_STRENGTH_BONUSES),
    skillLabels: {},
    skillPerks: {},
    skillRangeReduction: {},
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
    inventory: { version: 3, items: [], magazines: [], baseArmor: {}, baseMax: {} },
    effects: [],
    preferences: { parryMode: "off" },
    ammunition: [],
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
function enableActionResize(handle, { storageKey, defaultWidth, defaultHeight }) {
  if (!(handle instanceof HTMLElement)) return;
  const readSavedSize = () => {
    try {
      return JSON.parse(localStorage.getItem(storageKey) ?? "null");
    } catch {
      return null;
    }
  };
  const savedSize = readSavedSize();
  const initialWidth = clamp(Number(savedSize?.width) || defaultWidth, 360, 1200);
  const initialHeight = clamp(Number(savedSize?.height) || defaultHeight, 500, 1200);
  let currentWidth = initialWidth;
  let currentHeight = initialHeight;
  void lib_default.action.setWidth(initialWidth);
  void lib_default.action.setHeight(initialHeight);
  handle.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    handle.setPointerCapture?.(event.pointerId);
    const startWidth = currentWidth;
    const startHeight = currentHeight;
    const startX = event.clientX;
    const startY = event.clientY;
    let nextWidth = startWidth;
    let nextHeight = startHeight;
    let frame = null;
    const applySize = () => {
      frame = null;
      currentWidth = nextWidth;
      currentHeight = nextHeight;
      void lib_default.action.setWidth(nextWidth);
      void lib_default.action.setHeight(nextHeight);
    };
    const move = (moveEvent) => {
      nextWidth = clamp(startWidth + moveEvent.clientX - startX, 360, 1200);
      nextHeight = clamp(startHeight + moveEvent.clientY - startY, 500, 1200);
      if (frame == null) frame = requestAnimationFrame(applySize);
    };
    const stop = () => {
      if (frame != null) {
        cancelAnimationFrame(frame);
        applySize();
      }
      try {
        localStorage.setItem(storageKey, JSON.stringify({ width: nextWidth, height: nextHeight }));
      } catch {
      }
      if (handle.hasPointerCapture?.(event.pointerId)) handle.releasePointerCapture(event.pointerId);
      document.removeEventListener("pointermove", move);
      document.removeEventListener("pointerup", stop);
    };
    document.addEventListener("pointermove", move);
    document.addEventListener("pointerup", stop, { once: true });
  });
}
function numberOrFallback(value, fallback) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}
function getSlotAttackPenalty(slot) {
  if (slot === "head") return 30;
  if (["left-arm", "right-arm", "left-leg", "right-leg"].includes(slot)) return 15;
  return 0;
}
function getPartSlot(partName, source) {
  if (partName === "Torso") return "torso";
  return BODY_PART_SLOTS.includes(source?.slot) ? source.slot : BODY_SLOT_BY_PART[partName] ?? "other";
}
function sanitizeTrackerData(raw) {
  const next = deepClone(DEFAULT_TRACKER_DATA);
  if (!raw || typeof raw !== "object") raw = {};
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
  for (const partName of Object.keys(BODY_DEFAULTS)) {
    const source = raw.body?.[partName] ?? {};
    const part = next.body[partName];
    part.max = clamp(numberOrFallback(source.max, part.max), 0, 999);
    part.current = clamp(
      numberOrFallback(source.current, part.current),
      0,
      part.max
    );
    part.armor = clamp(numberOrFallback(source.armor, part.armor), 0, 999);
    part.minor = clamp(numberOrFallback(source.minor, part.minor), 0, 3);
    part.serious = clamp(numberOrFallback(source.serious, part.serious), 0, 1);
    part.slot = getPartSlot(partName, source);
    part.attackPenalty = clamp(numberOrFallback(source.attackPenalty, getSlotAttackPenalty(part.slot)), 0, 999);
    if (part.slot !== "other") part.attackPenalty = getSlotAttackPenalty(part.slot);
    part.hidden = partName === "Torso" ? false : source.hidden === true;
    part.label = String(source.label ?? partName).trim().slice(0, 40) || partName;
  }
  for (const [partName, source] of Object.entries(raw.body ?? {})) {
    if (Object.hasOwn(next.body, partName)) continue;
    const name = String(partName).trim().slice(0, 40);
    if (!name) continue;
    const max = clamp(Number(source?.max) || 0, 0, 999);
    next.body[name] = {
      current: clamp(Number(source?.current) || 0, 0, max),
      max,
      armor: clamp(Number(source?.armor) || 0, 0, 999),
      minor: clamp(Number(source?.minor) || 0, 0, 3),
      serious: clamp(Number(source?.serious) || 0, 0, 1),
      slot: getPartSlot(name, source),
      attackPenalty: clamp(numberOrFallback(source?.attackPenalty, getSlotAttackPenalty(getPartSlot(name, source))), 0, 999),
      hidden: source?.hidden === true,
      label: String(source?.label ?? name).trim().slice(0, 40) || name
    };
    if (next.body[name].slot !== "other") {
      next.body[name].attackPenalty = getSlotAttackPenalty(next.body[name].slot);
    }
  }
  reconcileEquipment(next);
  return next;
}
function sanitizeOdysseyData(raw) {
  var _a;
  const next = deepClone(DEFAULT_TRACKER_DATA.odyssey);
  if (!raw || typeof raw !== "object") raw = {};
  next.owner.playerId = String(raw.owner?.playerId ?? "").trim();
  next.owner.playerName = String(raw.owner?.playerName ?? "").trim();
  next.attackDraft.targetTokenId = String(raw.attackDraft?.targetTokenId ?? "").trim();
  next.attackDraft.targetTokenName = String(raw.attackDraft?.targetTokenName ?? "").trim();
  const rawSkills = raw.skills && typeof raw.skills === "object" ? raw.skills : {};
  const rawSkillCategories = raw.skillCategories && typeof raw.skillCategories === "object" ? raw.skillCategories : {};
  const rawSkillStrengthBonuses = raw.skillStrengthBonuses && typeof raw.skillStrengthBonuses === "object" ? raw.skillStrengthBonuses : {};
  const rawSkillLabels = raw.skillLabels && typeof raw.skillLabels === "object" ? raw.skillLabels : {};
  const rawSkillPerks = raw.skillPerks && typeof raw.skillPerks === "object" ? raw.skillPerks : {};
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
    next.skillCategories[normalizedKey] = categoryValue === COMBAT_SKILL_CATEGORY ? COMBAT_SKILL_CATEGORY : categoryValue === ABILITIES_SKILL_CATEGORY ? ABILITIES_SKILL_CATEGORY : APPLIED_SKILL_CATEGORY;
    next.skillStrengthBonuses[normalizedKey] = Boolean(
      rawSkillStrengthBonuses[normalizedKey] ?? rawSkillStrengthBonuses[key] ?? false
    );
  }
  for (const key of Object.keys(next.attributes)) {
    const fallbackValue = key === "Magic" ? raw.attributes?.[key] ?? raw.attributes?.Psionics ?? 0 : raw.attributes?.[key] ?? 0;
    next.attributes[key] = clamp(Number(fallbackValue) || 0, 0, 20);
  }
  next.skillRangeReduction = {};
  for (const skillName of Object.keys(next.skills)) {
    const reduction = sanitizeMagazine({ rangeReduction: raw.skillRangeReduction?.[skillName] }).rangeReduction;
    if (Object.keys(reduction).length) next.skillRangeReduction[skillName] = reduction;
    next.skillLabels[skillName] = String(rawSkillLabels[skillName] ?? skillName).trim().slice(0, 40) || skillName;
    next.skillPerks[skillName] = Array.isArray(rawSkillPerks[skillName]) ? rawSkillPerks[skillName].map((perk) => String(perk ?? "").trim().slice(0, 40)).filter(Boolean).slice(0, 12) : [];
  }
  next.weapons.melee = sanitizeWeapons(raw.weapons?.melee);
  next.weapons.ranged = sanitizeWeapons(raw.weapons?.ranged, true);
  next.ammunition = sanitizeAmmunition(raw.ammunition);
  next.inventory = normalizeInventory(raw.inventory, next.weapons.ranged);
  next.effects = normalizeEffects(raw.effects);
  next.preferences = { parryMode: ["off", "1", "2", "3", "4", "5"].includes(raw.preferences?.parryMode) ? raw.preferences.parryMode : "off" };
  for (const [key, label] of Object.entries(ARMOR_TRAINING)) {
    next.skills[key] = clamp(Number(raw.skills?.[key]) || 0, 0, 1);
    next.skillLabels[key] = label;
    next.skillCategories[key] = "passive";
    next.skillStrengthBonuses[key] = false;
    (_a = next.skillPerks)[key] ?? (_a[key] = []);
  }
  return next;
}
function sanitizeWeapons(raw, ranged = false) {
  if (!Array.isArray(raw)) return [];
  return raw.filter((item) => item && typeof item === "object").map((item, index) => ({
    id: String(item.id || `legacy-${ranged ? "ranged" : "melee"}-${index}`),
    name: String(item.name ?? "").trim() || "Weapon",
    damage: clamp(Number(item.damage ?? 0) || 0, -999, 999),
    accuracy: clamp(Number(item.accuracy ?? 0) || 0, -999, 999),
    ...ranged ? { ...Object.hasOwn(item, "magazineId") ? { magazineId: String(item.magazineId || "") } : {}, ...sanitizeMagazine(item), id: String(item.id || `legacy-ranged-${index}`) } : {}
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
  const lines = [];
  if (hasConfiguredShield(data)) {
    lines.push(
      `${SHIELD_PART_NAME} ${body[SHIELD_PART_NAME].current}/${body[SHIELD_PART_NAME].max}(${body[SHIELD_PART_NAME].armor})`
    );
  }
  if (hasConfiguredSpecial(data)) {
    lines.push(
      `${SPECIAL_PART_NAME} ${body[SPECIAL_PART_NAME].current}/${body[SPECIAL_PART_NAME].max}(${body[SPECIAL_PART_NAME].armor})`
    );
  }
  const parts = getBodyPartNames(data).map((partName) => {
    const part = body[partName];
    return `${getBodyPartLabel(data, partName)} ${part.current}/${part.max}(${part.armor})`;
  });
  for (let index = 0; index < parts.length; index += 3) {
    lines.push(parts.slice(index, index + 3).join(" | "));
  }
  return lines.join("\n");
}
function getOdysseyData(item) {
  return effectiveOdyssey(sanitizeOdysseyData(getTrackerData(item).odyssey));
}
function canPlayerControlToken(playerRole2, playerId2, token) {
  if (!token || !isCharacterToken(token)) return false;
  if (playerRole2 === "GM") return true;
  const odyssey = getOdysseyData(token);
  return Boolean(playerId2) && odyssey.owner.playerId === playerId2;
}
function getBodyTotals(data) {
  return getBodyPartNames(data).reduce(
    (accumulator, partName) => {
      accumulator.current += data.body[partName].current;
      accumulator.max += data.body[partName].max;
      return accumulator;
    },
    { current: 0, max: 0 }
  );
}
function getArmorTotal(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  return Object.values(body).reduce((total, part) => total + (Number(part?.armor) || 0), 0);
}
function getBodyPartNames(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  const names = Object.keys(body).filter((partName) => partName !== SHIELD_PART_NAME && partName !== SPECIAL_PART_NAME);
  return names.sort((left, right) => {
    if (left === "Torso") return -1;
    if (right === "Torso") return 1;
    const leftIndex = BODY_TOTAL_ORDER.indexOf(left);
    const rightIndex = BODY_TOTAL_ORDER.indexOf(right);
    if (leftIndex >= 0 || rightIndex >= 0) return (leftIndex < 0 ? 99 : leftIndex) - (rightIndex < 0 ? 99 : rightIndex);
    return left.localeCompare(right);
  });
}
function getBodyPartLabel(dataOrBody, partName) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  return String(body?.[partName]?.label ?? partName).trim() || partName;
}
function hasConfiguredShield(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody;
  const shield = body?.[SHIELD_PART_NAME];
  if (!shield || typeof shield !== "object") return false;
  return (Number(shield.max) || 0) > 0 || (Number(shield.current) || 0) > 0 || (Number(shield.armor) || 0) > 0;
}
function hasConfiguredSpecial(dataOrBody) {
  const body = dataOrBody?.body ?? dataOrBody;
  const special = body?.[SPECIAL_PART_NAME];
  if (!special || typeof special !== "object") return false;
  return (Number(special.max) || 0) > 0 || (Number(special.current) || 0) > 0 || (Number(special.armor) || 0) > 0;
}
function getTargetableBodyParts(dataOrBody, includeHidden = false) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  const parts = getBodyPartNames(body).filter(
    (partName) => includeHidden || body[partName]?.hidden !== true
  );
  return parts.length ? parts : ["Torso"];
}
function getBodyPartAttackPenalty(dataOrBody, partName) {
  const body = dataOrBody?.body ?? dataOrBody ?? {};
  const part = body?.[partName];
  const slot = getPartSlot(partName, part);
  return slot === "other" ? clamp(numberOrFallback(part?.attackPenalty, 0), 0, 999) : getSlotAttackPenalty(slot);
}
function getEffectiveSize(token, sceneGridDpi) {
  const scaleX = Math.abs(token.scale?.x ?? 1);
  const scaleY = Math.abs(token.scale?.y ?? 1);
  const imageWidth = Number(token.image?.width);
  const imageHeight = Number(token.image?.height);
  const imageGridDpi = Number(token.grid?.dpi);
  const baseWidth = imageWidth > 0 && imageGridDpi > 0 ? imageWidth / imageGridDpi * sceneGridDpi : Number(token.width) || sceneGridDpi;
  const baseHeight = imageHeight > 0 && imageGridDpi > 0 ? imageHeight / imageGridDpi * sceneGridDpi : Number(token.height) || sceneGridDpi;
  return {
    width: baseWidth * scaleX,
    height: baseHeight * scaleY
  };
}
async function getTokenMetrics(token, data) {
  const gridDpi = await getCachedGridDpi();
  const effectiveSize = getEffectiveSize(token, gridDpi);
  const center = token.position;
  const width = effectiveSize.width;
  const height = effectiveSize.height;
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
  const otherParts = getBodyPartNames(data).filter(
    (partName) => data.body?.[partName]?.slot === "other" && data.body?.[partName]?.hidden !== true
  );
  const extraInnerRadius = outerRadius + Math.max(4, visibleDiameter * 0.03);
  const extraOuterRadius = extraInnerRadius + outerThickness;
  const specialThickness = Math.max(4, visibleDiameter * 0.03);
  const specialInnerRadius = otherParts.length ? extraOuterRadius : outerRadius;
  const specialOuterRadius = specialInnerRadius + specialThickness;
  const shieldThickness = Math.max(4, visibleDiameter * 0.028);
  const shieldOuterRadius = Math.max(10, visibleDiameter * 0.1);
  const shieldInnerRadius = Math.max(4, shieldOuterRadius - shieldThickness);
  const shieldOffsetY = -(specialOuterRadius + shieldOuterRadius + Math.max(5, visibleDiameter * 0.035));
  return {
    center,
    visibleDiameter,
    outerRadius,
    outerInnerRadius,
    extraInnerRadius,
    extraOuterRadius,
    torsoOuterRadius,
    torsoInnerRadius,
    specialOuterRadius,
    specialInnerRadius,
    shieldOuterRadius,
    shieldInnerRadius,
    shieldOffsetY
  };
}
function getOverlayPartLayout(data, metrics) {
  const visibleParts = getBodyPartNames(data).filter(
    (partName) => partName !== "Torso" && data.body?.[partName]?.hidden !== true
  );
  const layouts = [];
  for (const { slot, angle } of BODY_SLOT_LAYOUT) {
    const parts = visibleParts.filter((partName) => data.body?.[partName]?.slot === slot);
    const span = 60 / Math.max(parts.length, 1);
    parts.forEach((partName, index) => {
      layouts.push({
        partName,
        angle: angle - 30 + span * (index + 0.5),
        span,
        innerRadius: metrics.outerInnerRadius,
        outerRadius: metrics.outerRadius
      });
    });
  }
  const otherParts = visibleParts.filter((partName) => data.body?.[partName]?.slot === "other");
  const otherSpan = Math.min(60, 360 / Math.max(otherParts.length, 1) * 0.8);
  otherParts.forEach((partName, index) => {
    layouts.push({
      partName,
      angle: -90 + otherSpan * index,
      span: otherSpan,
      innerRadius: metrics.extraInnerRadius,
      outerRadius: metrics.extraOuterRadius
    });
  });
  return layouts;
}
async function getCachedGridDpi(forceRefresh = false) {
  if (!forceRefresh && Number.isFinite(cachedGridDpi) && cachedGridDpi > 0) {
    return cachedGridDpi;
  }
  let gridDpi = 150;
  try {
    gridDpi = await lib_default.scene.grid.getDpi() || gridDpi;
  } catch (error) {
    console.warn("[Body HP] Unable to read grid dpi, using fallback size", error);
  }
  cachedGridDpi = Math.max(1, Number(gridDpi) || 150);
  return cachedGridDpi;
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
function buildAnnulusCommands(radiusOuter, radiusInner, offsetX = 0, offsetY = 0) {
  const outer = arcPoints(radiusOuter, -180, 180, 36).map((point) => ({
    x: point.x + offsetX,
    y: point.y + offsetY
  }));
  const inner = arcPoints(radiusInner, -180, 180, 36).map((point) => ({
    x: point.x + offsetX,
    y: point.y + offsetY
  }));
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
  const segments = Math.max(10, Math.ceil(Math.abs(spanAngle) / 5));
  const outer = arcPoints(radiusOuter, startAngle, endAngle, segments);
  const inner = arcPoints(radiusInner, endAngle, startAngle, segments);
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
function hexToRgb(hex) {
  const normalized = String(hex).replace("#", "");
  const value = normalized.length === 3 ? normalized.split("").map((char) => `${char}${char}`).join("") : normalized;
  return {
    r: Number.parseInt(value.slice(0, 2), 16) || 0,
    g: Number.parseInt(value.slice(2, 4), 16) || 0,
    b: Number.parseInt(value.slice(4, 6), 16) || 0
  };
}
function rgbToHex({ r, g, b }) {
  return `#${[r, g, b].map(
    (channel) => clamp(Math.round(channel), 0, 255).toString(16).padStart(2, "0").toUpperCase()
  ).join("")}`;
}
function mixHexColors(startHex, endHex, ratio) {
  const safeRatio = clamp(ratio, 0, 1);
  const start = hexToRgb(startHex);
  const end = hexToRgb(endHex);
  return rgbToHex({
    r: start.r + (end.r - start.r) * safeRatio,
    g: start.g + (end.g - start.g) * safeRatio,
    b: start.b + (end.b - start.b) * safeRatio
  });
}
function getHpColor(ratio) {
  const safeRatio = clamp(ratio, 0, 1);
  for (let index = 0; index < HP_COLOR_STOPS.length - 1; index += 1) {
    const upper = HP_COLOR_STOPS[index];
    const lower = HP_COLOR_STOPS[index + 1];
    if (safeRatio > upper.ratio || safeRatio < lower.ratio) {
      continue;
    }
    const span = upper.ratio - lower.ratio;
    if (span <= 0) return upper.color;
    const progress = (safeRatio - lower.ratio) / span;
    return mixHexColors(lower.color, upper.color, progress);
  }
  return HP_COLOR_STOPS.at(-1)?.color ?? RING_COLORS.base;
}
function getPartColor(part) {
  if (part.max <= 0) {
    return (Number(part?.armor) || 0) > 0 ? getHpColor(1) : getHpColor(0);
  }
  return getHpColor(part.current / part.max);
}
function getSpecialPartColor(part) {
  const ratio = (Number(part?.max) || 0) > 0 ? clamp((Number(part?.current) || 0) / (Number(part?.max) || 1), 0, 1) : (Number(part?.current) || 0) > 0 || (Number(part?.armor) || 0) > 0 ? 1 : 0;
  return mixHexColors("#000000", SPECIAL_RING_COLOR, ratio);
}
function buildRingItem(token, metrics, kind, commands, fillColor, zIndex = 0, fillRule = "nonzero", signature = "", itemVisible = true) {
  return buildPath().name(`${kind}: ${getCharacterName(token)}`).commands(commands).fillRule(fillRule).fillColor(fillColor).fillOpacity(1).strokeColor(RING_COLORS.border).strokeOpacity(1).strokeWidth(OVERLAY_STROKE_WIDTH).position(metrics.center).rotation(0).zIndex((token.zIndex ?? 0) + 100 + zIndex).visible(itemVisible && token.visible !== false).attachedTo(token.id).disableAttachmentBehavior(["ROTATION"]).layer("ATTACHMENT").locked(true).disableHit(true).metadata({
    [OVERLAY_KEY]: token.id,
    kind,
    visualVersion: VISUAL_VERSION,
    signature
  }).build();
}
function applyOverlayItemState(target, source) {
  target.name = source.name;
  target.commands = source.commands;
  target.fillRule = source.fillRule;
  target.fillColor = source.fillColor;
  target.fillOpacity = source.fillOpacity;
  target.strokeColor = source.strokeColor;
  target.strokeOpacity = source.strokeOpacity;
  target.strokeWidth = source.strokeWidth;
  target.position = source.position;
  target.rotation = source.rotation;
  target.zIndex = source.zIndex;
  target.visible = source.visible;
  target.metadata = {
    ...target.metadata ?? {},
    ...source.metadata ?? {}
  };
}
function hasPatchableOverlaySet(token, overlayItems, expectedKinds) {
  if (overlayItems.length !== expectedKinds.length) {
    return false;
  }
  const seenKinds = /* @__PURE__ */ new Set();
  return overlayItems.every((item) => {
    const kind = String(item.metadata?.kind ?? "");
    const valid = item.attachedTo === token.id && Number(item.metadata?.visualVersion ?? 0) === VISUAL_VERSION && expectedKinds.includes(kind) && !seenKinds.has(kind);
    seenKinds.add(kind);
    return valid;
  });
}
function roundMetric(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}
function buildOverlaySignature(token, data, metrics) {
  const bodySignature = getBodyPartNames(data).map((partName) => {
    const part = data.body?.[partName] ?? {};
    return `${partName}:${part.slot ?? "other"}:${part.hidden === true}:${getPartColor(part)}`;
  }).join("|");
  return [
    VISUAL_VERSION,
    roundMetric(metrics.visibleDiameter),
    roundMetric(metrics.outerRadius),
    roundMetric(metrics.outerInnerRadius),
    roundMetric(metrics.extraInnerRadius),
    roundMetric(metrics.extraOuterRadius),
    roundMetric(metrics.torsoOuterRadius),
    roundMetric(metrics.torsoInnerRadius),
    roundMetric(metrics.specialOuterRadius),
    roundMetric(metrics.specialInnerRadius),
    hasConfiguredSpecial(data),
    bodySignature
  ].join(";");
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
function buildOverlayItems(token, data, metrics, signature = "") {
  const items = [];
  const makeSignature = (kind, commands, color, visible) => `${VISUAL_VERSION}|${kind}|${roundMetric(metrics.visibleDiameter)}|${roundMetric(metrics.center.x)}|${roundMetric(metrics.center.y)}|${color}|${visible}|${JSON.stringify(commands)}`;
  for (const segment of getOverlayPartLayout(data, metrics)) {
    const part = data.body[segment.partName];
    const kind = `part-${segment.partName}`;
    const commands = buildSectorCommands(
      segment.outerRadius,
      segment.innerRadius,
      segment.angle,
      segment.span
    );
    items.push(buildRingItem(token, metrics, kind, commands, getPartColor(part), 1, "nonzero", makeSignature(kind, commands, getPartColor(part), true), true));
  }
  const torsoCommands = buildAnnulusCommands(metrics.torsoOuterRadius, metrics.torsoInnerRadius);
  items.push(buildRingItem(token, metrics, "torso-ring", torsoCommands, getPartColor(data.body.Torso), 2, "evenodd", makeSignature("torso-ring", torsoCommands, getPartColor(data.body.Torso), true), true));
  if (hasConfiguredSpecial(data)) {
    const specialCommands = buildAnnulusCommands(metrics.specialOuterRadius, metrics.specialInnerRadius);
    items.push(buildRingItem(token, metrics, "special-ring", specialCommands, getSpecialPartColor(data.body[SPECIAL_PART_NAME]), 3, "evenodd", makeSignature("special-ring", specialCommands, getSpecialPartColor(data.body[SPECIAL_PART_NAME]), true), true));
  }
  return items;
}
function getExpectedOverlayKinds(data) {
  return [
    ...getBodyPartNames(data).filter((partName) => partName !== "Torso" && data.body?.[partName]?.hidden !== true).map((partName) => `part-${partName}`),
    "torso-ring",
    ...hasConfiguredSpecial(data) ? ["special-ring"] : []
  ];
}
async function removeOverlaysForToken(tokenId, items) {
  const sceneItems2 = items ?? await lib_default.scene.items.getItems();
  const overlayIds = sceneItems2.filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId).map((item) => item.id);
  if (overlayIds.length) {
    await lib_default.scene.items.deleteItems(overlayIds);
  }
}
async function ensureOverlayForTokenInternal(tokenId, items) {
  const sceneItems2 = items ?? await lib_default.scene.items.getItems();
  const token = sceneItems2.find((item) => item.id === tokenId);
  if (!token || !isCharacterToken(token)) return;
  const overlayItems = sceneItems2.filter((item) => item.metadata?.[OVERLAY_KEY] === tokenId);
  if (!isTrackedCharacter(token) || token.visible === false) {
    if (overlayItems.length) {
      await removeOverlaysForToken(tokenId, sceneItems2);
    }
    return;
  }
  const data = getTrackerData(token);
  const metrics = await getTokenMetrics(token, data);
  const overlaySignature = buildOverlaySignature(token, data, metrics);
  const expectedKinds = getExpectedOverlayKinds(data);
  if (hasPatchableOverlaySet(token, overlayItems, expectedKinds)) {
    const nextOverlayItems = buildOverlayItems(token, data, metrics, overlaySignature);
    const nextOverlayByKind = new Map(
      nextOverlayItems.map((item) => [String(item.metadata?.kind ?? ""), item])
    );
    const itemsToUpdate = overlayItems.filter((item) => {
      const kind = String(item.metadata?.kind ?? "");
      const nextItem = nextOverlayByKind.get(kind);
      return String(item.metadata?.signature ?? "") !== String(nextItem?.metadata?.signature ?? "");
    });
    if (!itemsToUpdate.length) return;
    try {
      await lib_default.scene.items.updateItems(
        itemsToUpdate.map((item) => item.id),
        (sceneItemsToUpdate) => {
          for (const overlayItem of sceneItemsToUpdate) {
            const kind = String(overlayItem.metadata?.kind ?? "");
            const nextItem = nextOverlayByKind.get(kind);
            if (!nextItem) continue;
            applyOverlayItemState(overlayItem, nextItem);
          }
        }
      );
      return;
    } catch (error) {
      console.warn("[Body HP] Overlay patch failed, falling back to rebuild", error);
    }
  }
  await removeOverlaysForToken(tokenId);
  await lib_default.scene.items.addItems(
    buildOverlayItems(token, data, metrics, overlaySignature)
  );
}
function ensureOverlayForToken(tokenId) {
  let pending = overlayEnsureQueue.get(tokenId);
  if (!pending) {
    pending = { timer: null, running: false, waiters: [] };
    overlayEnsureQueue.set(tokenId, pending);
  }
  return new Promise((resolve, reject) => {
    pending.waiters.push({ resolve, reject });
    if (pending.running) return;
    if (pending.timer) clearTimeout(pending.timer);
    pending.timer = setTimeout(() => {
      pending.timer = null;
      void flushOverlayEnsure(tokenId, pending);
    }, OVERLAY_UPDATE_DELAY_MS);
  });
}
async function flushOverlayEnsure(tokenId, pending) {
  if (pending.running) return;
  pending.running = true;
  const waiters = pending.waiters.splice(0);
  try {
    await ensureOverlayForTokenInternal(tokenId);
    waiters.forEach(({ resolve }) => resolve());
  } catch (error) {
    waiters.forEach(({ reject }) => reject(error));
  } finally {
    pending.running = false;
    if (pending.waiters.length) {
      pending.timer = setTimeout(() => {
        pending.timer = null;
        void flushOverlayEnsure(tokenId, pending);
      }, OVERLAY_UPDATE_DELAY_MS);
    } else {
      overlayEnsureQueue.delete(tokenId);
    }
  }
}
async function syncTrackedOverlays() {
  const items = await lib_default.scene.items.getItems();
  const overlayIds = items.filter(isOverlayItem).map((item) => item.id);
  if (overlayIds.length) {
    await lib_default.scene.items.deleteItems(overlayIds);
  }
  const trackedTokens = items.filter((item) => isTrackedCharacter(item) && item.visible !== false);
  for (const token of trackedTokens) {
    await ensureOverlayForToken(token.id);
  }
}

// odyssey_rules.js
function rollPercent() {
  return Math.floor(Math.random() * 100) + 1;
}
function rollDice(sides, modifier2 = 0, count = 1) {
  const safeSides = clamp(Number(sides) || 0, 2, Number.MAX_SAFE_INTEGER);
  const safeCount = clamp(Number(count) || 0, 1, 100);
  const rolls = Array.from({ length: safeCount }, () => Math.floor(Math.random() * safeSides) + 1);
  const subtotal = rolls.reduce((sum, roll) => sum + roll, 0);
  return {
    roll: rolls[0] ?? 0,
    rolls,
    count: safeCount,
    sides: safeSides,
    subtotal,
    modifier: Number(modifier2) || 0,
    total: subtotal + (Number(modifier2) || 0)
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
    crit = Math.ceil(damageDiff / 30) - 1;
    label = `Critical damage: ${crit} Crit.`;
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
function getAttackOutcomeIcon(outcome) {
  switch (outcome) {
    case "critical-success":
      return "\u{1F3AF}";
    case "success":
      return "\u2705";
    case "critical-failure":
      return "\u{1F480}";
    default:
      return "\u274C";
  }
}
function formatAttackOutcomeLabel(outcome) {
  switch (outcome) {
    case "critical-success":
      return `${getAttackOutcomeIcon(outcome)} Critical Success`;
    case "success":
      return `${getAttackOutcomeIcon(outcome)} Success`;
    case "critical-failure":
      return `${getAttackOutcomeIcon(outcome)} Critical Failure`;
    default:
      return `${getAttackOutcomeIcon(outcome)} Failure`;
  }
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
  const requestedPart = String(targetPart ?? "").trim();
  const part = requestedPart && requestedPart !== SPECIAL_PART_NAME ? requestedPart : "Torso";
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
    const crit = Math.max(baseDamage.crit || 0, 2);
    damage = {
      ...baseDamage,
      label: `Critical hit: ${crit} Crit.`,
      crit,
      serious: 0,
      minor: 0
    };
    bodyDelta = -crit;
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
    return `${getAttackOutcomeIcon(outcome)} Critical success to ${part}. Roll ${attackRoll}; ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  if (outcome === "critical-failure") {
    return `${getAttackOutcomeIcon(outcome)} Critical failure. Roll ${attackRoll}.`;
  }
  if (outcome === "success") {
    return `${getAttackOutcomeIcon(outcome)} Hit ${part}. ${attackTotal} vs ${defenseTotal}. ${damage?.label ?? ""}`.trim();
  }
  return `${getAttackOutcomeIcon(outcome)} Missed ${part}. ${attackTotal} vs ${defenseTotal}.`;
}

// main.js
var DEBUG_LOG_KEY = "com.codex.body-hp/debugLog";
var DEBUG_BROADCAST_CHANNEL = "com.codex.body-hp/debug";
var DEBUG_ENTRY_LIMIT = 50;
var TARGET_PICK_TOOL_ID = "com.codex.body-hp/attack-target-picker";
var TARGET_PICK_MODE_ID = "pick-target";
var TARGET_HIGHLIGHT_KEY = "com.codex.body-hp/local-attack-target";
var LOCAL_SELF_VIEW_KEY = "com.codex.body-hp/local-self-view";
var ATTACK_TARGET_CONTEXT_MENU_ID = "com.codex.body-hp/set-attack-target";
var SHOW_EMBEDDED_PUBLIC_LOG = false;
var EXTENSION_ICON_URL = new URL("./icon.svg", window.location.href).href;
var TARGET_PICK_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
    <circle cx="16" cy="16" r="3.25" fill="#ef4444" stroke="#7f1d1d" stroke-width="1.5"/>
    <path d="M16 2v8M16 22v8M2 16h8M22 16h8" stroke="#ef4444" stroke-width="2.5" stroke-linecap="round"/>
    <path d="M16 6v4M16 22v4M6 16h4M22 16h4" stroke="#7f1d1d" stroke-width="1.2" stroke-linecap="round"/>
  </svg>`
)}") 16 16, crosshair`;
var HUD_MODE = new URLSearchParams(window.location.search).get("hud") === "1";
var measuredTargetDistance = null;
var targetDistance = null;
function paintTargetDistance(reading) {
  measuredTargetDistance = reading;
  for (const node of document.querySelectorAll("[data-target-distance]")) node.textContent = reading?.sourceId === activeTokenId ? reading.label : "\u2014";
  paintRangeAdjustments();
}
function paintRangeAdjustments() {
  const token = getCharacterById(activeTokenId);
  if (!token) return;
  const odyssey = getOdysseyData(token);
  for (const node of document.querySelectorAll("[data-range-modifier]")) {
    const manual = node.dataset.rangeModifier === "manual", prefix = manual ? "data-manual-attack-field" : "data-attack-field";
    const weapon = getWeaponByName(token, getActionFieldValue(`[${prefix}="weaponName"]`));
    const skill = getActionFieldValue(`[${prefix}="skill"]`), raw = getActionFieldValue('[data-manual-attack-field="distanceMeters"]');
    const meters = manual ? String(raw ?? "").trim() === "" ? null : Number(raw) : measuredTargetDistance?.sourceId === token.id ? measuredTargetDistance.meters : null;
    const result = rangeAdjustment(weapon, meters, odyssey, skill);
    node.hidden = !result.category;
    node.textContent = rangeAdjustmentLabel(result);
    node.classList.toggle("bad", result.modifier < 0);
    node.classList.toggle("good", result.modifier > 0);
  }
}
function refreshTargetDistance() {
  if (!targetDistance) return;
  if (targetPickState.active) {
    const source2 = getCharacterById(targetPickState.attackerTokenId);
    if (!source2 || !canUseToken(source2) || source2.id !== activeTokenId) void stopTargetPick();
    return;
  }
  const source = getCharacterById(activeTokenId), targetId = source ? getOdysseyData(source).attackDraft.targetTokenId : "";
  const target = getCharacterById(targetId);
  if (source && target && target.visible !== false && canUseToken(source)) void targetDistance.measure(source, target);
  else paintTargetDistance(null);
}
document.documentElement.classList.toggle("od-hud-mode", HUD_MODE);
var CORE_COMBAT_SKILLS = Object.keys(DEFAULT_ODYSSEY_SKILLS);
var ATTACK_ONLY_EXCLUDED_SKILLS = /* @__PURE__ */ new Set([PARRY_SKILL_NAME]);
var UNARMED_WEAPON_NAME = "Not Armed";
var TOKEN_EXPORT_FORMAT = "odyssey-token-profile";
var TOKEN_EXPORT_VERSION = 1;
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
var characterTools = null;
var profileRollPending = false;
var playerRole = "PLAYER";
var playerId = "";
var playerName = "";
var playerColor = "#facc15";
var sceneItems = [];
var localSceneItems = [];
var selectionIds = [];
var activeTokenId = null;
var debugEntries = [];
var partyPlayers = [];
var sortedPartyPlayersCache = [];
var characterListCache = [];
var trackedCharacterListCache = [];
var charactersByIdCache = /* @__PURE__ */ new Map();
var gmPrivateEntries = [];
var ammunitionAddOpen = false;
var weaponAddOpen = false;
var bodyPartAddOpen = false;
var skillAddOpen = false;
var skillPerkAddSkill = "";
var pendingLocalDebugEntryIds = /* @__PURE__ */ new Set();
var pendingLocalDebugClear = false;
var collapsibleSectionState = /* @__PURE__ */ new Map();
var attackFormDrafts = /* @__PURE__ */ new Map();
var inputAutosaveTimers = /* @__PURE__ */ new Map();
var selectionPollTimer = null;
var renderScheduled = false;
var sceneSideEffectsTimer = null;
var targetPickState = {
  active: false,
  attackerTokenId: null,
  previousToolId: "",
  previousModeId: void 0,
  toolReady: false,
  restoring: false
};
function rebuildSceneCaches() {
  characterListCache = sortCharacters(sceneItems.filter(isCharacterToken));
  trackedCharacterListCache = characterListCache.filter(isTrackedCharacter);
  charactersByIdCache = new Map(characterListCache.map((item) => [item.id, item]));
}
function setSceneItems(items) {
  sceneItems = Array.isArray(items) ? items : [];
  rebuildSceneCaches();
}
function setLocalSceneItems(items) {
  localSceneItems = Array.isArray(items) ? items : [];
}
function setSelectionState(selection) {
  selectionIds = selection ?? [];
}
function setPartyPlayersState(players) {
  partyPlayers = players ?? [];
  sortedPartyPlayersCache = [...partyPlayers].sort(
    (left, right) => String(left?.name ?? "").localeCompare(String(right?.name ?? ""))
  );
}
function removeLocalItemsFromCache(ids) {
  if (!ids.length) return;
  const removeSet = new Set(ids);
  setLocalSceneItems(localSceneItems.filter((item) => !removeSet.has(item.id)));
}
function upsertLocalItemsInCache(items) {
  if (!items.length) return;
  const byId = new Map(localSceneItems.map((item) => [item.id, item]));
  for (const item of items) {
    byId.set(item.id, item);
  }
  setLocalSceneItems([...byId.values()]);
}
function scheduleRender() {
  if (renderScheduled) return;
  renderScheduled = true;
  window.requestAnimationFrame(() => {
    renderScheduled = false;
    render();
  });
}
function scheduleSceneSideEffects(items = sceneItems) {
  if (sceneSideEffectsTimer) {
    clearTimeout(sceneSideEffectsTimer);
  }
  sceneSideEffectsTimer = window.setTimeout(() => {
    sceneSideEffectsTimer = null;
    void syncLocalOwnedHiddenTokenViews(items).catch((error) => {
      console.warn("[Body HP] Unable to sync local hidden-token views", error);
    });
    void syncTargetHighlight().catch((error) => {
      console.warn("[Body HP] Unable to sync target highlight", error);
    });
  }, 60);
}
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
  const category = odyssey?.skillCategories?.[skillName];
  if (category === COMBAT_SKILL_CATEGORY) return COMBAT_SKILL_CATEGORY;
  if (category === ABILITIES_SKILL_CATEGORY) return ABILITIES_SKILL_CATEGORY;
  return APPLIED_SKILL_CATEGORY;
}
function getSkillStrengthBonusFlag(odyssey, skillName) {
  return Boolean(odyssey?.skillStrengthBonuses?.[skillName]);
}
function getSkillLabel(odyssey, skillName) {
  return String(odyssey?.skillLabels?.[skillName] ?? skillName).trim() || skillName;
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
function getAbilitiesSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === ABILITIES_SKILL_CATEGORY
  );
}
function getAttackSkillEntries(odyssey) {
  return [...getCombatSkillEntries(odyssey), ...getAbilitiesSkillEntries(odyssey)].filter(
    ([skillName]) => !ATTACK_ONLY_EXCLUDED_SKILLS.has(skillName)
  );
}
function getAppliedSkillEntries(odyssey) {
  return getSortedSkillEntries(odyssey).filter(
    ([skillName]) => getSkillCategory(odyssey, skillName) === APPLIED_SKILL_CATEGORY
  );
}
function buildSkillOptions(skillEntries, selectedValue = "", odyssey = null) {
  return skillEntries.map(
    ([key, value]) => `<option value="${escapeHtml(key)}" ${key === selectedValue ? "selected" : ""}>${escapeHtml(getSkillLabel(odyssey, key))} (${value})</option>`
  ).join("");
}
function buildGroupedSkillOptions(odyssey, selectedValue = "") {
  const combatOptions = buildSkillOptions(getCombatSkillEntries(odyssey), selectedValue, odyssey);
  const abilitiesOptions = buildSkillOptions(getAbilitiesSkillEntries(odyssey), selectedValue, odyssey);
  const appliedOptions = buildSkillOptions(getAppliedSkillEntries(odyssey), selectedValue, odyssey);
  return [
    combatOptions ? `<optgroup label="Combat">${combatOptions}</optgroup>` : "",
    abilitiesOptions ? `<optgroup label="Abilities">${abilitiesOptions}</optgroup>` : "",
    appliedOptions ? `<optgroup label="Applied">${appliedOptions}</optgroup>` : ""
  ].filter(Boolean).join("");
}
function getTransientFieldKey(field) {
  if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) {
    return "";
  }
  if (field.dataset.attackField === "ammoId" || field.dataset.manualAttackField === "ammoId") return "";
  if (field.dataset.attackField) return `attack:${field.dataset.attackField}`;
  if (field.dataset.manualAttackField) return `manual-attack:${field.dataset.manualAttackField}`;
  if (field.dataset.weaponMag !== void 0) return `inventory:weapon-mag:${field.dataset.weaponMag}`;
  if (field.dataset.magAmmo) return `inventory:mag-ammo:${field.dataset.magAmmo}`;
  if (field.dataset.equipPart) return `inventory:equip-part:${field.dataset.equipPart}`;
  if (field.dataset.profileModifier) return `profile:${field.dataset.profileModifier}`;
  if (field.dataset.rollField) return `roll:${field.dataset.rollField}`;
  if (field.dataset.rollCharField) return `roll-char:${field.dataset.rollCharField}`;
  if (field.dataset.rollSkillField) return `roll-skill:${field.dataset.rollSkillField}`;
  if (field.dataset.gmRollField) return `gm-roll:${field.dataset.gmRollField}`;
  if (field.dataset.skillField) return `new-skill:${field.dataset.skillField}`;
  if (field.dataset.weaponField) return `new-weapon:${field.dataset.weaponField}`;
  if (field.dataset.action === "select-owner-player") return "owner";
  if (field.dataset.action === "set-odyssey-skill") return `skill:${field.dataset.skill ?? ""}`;
  if (field.dataset.action === "set-skill-strength-bonus") {
    return `skill-strength:${field.dataset.skill ?? ""}`;
  }
  if (field.dataset.action === "set-odyssey-attribute") {
    return `attribute:${field.dataset.attribute ?? ""}`;
  }
  if (field.dataset.action === "set-weapon-name") {
    return `weapon-name:${field.dataset.weaponIndex ?? ""}`;
  }
  if (field.dataset.action === "set-weapon-damage") {
    return `weapon-damage:${field.dataset.weaponIndex ?? ""}`;
  }
  if (field.dataset.action === "set-weapon-accuracy") {
    return `weapon-accuracy:${field.dataset.weaponIndex ?? ""}`;
  }
  if (field.dataset.action === "set-field") {
    return `part:${field.dataset.part ?? ""}:${field.dataset.field ?? ""}`;
  }
  return "";
}
function shouldPreserveFieldValue(fieldKey, focusedKey) {
  return fieldKey === focusedKey || fieldKey.startsWith("attack:") || fieldKey.startsWith("manual-attack:") || fieldKey.startsWith("inventory:") || fieldKey.startsWith("profile:") || fieldKey.startsWith("roll:") || fieldKey.startsWith("roll-char:") || fieldKey.startsWith("roll-skill:") || fieldKey.startsWith("gm-roll:") || fieldKey.startsWith("new-skill:") || fieldKey.startsWith("new-weapon:");
}
function captureSelectedPanelState() {
  ui.selectedTokenPanel.querySelectorAll("details[data-section-key]").forEach((section) => collapsibleSectionState.set(section.dataset.sectionKey, section.open));
  const renderedTokenId = String(ui.selectedTokenPanel.dataset.tokenId ?? "").trim();
  if (!renderedTokenId || !ui.selectedTokenPanel.childElementCount) return null;
  let focusedKey = "";
  let selectionStart = null;
  let selectionEnd = null;
  const activeField = document.activeElement;
  if ((activeField instanceof HTMLInputElement || activeField instanceof HTMLSelectElement) && ui.selectedTokenPanel.contains(activeField)) {
    focusedKey = getTransientFieldKey(activeField);
    if (activeField instanceof HTMLInputElement && typeof activeField.selectionStart === "number" && typeof activeField.selectionEnd === "number") {
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
    tokenId: renderedTokenId,
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
      field.value = field.dataset.attackField === "rounds" || field.dataset.manualAttackField === "rounds" ? String(Math.max(1, Math.min(Number(normalizedValue) || 1, Number(field.max)))) : normalizedValue;
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
  return sortedPartyPlayersCache;
}
function getCharacters() {
  return characterListCache;
}
function getTrackedCharacters() {
  return trackedCharacterListCache;
}
function getControllableCharacters() {
  return getTrackedCharacters().filter(
    (token) => canPlayerControlToken(playerRole, playerId, token)
  );
}
function getCharacterById(tokenId) {
  return charactersByIdCache.get(tokenId) ?? null;
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
  if (playerRole !== "GM") {
    const firstControllable = getControllableCharacters()[0];
    if (firstControllable) return firstControllable.id;
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
function canViewAttackBlock(token) {
  return canUseToken(token);
}
function canViewBodyHpSummary(token) {
  return canUseToken(token);
}
function canViewDiceBlock(token) {
  return canUseToken(token);
}
function canViewOverlayPreview(token) {
  return canUseToken(token);
}
function applyLocalTrackerDataUpdate(tokenId, updater) {
  const nextItems = sceneItems.map((item) => {
    if (item.id !== tokenId) return item;
    const nextItem = structuredClone(item);
    nextItem.metadata ?? (nextItem.metadata = {});
    nextItem.metadata[META_KEY] = updater(getTrackerData(item));
    return nextItem;
  });
  setSceneItems(nextItems);
}
async function updateTrackerData2(tokenId, updater) {
  applyLocalTrackerDataUpdate(tokenId, updater);
  scheduleRender();
  try {
    await updateTrackerData(tokenId, updater);
  } catch (error) {
    setSceneItems(await lib_default.scene.items.getItems());
    scheduleRender();
    throw error;
  }
}
async function initializeCharacterToken(tokenId) {
  const token = getCharacterById(tokenId);
  if (!token || !isCharacterToken(token)) return false;
  const shouldInitialize = !isTrackedCharacter(token);
  if (!shouldInitialize) return false;
  await updateTrackerData2(tokenId, (current2) => current2);
  await ensureOverlayForToken(tokenId);
  return shouldInitialize;
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
  characterTools?.syncPool();
  pendingLocalDebugEntryIds.add(entry.id);
  await lib_default.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-entry", entry },
    { destination: "ALL" }
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
  pendingLocalDebugClear = true;
  await lib_default.broadcast.sendMessage(
    DEBUG_BROADCAST_CHANNEL,
    { type: "debug-clear" },
    { destination: "ALL" }
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
async function handleSelectionChange(selection, allowInitialize = true) {
  const nextSelection = selection ?? [];
  setSelectionState(nextSelection);
  const selectedCharacterId = nextSelection.find((id) => charactersByIdCache.has(id));
  if (selectedCharacterId) {
    activeTokenId = selectedCharacterId;
    if (allowInitialize) {
      await initializeCharacterToken(selectedCharacterId);
    }
  } else if (activeTokenId && !charactersByIdCache.has(activeTokenId)) {
    activeTokenId = null;
  } else {
    activeTokenId = resolveActiveTokenId();
  }
  scheduleRender();
  await syncTargetHighlight();
}
function startSelectionPolling() {
  if (selectionPollTimer) {
    clearInterval(selectionPollTimer);
  }
  selectionPollTimer = setInterval(() => {
    void lib_default.player.getSelection().then((selection) => selection ?? []).then((selection) => {
      if (!arraysEqual(selectionIds, selection)) {
        return handleSelectionChange(selection);
      }
      return null;
    }).catch((error) => {
      console.warn("[Body HP] Selection polling failed", error);
    });
  }, 1500);
}
function formatAttackDebug({
  range,
  distance,
  attackerName,
  targetName,
  targetPart,
  attackSkillName,
  attackSkillValue,
  weaponDamage,
  weaponAccuracy,
  strengthBonus,
  manualAttackBonuses,
  totalAttackBonuses,
  manualAttackPenalties,
  automaticTargetPenalty,
  totalAttackPenalties,
  defenseBonuses,
  defensePenalties,
  baseTargetParry,
  targetParry,
  parryMode,
  targetArmor,
  specialArmor,
  specialActive,
  result,
  beforeHp,
  afterHp,
  specialBeforeHp,
  specialAfterHp,
  beforeMinor,
  afterMinor,
  beforeSerious,
  afterSerious,
  critApplied,
  damageAppliedLabel
}) {
  const accuracyTable = formatTextTable(
    ["Side", "Attacking", "Defending"],
    [
      [
        "Accuracy",
        `${result.attackRoll} + ${attackSkillValue * 10} + ${totalAttackBonuses} - ${totalAttackPenalties} = ${result.attackTotal}`,
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
  const damageRows = [
    ["Attacker", attackerName],
    ["Target", `${targetName} -> ${targetPart}`],
    ["Attack Skill", `${attackSkillName} (${attackSkillValue})`],
    ["Strength Bonus", strengthBonus],
    ["Weapon Accuracy", weaponAccuracy],
    ...range?.category ? [["Weapon range", rangeAdjustmentLabel(range)]] : [],
    ["Manual Attack Bonus", manualAttackBonuses],
    ["Total Attack Bonus", totalAttackBonuses],
    ["Manual Attack Penalty", manualAttackPenalties],
    ["Auto Target Penalty", automaticTargetPenalty],
    ["Parry Mode", getParryModeLabel(parryMode)],
    ...distance ? [["Target distance", distance.label]] : [],
    ["Base Parry", baseTargetParry],
    ["Effective Parry", targetParry],
    ["Armor", targetArmor]
  ];
  if (specialActive) {
    damageRows.push(["Special Armor", specialArmor]);
    damageRows.push(["Special HP", formatStateTransition(specialBeforeHp, specialAfterHp)]);
  }
  damageRows.push(
    ["Outcome", formatAttackOutcomeLabel(result.outcome)],
    ["Applied Damage", damageAppliedLabel ?? formatAppliedDamageLabel(result.damage, critApplied)],
    ["Damage Diff", result.damage?.damageDiff ?? 0],
    ["Damage Label", result.damage?.label ?? "No damage"],
    ["Applied Min/Sir/Crit", `${result.damage?.minor ?? 0} / ${result.damage?.serious ?? 0} / ${result.damage?.crit ?? 0}`],
    ["Converted Crit", critApplied]
  );
  const damageTable = formatTextTable(
    ["Parameter", "Value"],
    damageRows
  );
  return [
    `Damage Applied: ${damageAppliedLabel ?? formatAppliedDamageLabel(result.damage, critApplied)}`,
    `Result: ${formatAttackOutcomeLabel(result.outcome)}`,
    "",
    accuracyTable,
    "",
    damageTable
  ].join("\n");
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
function getNormalizedPartState(part) {
  return {
    current: Number(part?.current) || 0,
    max: Number(part?.max) || 0,
    armor: Number(part?.armor) || 0,
    minor: Number(part?.minor) || 0,
    serious: Number(part?.serious) || 0
  };
}
function projectDamageWithSpecialProtection({
  specialPart,
  targetPart,
  damage,
  targetPartName
}) {
  const normalizedSpecial = getNormalizedPartState(specialPart);
  const normalizedTarget = getNormalizedPartState(targetPart);
  if (!damage) {
    return {
      specialProjectedState: normalizedSpecial,
      projectedTargetState: {
        ...normalizedTarget,
        critApplied: 0
      },
      specialActive: false,
      specialArmor: 0,
      damageAppliedLabel: "No Damage"
    };
  }
  const specialActive = hasConfiguredSpecial({ body: { [SPECIAL_PART_NAME]: normalizedSpecial } }) && normalizedSpecial.max > 0 && normalizedSpecial.current > 0;
  if (!specialActive) {
    const projectedTargetState2 = projectPartDamage(normalizedTarget, damage);
    return {
      specialProjectedState: normalizedSpecial,
      projectedTargetState: projectedTargetState2,
      specialActive: false,
      specialArmor: 0,
      damageAppliedLabel: formatAppliedDamageLabel(damage, projectedTargetState2.critApplied ?? 0)
    };
  }
  const specialProjectedState = projectPartDamage(normalizedSpecial, damage);
  const absorbedHp = Math.max(0, normalizedSpecial.current - specialProjectedState.current);
  const totalSpecialCrit = Math.max(0, Number(specialProjectedState.critApplied) || 0);
  const specialCritApplied = Math.min(totalSpecialCrit, absorbedHp);
  const overflowCrit = Math.max(0, totalSpecialCrit - specialCritApplied);
  const overflowDamage = overflowCrit > 0 ? { crit: overflowCrit, serious: 0, minor: 0 } : null;
  const projectedTargetState = overflowDamage ? projectPartDamage(normalizedTarget, overflowDamage) : {
    ...normalizedTarget,
    critApplied: 0
  };
  const specialAppliedLabel = formatAppliedDamageLabel(
    {
      serious: damage.serious,
      minor: damage.minor
    },
    specialCritApplied
  );
  const targetAppliedLabel = overflowDamage ? formatAppliedDamageLabel(overflowDamage, projectedTargetState.critApplied ?? 0) : "No Damage";
  let damageAppliedLabel = "No Damage";
  if (specialAppliedLabel !== "No Damage") {
    damageAppliedLabel = `Special ${specialAppliedLabel}`;
  }
  if (targetAppliedLabel !== "No Damage") {
    damageAppliedLabel = damageAppliedLabel === "No Damage" ? `${targetPartName} ${targetAppliedLabel}` : `${damageAppliedLabel}; ${targetPartName} ${targetAppliedLabel}`;
  }
  return {
    specialProjectedState,
    projectedTargetState,
    specialActive: true,
    specialArmor: normalizedSpecial.armor,
    damageAppliedLabel
  };
}
function formatRawDiceRolls(result) {
  return result.rolls.join(", ");
}
function formatDiceRollsWithModifier(result) {
  const modifier2 = Number(result.modifier) || 0;
  return result.rolls.map((roll) => (Number(roll) || 0) + modifier2).join(", ");
}
function buildDiceRollSummary(diceLabel, result) {
  return `Rolled ${diceLabel}: raw [${formatRawDiceRolls(result)}], sum ${result.subtotal}, with modifier ${formatDiceRollsWithModifier(result)}`;
}
function formatOverlayPreviewText(data) {
  return `Armor Total ${getArmorTotal(data)}
${formatOverlayText(data)}`;
}
function formatDiceDebug({ tokenName, result }) {
  return formatTextTable(
    ["Parameter", "Value"],
    [
      ["Actor", tokenName],
      ["Dice", `${result.count}d${result.sides}`],
      ["Raw Dice", formatRawDiceRolls(result)],
      ["Dice Sum", result.subtotal],
      ["With Modifier", formatDiceRollsWithModifier(result)]
    ]
  );
}
function getCurrentPlayerColor() {
  return String(
    playerColor || partyPlayers.find((player) => player?.id === playerId)?.color || "#facc15"
  );
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
function formatAppliedDamageLabel(damage, critApplied = 0) {
  const parts = [];
  const totalCrit = Math.max(0, Number(critApplied) || 0);
  const serious = Math.max(0, Number(damage?.serious) || 0);
  const minor = Math.max(0, Number(damage?.minor) || 0);
  if (totalCrit > 0) {
    parts.push(`${totalCrit} Crit`);
  }
  if (serious > 0) {
    parts.push(`${serious} Serious`);
  }
  if (minor > 0) {
    parts.push(`${minor} Minor`);
  }
  return parts.length ? parts.join(", ") : "No Damage";
}
function formatStateTransition(before, after) {
  if (before == null || after == null) return "-";
  return `${before} -> ${after}`;
}
function getResolvedCheckResultIcon(resultLabel) {
  const normalized = String(resultLabel).trim();
  if (normalized === "Critical Success") return "\u{1F3AF}";
  if (normalized === "Critical Failure") return "\u{1F480}";
  if (normalized === "Check Passed") return "\u2705";
  return "\u274C";
}
function isResolvedCheckResultSuccess(resultLabel) {
  const normalized = String(resultLabel).trim();
  return normalized === "Check Passed" || normalized === "Critical Success";
}
function formatResolvedCheckResultLabel(resultLabel) {
  const normalized = String(resultLabel).trim() || "Check Failed";
  return `${getResolvedCheckResultIcon(normalized)} ${normalized}`;
}
function formatRollCharDebug({ tokenName, attributeLabel, result }) {
  return [
    `Character: ${tokenName}`,
    `Characteristic: ${attributeLabel}`,
    `${formatResolvedCheckResultLabel(result.result)}`,
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
    `${formatResolvedCheckResultLabel(result.result)}`,
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
function formatTextTable(headers, rows2) {
  const normalizedHeaders = headers.map((cell) => String(cell ?? ""));
  const normalizedRows = rows2.map((row) => row.map((cell) => String(cell ?? "")));
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
  const availableWeapons = getAttackSelectableWeapons(token);
  const defaultWeapon = getDefaultAttackWeapon(token);
  const stored = attackFormDrafts.get(token.id) ?? {};
  const persistedTargetTokenId = String(data.odyssey?.attackDraft?.targetTokenId ?? "").trim();
  const persistedTargetTokenName = String(data.odyssey?.attackDraft?.targetTokenName ?? "").trim();
  const combatSkillNames = getAttackSkillEntries(data.odyssey).map(([skillName]) => skillName);
  const fallbackSkill = combatSkillNames[0] ?? CORE_COMBAT_SKILLS.find((key) => key in data.odyssey.skills) ?? CORE_COMBAT_SKILLS[0];
  const hasStoredTargetTokenId = Object.prototype.hasOwnProperty.call(stored, "targetTokenId");
  const draftTargetTokenId = persistedTargetTokenId || (hasStoredTargetTokenId ? String(stored.targetTokenId ?? "").trim() : null);
  const draftTargetTokenName = persistedTargetTokenName || String(stored.targetTokenName ?? "").trim();
  const targetTokenId = draftTargetTokenId && draftTargetTokenId !== token.id ? draftTargetTokenId : "";
  const resolvedTarget = targetTokenId ? targetCharacters.find((target) => target.id === targetTokenId) ?? null : null;
  const storedWeaponName = String(stored.weaponName ?? "").trim();
  const selectedWeapon = availableWeapons.find((weapon) => (weapon.selectionKey ?? weapon.name) === storedWeaponName) ?? defaultWeapon;
  return {
    skill: combatSkillNames.includes(stored.skill) ? stored.skill : fallbackSkill,
    targetTokenId,
    targetTokenName: resolvedTarget ? getCharacterName(resolvedTarget) : draftTargetTokenName,
    targetPart: String(stored.targetPart ?? "").trim() && stored.targetPart !== SPECIAL_PART_NAME ? String(stored.targetPart).trim() : "Torso",
    weaponName: selectedWeapon?.selectionKey ?? selectedWeapon?.name ?? defaultWeapon.name,
    weaponDamage: stored.weaponDamage ?? String(selectedWeapon?.damage ?? defaultWeapon.damage ?? 0),
    weaponAccuracy: String(selectedWeapon?.accuracy ?? defaultWeapon.accuracy ?? 0),
    attackBonuses: stored.attackBonuses ?? "0",
    attackPenalties: stored.attackPenalties ?? "0",
    manualAttackBonuses: stored.manualAttackBonuses ?? stored.attackBonuses ?? "0",
    manualAttackPenalties: stored.manualAttackPenalties ?? stored.attackPenalties ?? "0",
    defenseBonuses: stored.defenseBonuses ?? "0",
    defensePenalties: stored.defensePenalties ?? "0",
    distanceMeters: stored.distanceMeters ?? "",
    manualArmor: stored.manualArmor ?? "0",
    manualParry: stored.manualParry ?? "0",
    rounds: stored.rounds ?? "1",
    parryMode: ["off", "1", "2", "3", "4", "5"].includes(String(stored.parryMode)) ? String(stored.parryMode) : getOdysseyData(token).preferences?.parryMode || "off"
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
function getSharedAttackDraftField(manualField) {
  if (manualField === "distanceMeters") return "distanceMeters";
  if (manualField === "rounds") return "rounds";
  if (manualField === "targetPart") return "targetPart";
  if (manualField === "skill") return "skill";
  if (manualField === "weaponName") return "weaponName";
  if (manualField === "weaponDamage") return "weaponDamage";
  if (manualField === "attackBonuses") return "manualAttackBonuses";
  if (manualField === "attackPenalties") return "manualAttackPenalties";
  if (manualField === "manualArmor") return "manualArmor";
  if (manualField === "manualParry") return "manualParry";
  return "";
}
function buildWeaponOptions(weapons, selectedWeaponName = "") {
  return weapons.map(
    (weapon) => `<option value="${escapeHtml(weapon.selectionKey ?? weapon.name)}" ${(weapon.selectionKey ?? weapon.name) === selectedWeaponName ? "selected" : ""}>${escapeHtml(weapon.name)} (${weapon.rangedIndex == null ? `DMG ${weapon.damage}, ` : `Ammo ${weapon.loaded}/${weapon.capacity}, `}ACC ${weapon.accuracy >= 0 ? "+" : ""}${weapon.accuracy})</option>`
  ).join("");
}
function getAttackSelectableWeapons(token) {
  const odyssey = getOdysseyData(token);
  const meleeWeapons = odyssey.weapons?.melee ?? [];
  return [
    { name: UNARMED_WEAPON_NAME, damage: 0, accuracy: 0 },
    ...meleeWeapons,
    ...(odyssey.weapons?.ranged ?? []).map((weapon, rangedIndex) => ({ ...weapon, damage: 0, rangedIndex, selectionKey: `ranged:${weapon.id}` }))
  ];
}
function getDefaultAttackWeapon(token) {
  const odyssey = getOdysseyData(token);
  const meleeWeapons = odyssey.weapons?.melee ?? [];
  return meleeWeapons[0] ?? { name: UNARMED_WEAPON_NAME, damage: 0, accuracy: 0 };
}
function getWeaponByName(token, weaponName) {
  const normalizedWeaponName = String(weaponName ?? "").trim();
  const weapons = getAttackSelectableWeapons(token);
  return weapons.find((weapon) => (weapon.selectionKey ?? weapon.name) === normalizedWeaponName) ?? getDefaultAttackWeapon(token);
}
function syncAttackWeaponInputs(tokenId, weaponName, weaponDamage, weaponAccuracy = 0) {
  scheduleRender();
  if (!tokenId) return;
  saveAttackDraftValue(tokenId, "weaponName", weaponName);
  saveAttackDraftValue(tokenId, "weaponDamage", String(weaponDamage));
  saveAttackDraftValue(tokenId, "weaponAccuracy", String(weaponAccuracy));
  const attackWeaponSelect = ui.selectedTokenPanel.querySelector('[data-attack-field="weaponName"]');
  if (attackWeaponSelect instanceof HTMLSelectElement) {
    const hasOption = Array.from(attackWeaponSelect.options).some(
      (option) => option.value === weaponName
    );
    if (hasOption) {
      attackWeaponSelect.value = weaponName;
    }
  }
  const manualWeaponSelect = ui.selectedTokenPanel.querySelector('[data-manual-attack-field="weaponName"]');
  if (manualWeaponSelect instanceof HTMLSelectElement) {
    const hasOption = Array.from(manualWeaponSelect.options).some(
      (option) => option.value === weaponName
    );
    if (hasOption) {
      manualWeaponSelect.value = weaponName;
    }
  }
  ui.selectedTokenPanel.querySelectorAll('[data-attack-field="weaponDamage"], [data-manual-attack-field="weaponDamage"]').forEach((field) => {
    if (field instanceof HTMLInputElement) {
      field.value = String(weaponDamage);
    }
  });
  ui.selectedTokenPanel.querySelectorAll("[data-weapon-accuracy-display]").forEach((field) => {
    field.textContent = `${weaponAccuracy >= 0 ? "+" : ""}${weaponAccuracy}`;
  });
}
function sanitizeFileName(value) {
  return String(value ?? "").trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-").replace(/\s+/g, "_").slice(0, 80) || "token";
}
function buildTokenExportPayload(token) {
  return { format: TOKEN_EXPORT_FORMAT, version: TOKEN_EXPORT_VERSION, exportedAt: (/* @__PURE__ */ new Date()).toISOString(), tokenName: getCharacterName(token), data: profileData(getTrackerData(token)) };
}
function normalizeImportedTokenPayload(raw) {
  if (!raw || typeof raw !== "object") throw new Error("Invalid token profile.");
  const source = raw.format === TOKEN_EXPORT_FORMAT && raw.data ? raw.data : raw.trackerData ?? raw;
  return profileData(sanitizeTrackerData(source));
}
function downloadTextFile(filename, content, mimeType = "application/json") {
  const blob = new Blob([content], { type: `${mimeType};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 0);
}
async function exportSelectedTokenData() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can export token data.", "error");
    return;
  }
  const payload = buildTokenExportPayload(token);
  const timestamp = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  const fileName = `${sanitizeFileName(getCharacterName(token))}-${timestamp}.odyssey-token.json`;
  downloadTextFile(fileName, JSON.stringify(payload, null, 2));
  setStatus(`Exported ${getCharacterName(token)} data to ${fileName}.`, "success");
}
function openTokenImportPicker() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can import token data.", "error");
    return;
  }
  const input = ui.selectedTokenPanel.querySelector("[data-token-import-input]");
  if (!(input instanceof HTMLInputElement)) {
    setStatus("Import control is not available for this token.", "error");
    return;
  }
  input.value = "";
  input.click();
}
async function importSelectedTokenData(file) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can import token data.", "error");
    return;
  }
  if (!(file instanceof File)) {
    setStatus("Choose an import file first.", "error");
    return;
  }
  const text2 = await file.text();
  let parsed;
  try {
    parsed = JSON.parse(text2);
  } catch (error) {
    throw new Error("Import file must contain valid JSON.");
  }
  const imported = normalizeImportedTokenPayload(parsed);
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    next.enabled = imported.enabled;
    next.minor = imported.minor;
    next.serious = imported.serious;
    next.body = structuredClone(imported.body);
    next.odyssey.skills = structuredClone(imported.odyssey.skills);
    next.odyssey.skillCategories = structuredClone(imported.odyssey.skillCategories);
    next.odyssey.skillStrengthBonuses = structuredClone(imported.odyssey.skillStrengthBonuses);
    next.odyssey.skillLabels = structuredClone(imported.odyssey.skillLabels);
    next.odyssey.skillPerks = structuredClone(imported.odyssey.skillPerks);
    next.odyssey.attributes = structuredClone(imported.odyssey.attributes);
    next.odyssey.weapons = structuredClone(imported.odyssey.weapons);
    next.odyssey.ammunition = structuredClone(imported.odyssey.ammunition);
    next.odyssey.inventory = structuredClone(imported.odyssey.inventory);
    return next;
  });
  await ensureOverlayForToken(token.id);
  const fallbackWeapon = getDefaultAttackWeapon(getCharacterById(token.id) ?? token);
  syncAttackWeaponInputs(
    token.id,
    fallbackWeapon.name,
    fallbackWeapon.damage,
    fallbackWeapon.accuracy ?? 0
  );
  setStatus(`Imported token data into ${getCharacterName(token)}.`, "success");
}
async function persistAttackTargetToken(tokenId, targetTokenId) {
  if (!tokenId) return;
  const token = getCharacterById(tokenId);
  if (!token) return;
  const normalizedTargetTokenId = String(targetTokenId ?? "").trim();
  const targetToken = normalizedTargetTokenId ? getCharacterById(normalizedTargetTokenId) : null;
  const normalizedTargetTokenName = targetToken ? getCharacterName(targetToken) : "";
  const currentAttackDraft = getOdysseyData(token).attackDraft ?? {};
  const currentTargetTokenId = String(currentAttackDraft.targetTokenId ?? "").trim();
  const currentTargetTokenName = String(currentAttackDraft.targetTokenName ?? "").trim();
  if (currentTargetTokenId === normalizedTargetTokenId && currentTargetTokenName === normalizedTargetTokenName) {
    return;
  }
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    next.odyssey ?? (next.odyssey = structuredClone(getTrackerData(token).odyssey));
    (_a = next.odyssey).attackDraft ?? (_a.attackDraft = { targetTokenId: "", targetTokenName: "" });
    next.odyssey.attackDraft.targetTokenId = normalizedTargetTokenId;
    next.odyssey.attackDraft.targetTokenName = normalizedTargetTokenName;
    return next;
  });
}
async function assignAttackTarget(attacker, target, successMessage) {
  if (!attacker || !isCharacterToken(attacker)) {
    setStatus("Select an attacker token first.", "error");
    return false;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return false;
  }
  if (!target || !isCharacterToken(target)) {
    setStatus("Choose a valid target token.", "error");
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
  saveAttackDraftValue(attacker.id, "targetTokenName", getCharacterName(target));
  await persistAttackTargetToken(attacker.id, target.id);
  activeTokenId = attacker.id;
  scheduleRender();
  const targetField = ui.selectedTokenPanel.querySelector('[data-attack-field="targetTokenId"]');
  if (targetField instanceof HTMLInputElement || targetField instanceof HTMLSelectElement) {
    targetField.value = target.id;
  }
  await syncTargetHighlight();
  setStatus(successMessage || `Target set to ${getCharacterName(target)}.`, "success");
  return true;
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
  const highlightIds = localSceneItems.filter(isLocalTargetHighlight).map((item) => item.id);
  if (highlightIds.length) {
    await lib_default.scene.local.deleteItems(highlightIds);
    removeLocalItemsFromCache(highlightIds);
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
  if (!attacker || !isCharacterToken(attacker) || !canViewAttackBlock(attacker)) {
    await clearTargetHighlight();
    return;
  }
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== attacker.id && item.visible !== false
  );
  const draft2 = getAttackDraft(attacker, getTrackerData(attacker), targetCharacters);
  const target = getCharacterById(draft2.targetTokenId);
  const existingHighlights = localSceneItems.filter(isLocalTargetHighlight);
  if (!target || !isCharacterToken(target) || target.visible === false || target.id === attacker.id) {
    if (existingHighlights.length) {
      const highlightIds = existingHighlights.map((item) => item.id);
      await lib_default.scene.local.deleteItems(highlightIds);
      removeLocalItemsFromCache(highlightIds);
    }
    return;
  }
  if (existingHighlights.length === 1 && existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.targetTokenId === target.id && existingHighlights[0]?.metadata?.[TARGET_HIGHLIGHT_KEY]?.attackerTokenId === attacker.id) {
    return;
  }
  if (existingHighlights.length) {
    const highlightIds = existingHighlights.map((item) => item.id);
    await lib_default.scene.local.deleteItems(highlightIds);
    removeLocalItemsFromCache(highlightIds);
  }
  const nextHighlight = await buildTargetHighlightItem(target);
  await lib_default.scene.local.addItems([nextHighlight]);
  upsertLocalItemsInCache([nextHighlight]);
}
function isLocalSelfViewItem(item) {
  return Boolean(item?.metadata?.[LOCAL_SELF_VIEW_KEY]);
}
function buildLocalSelfViewSignature(token) {
  return [
    token.lastModified,
    token.position?.x ?? 0,
    token.position?.y ?? 0,
    token.rotation ?? 0,
    token.scale?.x ?? 1,
    token.scale?.y ?? 1,
    token.zIndex ?? 0
  ].join("|");
}
async function buildLocalSelfViewItem(token) {
  return buildImage(token.image, token.grid).name(`${getCharacterName(token)} (Local View)`).position(token.position).rotation(token.rotation ?? 0).scale(token.scale ?? { x: 1, y: 1 }).zIndex((token.zIndex ?? 0) + 1).visible(true).layer("CHARACTER").locked(true).disableHit(true).metadata({
    [LOCAL_SELF_VIEW_KEY]: {
      tokenId: token.id,
      ownerPlayerId: playerId,
      signature: buildLocalSelfViewSignature(token)
    }
  }).build();
}
async function clearLocalSelfViews() {
  const localViewIds = localSceneItems.filter(isLocalSelfViewItem).map((item) => item.id);
  if (localViewIds.length) {
    await lib_default.scene.local.deleteItems(localViewIds);
    removeLocalItemsFromCache(localViewIds);
  }
}
async function syncLocalOwnedHiddenTokenViews(items = sceneItems) {
  if (playerRole === "GM" || !playerId) {
    await clearLocalSelfViews();
    return;
  }
  const desiredTokens = items.filter(
    (item) => isCharacterToken(item) && item.visible === false && canUseToken(item)
  );
  const desiredTokenIds = new Set(desiredTokens.map((token) => token.id));
  const existingViews = localSceneItems.filter(isLocalSelfViewItem);
  const staleViewIds = existingViews.filter((item) => !desiredTokenIds.has(String(item.metadata?.[LOCAL_SELF_VIEW_KEY]?.tokenId ?? ""))).map((item) => item.id);
  if (staleViewIds.length) {
    await lib_default.scene.local.deleteItems(staleViewIds);
    removeLocalItemsFromCache(staleViewIds);
  }
  const currentViews = localSceneItems.filter(isLocalSelfViewItem);
  const viewIdsToReplace = [];
  const itemsToAdd = [];
  for (const token of desiredTokens) {
    const existingView = currentViews.find(
      (item) => item.metadata?.[LOCAL_SELF_VIEW_KEY]?.tokenId === token.id
    );
    const nextSignature = buildLocalSelfViewSignature(token);
    if (existingView?.metadata?.[LOCAL_SELF_VIEW_KEY]?.signature === nextSignature) {
      continue;
    }
    if (existingView) {
      viewIdsToReplace.push(existingView.id);
    }
    itemsToAdd.push(await buildLocalSelfViewItem(token));
  }
  if (viewIdsToReplace.length) {
    await lib_default.scene.local.deleteItems(viewIdsToReplace);
    removeLocalItemsFromCache(viewIdsToReplace);
  }
  if (itemsToAdd.length) {
    await lib_default.scene.local.addItems(itemsToAdd);
    upsertLocalItemsInCache(itemsToAdd);
  }
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
  await targetDistance?.clear();
  scheduleRender();
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
    onToolMove: (_context, event) => {
      if (!targetPickState.active || !targetDistance) return;
      const source = getCharacterById(targetPickState.attackerTokenId);
      if (!source || !canUseToken(source)) return;
      const hovered = event.target && isCharacterToken(event.target) && event.target.visible !== false ? getCharacterById(event.target.id) || event.target : null;
      const target = hovered || event.pointerPosition;
      if (target) targetDistance.preview(source, target);
    },
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
      const assigned = await assignAttackTarget(attacker, target);
      if (!assigned) return false;
      await stopTargetPick();
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
  measuredTargetDistance = null;
  await ensureTargetPickerTool();
  await lib_default.tool.activateTool(TARGET_PICK_TOOL_ID);
  await lib_default.tool.activateMode(TARGET_PICK_TOOL_ID, TARGET_PICK_MODE_ID);
  scheduleRender();
  setStatus("Click a visible target token on the map to assign it.", "info");
}
async function removeAttackTargetContextMenu() {
  try {
    await lib_default.contextMenu.remove(ATTACK_TARGET_CONTEXT_MENU_ID);
  } catch (_error) {
  }
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
function rollCharacterCheck(attributeValue, modifier2 = 0) {
  const baseAttribute = Number(attributeValue) || 0;
  const finalAttribute = Math.max(0, baseAttribute + (Number(modifier2) || 0));
  const roll = Math.floor(Math.random() * 20) + 1;
  let result = roll <= finalAttribute ? "Check Passed" : "Check Failed";
  let outcome = result === "Check Passed" ? "success" : "failure";
  if (roll === 1) {
    result = "Critical Success";
    outcome = "critical-success";
  } else if (roll === 20) {
    result = "Critical Failure";
    outcome = "critical-failure";
  }
  return {
    roll,
    baseAttribute,
    modifier: Number(modifier2) || 0,
    finalAttribute,
    result,
    outcome
  };
}
function rollSkillCheck(skillValue, modifier2 = 0) {
  const baseSkill = Number(skillValue) || 0;
  const rollPrimary = Math.floor(Math.random() * 100) + 1;
  const rollSecondary = Math.floor(Math.random() * 100) + 1;
  const totalPrimary = rollPrimary + baseSkill * 10 + (Number(modifier2) || 0);
  const totalSecondary = rollSecondary;
  let result = totalPrimary > totalSecondary ? "Check Passed" : "Check Failed";
  let outcome = result === "Check Passed" ? "success" : "failure";
  if (rollPrimary >= 95) {
    result = "Critical Success";
    outcome = "critical-success";
  } else if (rollPrimary <= 5) {
    result = "Critical Failure";
    outcome = "critical-failure";
  }
  return {
    rollPrimary,
    rollSecondary,
    baseSkill,
    modifier: Number(modifier2) || 0,
    totalPrimary,
    totalSecondary,
    result,
    outcome
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
      <div class="ownership-transfer">
        <input type="file" accept=".json,.txt,.odyssey-token.json" data-token-import-input hidden>
        ${isEditable() && characterTools ? characterTools.renderTemplates() : ""}
      </div>
    `,
    false
  );
}
function renderCharacteristicsBlock(data) {
  return renderCollapsibleSection("Characteristics", attributeCards(data.odyssey, ATTRIBUTE_UI_FIELDS, isEditable()), true);
}
function renderEnglishSkillsBlock(data) {
  const groups = [["Combat", getCombatSkillEntries(data.odyssey)], ["Abilities", getAbilitiesSkillEntries(data.odyssey)], ["Applied", getAppliedSkillEntries(data.odyssey)]];
  return renderCollapsibleSection("Skills", skillCards(data.odyssey, groups.map(([label, entries]) => [label, entries.filter(([key]) => !Object.hasOwn(ARMOR_TRAINING, key))]), isEditable()), true);
}
function toggleWeaponDraftFields(select) {
  const row = select.closest(".weapon-entry");
  if (!row) return;
  const ranged = select.value === "ranged";
  row.querySelectorAll(".weapon-melee-fields").forEach((field) => field.classList.toggle("hidden", ranged));
  row.querySelectorAll(".weapon-ranged-fields").forEach((field) => field.classList.toggle("hidden", !ranged));
}
async function saveWeapon(node) {
  if (!isEditable()) {
    throw new Error("Only the GM can edit weapons.");
  }
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a token first.");
  const row = node.closest(".weapon-entry");
  if (!row) throw new Error("Weapon form is missing.");
  const read = (field) => row.querySelector(`[data-weapon-draft="${field}"]`);
  const name = String(read("name")?.value ?? "").trim();
  const nextType = read("type")?.value === "ranged" ? "ranged" : "melee";
  const type = node.dataset.weaponType;
  const index = Number(node.dataset.weaponIndex);
  const accuracy = clamp(Number(read("accuracy")?.value) || 0, -999, 999);
  const damage = clamp(Number(read("damage")?.value) || 0, -999, 999);
  const capacity = clamp(Number(read("capacity")?.value) || 0, 0, 999);
  const ammoIds = Array.from(row.querySelectorAll('[data-weapon-draft="compatible"]:checked')).map((field) => field.dataset.ammoId);
  if (!name) throw new Error("Enter a weapon name first.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const source = type ? next.odyssey.weapons[type] ?? [] : [];
    const weapon = Number.isInteger(index) ? source[index] : null;
    const allWeapons = [...next.odyssey.weapons.melee ?? [], ...next.odyssey.weapons.ranged ?? []];
    if (allWeapons.some((item) => item !== weapon && item.name === name)) throw new Error("A weapon with this name already exists.");
    if (weapon && type !== nextType) {
      if (type === "ranged" && weapon.magazineId) swapMagazine(next.odyssey, index, "");
      if (type === "ranged" && weapon.loaded > 0) {
        const loadedAmmo = next.odyssey.ammunition.find((ammo) => ammo.id === weapon.loadedAmmoId);
        if (loadedAmmo) loadedAmmo.quantity += weapon.loaded;
      }
      source.splice(index, 1);
    }
    if (nextType === "ranged") {
      const ranged = weapon && type === "ranged" ? weapon : { id: crypto.randomUUID(), loaded: 0, loadedAmmoId: "", ammoIds: [], damage: 0 };
      if (capacity < ranged.loaded) throw new Error("Unload rounds before reducing capacity.");
      if (ranged.loaded > 0 && !ammoIds.includes(ranged.loadedAmmoId)) throw new Error("Loaded ammunition must remain compatible.");
      Object.assign(ranged, { name, accuracy, capacity, ammoIds });
      if (!weapon || type !== "ranged") next.odyssey.weapons.ranged.push(ranged);
    } else {
      const melee = weapon && type === "melee" ? weapon : {};
      Object.assign(melee, { name, damage, accuracy });
      if (!weapon || type !== "melee") next.odyssey.weapons.melee.push(melee);
    }
    return next;
  });
}
async function saveAmmunition(node) {
  if (!isEditable()) throw new Error("Only the GM can edit ammunition.");
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a token first.");
  const row = node.closest(".ammo-entry");
  if (!row) throw new Error("Ammunition form is missing.");
  const read = (field) => row.querySelector(`[data-ammo-draft="${field}"]`);
  const name = String(read("name")?.value ?? "").trim();
  const ammoId = node.dataset.ammoId;
  if (!name) throw new Error("Enter an ammunition name first.");
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    const ammunition = (_a = next.odyssey).ammunition ?? (_a.ammunition = []);
    const ammo = ammunition.find((item) => item.id === ammoId);
    if (ammunition.some((item) => item !== ammo && item.name === name)) throw new Error("Ammunition with this name already exists.");
    const values = {
      name,
      damage: clamp(Number(read("damage")?.value) || 0, 0, 999),
      penetration: clamp(Number(read("penetration")?.value) || 0, -999, 999),
      quantity: clamp(Number(read("quantity")?.value) || 0, 0, 999999)
    };
    if (ammo) Object.assign(ammo, values);
    else ammunition.push({ id: crypto.randomUUID(), ...values });
    return next;
  });
}
function renderAttackAmmunition(token, draft2, manual, disabledAttr) {
  const weapon = getWeaponByName(token, draft2.weaponName);
  if (weapon.rangedIndex == null) return "";
  const ammo = getOdysseyData(token).ammunition ?? [];
  const prefix = manual ? "data-manual-attack-field" : "data-attack-field";
  return `<label class="field-stack"><span class="field-label">Ammunition (${weapon.loaded}/${weapon.capacity})</span>
    <select ${prefix}="ammoId" ${disabledAttr}>
      <option value="" ${!weapon.loadedAmmoId ? "selected" : ""}>Empty</option>
      ${ammo.filter((item) => weapon.ammoIds.includes(item.id)).map((item) => `<option value="${escapeHtml(item.id)}"
        ${item.id === weapon.loadedAmmoId ? "selected" : "disabled"}>${escapeHtml(item.name)} (DMG ${item.damage}, AP ${item.penetration})${item.id === weapon.loadedAmmoId ? "" : " - Not loaded"}</option>`).join("")}
    </select></label>
    <label class="field-stack"><span class="field-label">Rounds / Attack</span>
      <input type="number" min="1" max="${weapon.loaded}" step="1" value="${Math.min(Number(draft2.rounds) || 1, weapon.loaded) || 1}" ${prefix}="rounds" ${disabledAttr} ${weapon.loaded ? "" : "disabled"}></label>`;
}
async function editAmmunition(action, node) {
  if (!isEditable()) throw new Error("Only the GM can edit ammunition and reload weapons.");
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a token first.");
  const index = Number(node.dataset.rangedIndex);
  const ammoId = node.dataset.ammoId;
  const reloadId = action === "ammo-reload" ? getActionFieldValue(`[data-reload-ammo="${index}"]`) : "";
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const data = next.odyssey;
    data.ammunition ?? (data.ammunition = []);
    const weapon = data.weapons.ranged[index];
    if (action === "ammo-remove") {
      if (data.inventory.magazines.some((m) => m.loadedAmmoId === ammoId && m.loaded > 0)) throw new Error("Empty all magazines using this ammo first.");
      if (data.weapons.ranged.some((item) => item.loadedAmmoId === ammoId && item.loaded > 0)) throw new Error("Unload this ammunition before removing it.");
      data.ammunition = data.ammunition.filter((item) => item.id !== ammoId);
      for (const item of data.weapons.ranged) item.ammoIds = item.ammoIds.filter((id) => id !== ammoId);
    }
    if (action === "ammo-reload") {
      if (weapon && !weapon.magazineId) {
        const id = `reload-mag-${weapon.id}`;
        if (!data.inventory.magazines.some((m) => m.id === id)) data.inventory.magazines.push({ id, name: weapon.name + " magazine", capacity: weapon.capacity, loaded: 0, loadedAmmoId: "" });
        swapMagazine(data, index, id);
      }
      reloadMagazine(data, index, reloadId);
    }
    if (action === "ranged-remove" && weapon?.magazineId) swapMagazine(data, index, "");
    if ((action === "ranged-remove" || action === "ammo-unload") && weapon) {
      const loadedAmmo = data.ammunition.find((item) => item.id === weapon.loadedAmmoId);
      if (weapon.loaded && !loadedAmmo) throw new Error("Loaded ammunition is missing.");
      if (loadedAmmo) {
        if (loadedAmmo.quantity + weapon.loaded > 999999) throw new Error("Ammo reserve is full.");
        loadedAmmo.quantity += weapon.loaded;
      }
      weapon.loaded = 0;
      weapon.loadedAmmoId = "";
      if (action === "ranged-remove") data.weapons.ranged.splice(index, 1);
    }
    return next;
  });
}
function renderEnglishAttackBlock(token, data, tokenLocked) {
  if (!canViewAttackBlock(token)) return "";
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false
  );
  const disabledAttr = tokenLocked ? "disabled" : "";
  const pickDisabledAttr = tokenLocked || !targetCharacters.length ? "disabled" : "";
  const draft2 = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft2.skill, data.odyssey);
  const weaponOptions = buildWeaponOptions(getAttackSelectableWeapons(token), draft2.weaponName);
  const selectedTarget = targetCharacters.find((target) => target.id === draft2.targetTokenId) ?? null;
  const selectedTargetData = selectedTarget ? getTrackerData(selectedTarget) : null;
  const targetableBodyParts = getTargetableBodyParts(selectedTargetData, isEditable());
  const targetName = selectedTarget ? getCharacterName(selectedTarget) : draft2.targetTokenName || "No target selected";
  const isPickingTarget = targetPickState.active && targetPickState.attackerTokenId === token.id;
  const attackDisabledAttr = tokenLocked || !draft2.targetTokenId ? "disabled" : "";
  const weaponAccuracyDisplay = Number(draft2.weaponAccuracy) || 0;
  return renderCollapsibleSection(
    "Attack",
    `
      <div data-combat-layout="attack">
      <div class="combat-form-grid">
        <input type="hidden" value="${escapeHtml(draft2.targetTokenId)}" data-attack-field="targetTokenId">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon</span>
          <select data-attack-field="weaponName" ${disabledAttr}>${weaponOptions}</select>
          <span class="od-range-readout cp-muted" data-range-modifier="attack"></span>
        </label>
        <label class="field-stack">
          <span class="field-label">Current Target</span>
          <div class="combat-value">${escapeHtml(targetName)}</div>
          <span class="od-target-distance" data-target-distance title="Distance between token centres">${measuredTargetDistance?.sourceId === token.id ? escapeHtml(measuredTargetDistance.label) : "\u2014"}</span>
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
            ${targetableBodyParts.map(
      (part) => `<option value="${part}" ${part === draft2.targetPart ? "selected" : ""}>${escapeHtml(getBodyPartLabel(selectedTargetData, part))}</option>`
    ).join("")}
          </select>
        </label>
        <label class="field-stack ${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? "hidden" : ""}">
          <span class="field-label">Weapon Damage</span>
          <input type="number" value="${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? 0 : draft2.weaponDamage}" data-attack-field="weaponDamage" ${disabledAttr} ${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? "disabled" : ""}>
        </label>
        ${renderAttackAmmunition(token, draft2, false, disabledAttr)}
        <label class="field-stack">
          <span class="field-label">Weapon Accuracy</span>
          <div class="combat-value" data-weapon-accuracy-display>${weaponAccuracyDisplay >= 0 ? "+" : ""}${weaponAccuracyDisplay}</div>
        </label>
      </div>
      <div class="combat-modifiers">
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
      </div>
      <details class="combat-notes">
        <summary>Attack rules</summary>
        <div>Weapon Accuracy is added to Attack Bonus. Called shots: Head -30, arms/legs -15. STR Bonus affects enabled attack skills; ${escapeHtml(PARRY_SKILL_NAME)} is counted unless disabled.</div>
      </details>
      <div class="combat-action">
        <button type="button" class="success" data-action="perform-attack" ${attackDisabledAttr}>Attack</button>
      </div>
      </div>
    `,
    true
  );
}
function renderEnglishNoTargetAttackBlock(token, data, tokenLocked) {
  if (!canViewAttackBlock(token)) return "";
  const targetCharacters = getCharacters().filter(
    (item) => item.id !== token.id && item.visible !== false
  );
  const draft2 = getAttackDraft(token, data, targetCharacters);
  const skillOptions = buildSkillOptions(getAttackSkillEntries(data.odyssey), draft2.skill, data.odyssey);
  const weaponOptions = buildWeaponOptions(getAttackSelectableWeapons(token), draft2.weaponName);
  const disabledAttr = tokenLocked ? "disabled" : "";
  const weaponAccuracyDisplay = Number(draft2.weaponAccuracy) || 0;
  return renderCollapsibleSection(
    "No Target Attack",
    `
      <div data-combat-layout="manual">
      <div class="combat-form-grid">
        <label class="field-stack">
          <span class="field-label">Attack Skill</span>
          <select data-manual-attack-field="skill" ${disabledAttr}>${skillOptions}</select>
        </label>
        <label class="field-stack">
          <span class="field-label">Weapon</span>
          <select data-manual-attack-field="weaponName" ${disabledAttr}>${weaponOptions}</select>
          <span class="od-range-readout cp-muted" data-range-modifier="manual"></span>
        </label>
        <label class="field-stack"><span class="field-label">Distance \xB7 metres</span><input type="number" min="0" step="0.1" value="${escapeHtml(draft2.distanceMeters)}" data-manual-attack-field="distanceMeters" ${disabledAttr}></label>
        <label class="field-stack ${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? "hidden" : ""}">
          <span class="field-label">Weapon Damage</span>
          <input type="number" value="${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? 0 : draft2.weaponDamage}" data-manual-attack-field="weaponDamage" ${disabledAttr} ${getWeaponByName(token, draft2.weaponName).rangedIndex != null ? "disabled" : ""}>
        </label>
        ${renderAttackAmmunition(token, draft2, true, disabledAttr)}
        <label class="field-stack">
          <span class="field-label">Weapon Accuracy</span>
          <div class="combat-value" data-weapon-accuracy-display>${weaponAccuracyDisplay >= 0 ? "+" : ""}${weaponAccuracyDisplay}</div>
        </label>
      </div>
      <div class="combat-modifiers">
        <label class="field-stack">
          <span class="field-label">Attack Bonus</span>
          <input type="number" value="${draft2.manualAttackBonuses}" data-manual-attack-field="attackBonuses" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Attack Penalty</span>
          <input type="number" value="${draft2.manualAttackPenalties}" data-manual-attack-field="attackPenalties" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Armor</span>
          <input type="number" min="0" max="999" value="${draft2.manualArmor}" data-manual-attack-field="manualArmor" ${disabledAttr}>
        </label>
        <label class="field-stack">
          <span class="field-label">Parry</span>
          <input type="number" min="0" max="10" value="${draft2.manualParry}" data-manual-attack-field="manualParry" ${disabledAttr}>
        </label>
      </div>
      <details class="combat-notes">
        <summary>Manual defense</summary>
        <div>Uses Armor and Parry entered here and does not change the saved map target. Weapon Accuracy is added to Attack Bonus.</div>
      </details>
      <div class="combat-action">
        <button type="button" class="success" data-action="perform-manual-attack" ${disabledAttr}>No Target Attack</button>
      </div>
      </div>
    `,
    false
  );
}
function renderEnglishDiceBlock(token, data, tokenLocked) {
  if (!canViewDiceBlock(token)) return "";
  const attributeOptions = ATTRIBUTE_UI_FIELDS.map(
    ([key, label]) => `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${data.odyssey.attributes[key] ?? 0})</option>`
  ).join("");
  const skillOptions = buildGroupedSkillOptions(data.odyssey);
  const tokenLockedAttr = tokenLocked ? "disabled" : "";
  return renderCollapsibleSection(
    "Dice",
    `
      <div class="dice-rolls">
        <div class="dice-roll-row">
          <div class="dice-roll-fields dice-roll-fields-three">
            <label class="field-stack"><span class="field-label">Dice Sides</span><input type="number" min="2" value="20" data-roll-field="dice"></label>
            <label class="field-stack"><span class="field-label">Dice Count</span><input type="number" min="1" max="100" value="1" data-roll-field="count"></label>
            <label class="field-stack"><span class="field-label">Modifier</span><input type="number" value="0" data-roll-field="modifier"></label>
          </div>
          <button type="button" data-action="perform-roll-dice">Roll Dice</button>
        </div>
        <div class="dice-roll-row">
          <div class="dice-roll-fields">
            <label class="field-stack"><span class="field-label">Characteristic</span><select data-roll-char-field="attribute" ${tokenLockedAttr}>${attributeOptions}</select></label>
            <label class="field-stack"><span class="field-label">Bonus / Penalty</span><input type="number" value="0" data-roll-char-field="modifier" ${tokenLockedAttr}></label>
          </div>
          <button type="button" data-action="perform-roll-char" ${tokenLockedAttr}>Roll Characteristic</button>
        </div>
        <div class="dice-roll-row">
          <div class="dice-roll-fields">
            <label class="field-stack"><span class="field-label">Skill</span><select data-roll-skill-field="skill" ${tokenLockedAttr}>${skillOptions}</select></label>
            <label class="field-stack"><span class="field-label">Bonus / Penalty</span><input type="number" value="0" data-roll-skill-field="modifier" ${tokenLockedAttr}></label>
          </div>
          <button type="button" data-action="perform-roll-skill" ${tokenLockedAttr}>Roll Skill</button>
        </div>
      </div>
    `,
    false
  );
}
function renderTokenHud(token, data) {
  const targets = getCharacters().filter((t) => t.id !== token.id && t.visible !== false);
  const draft2 = getAttackDraft(token, data, targets);
  const target = targets.find((t) => t.id === draft2.targetTokenId);
  const targetData = target ? getTrackerData(target) : null;
  return compactHud({
    token,
    floating: HUD_MODE,
    tracker: data,
    target,
    targetData,
    draft: draft2,
    skills: buildSkillOptions(getAttackSkillEntries(data.odyssey), draft2.skill, data.odyssey),
    weapons: buildWeaponOptions(getAttackSelectableWeapons(token), draft2.weaponName),
    parts: getTargetableBodyParts(targetData, isEditable()).map((k) => [k, getBodyPartLabel(targetData, k)]),
    ammunition: renderAttackAmmunition(token, draft2, false, "")
  });
}
function renderSelectedToken() {
  const panelState = captureSelectedPanelState();
  activeTokenId = resolveActiveTokenId();
  const token = getCharacterById(activeTokenId);
  if (!token) {
    ui.selectionHint.textContent = "No character token selected";
    delete ui.selectedTokenPanel.dataset.tokenId;
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
  ui.selectedTokenPanel.dataset.tokenId = token.id;
  ui.selectedTokenPanel.innerHTML = `
    <div class="selected-card">
      ${HUD_MODE && canUseToken(token) ? renderTokenHud(token, data) : ""}
      <div class="selected-head">
        <div>
          <div class="token-name">${escapeHtml(getCharacterName(token))}</div>
          <div class="token-meta">${escapeHtml(token.id.slice(0, 8))} - ${tracked ? "Initialized" : "Auto-init on selection"} - ${tokenLocked ? "Read only" : "Controllable"}</div>
        </div>
        <div class="row row-gap">
          <button type="button" data-action="open-character-hud" class="secondary">Floating HUD</button>
          <button type="button" data-action="focus-token" class="secondary">Select On Map</button>
          <button type="button" data-action="reload-token-visuals" class="secondary" ${tokenLocked ? "disabled" : ""}>Reload Token</button>
        </div>
      </div>

      <div class="summary-strip">
        ${canViewBodyHpSummary(token) ? `
              <div class="stat-chip">
                <span class="chip-label">Body HP</span>
                <span class="chip-value">${totals.current}/${totals.max}</span>
              </div>
            ` : ""}
        <div class="stat-chip">
          <span class="chip-label">Assigned Player</span>
          <span class="chip-value">${escapeHtml(odyssey.owner.playerName || odyssey.owner.playerId || "Unassigned")}</span>
        </div>
      </div>

      ${isEditable() ? `
            ${renderOwnerFields({ odyssey }, gmOnlyDisabled)}
          ` : ""}
      ${canUseToken(token) ? effectChips(odyssey, isEditable()) : ""}
      ${canUseToken(token) ? renderCharacteristicsBlock({ odyssey }) : ""}
      ${canUseToken(token) ? renderEnglishSkillsBlock({ odyssey }, gmOnlyDisabled) : ""}
      ${canUseToken(token) && characterTools ? renderCollapsibleSection("Items", characterTools.renderInventory(data), false) : ""}



      ${renderEnglishAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishNoTargetAttackBlock(token, { odyssey }, tokenLocked)}
      ${renderEnglishDiceBlock(token, { odyssey }, tokenLocked)}
      ${SHOW_EMBEDDED_PUBLIC_LOG ? renderCollapsibleSection(
    "Last Roll",
    `<pre class="console-output">${lastRollText}</pre>`,
    false
  ) : ""}
      ${showPartBlock ? renderCollapsibleSection(
    "Body Parts",
    `
                <div class="row row-gap">
                  <button type="button" class="secondary" data-action="heal-limbs" ${bodyFieldDisabled}>Heal</button>
                </div>
                <div class="body-table-wrap">
                  <table class="body-table body-table-compact">
                    <thead><tr><th>Part</th><th>HP</th><th>Max</th><th>Armor</th><th>Position</th><th>Other</th></tr></thead>
                    <tbody>${[SPECIAL_PART_NAME, ...getBodyPartNames(data)].map((partName) => {
      const part = data.body[partName];
      const fixedPart = partName === "Torso" || partName === SPECIAL_PART_NAME;
      const penaltyDisabled = part.slot !== "other" ? "disabled" : bodyFieldDisabled;
      return `<tr>
                        <td><div class="body-part-name-line">${fixedPart ? `<span class="body-part-eye-spacer" title="Always visible"></span>` : `<label class="check-label body-part-visible" title="Visible to players"><input type="checkbox" data-action="toggle-part-hidden" data-part="${escapeHtml(partName)}" ${part.hidden ? "" : "checked"} ${bodyFieldDisabled}><span></span></label>`}<input class="body-part-name-input" type="text" maxlength="40" value="${escapeHtml(getBodyPartLabel(data, partName))}" data-action="rename-body-part" data-part="${escapeHtml(partName)}" ${fixedPart ? "disabled" : bodyFieldDisabled}>${fixedPart ? "" : `<button type="button" class="body-part-remove" data-action="remove-body-part" data-part="${escapeHtml(partName)}" title="Remove body part" aria-label="Remove ${escapeHtml(partName)}" ${bodyFieldDisabled}>\xD7</button>`}</div></td>
                        <td><div class="inline-stepper"><button type="button" data-action="change-part" data-part="${escapeHtml(partName)}" data-field="current" data-delta="-1" ${bodyFieldDisabled}>-</button><input type="text" inputmode="numeric" maxlength="3" value="${part.current}" data-action="set-field" data-part="${escapeHtml(partName)}" data-field="current" ${bodyFieldDisabled}><button type="button" data-action="change-part" data-part="${escapeHtml(partName)}" data-field="current" data-delta="1" ${bodyFieldDisabled}>+</button></div></td>
                        <td><input class="compact-input" type="text" inputmode="numeric" maxlength="3" value="${part.max}" data-action="set-field" data-part="${escapeHtml(partName)}" data-field="max" ${bodyFieldDisabled}></td>
                        <td><input class="compact-input" type="text" inputmode="numeric" maxlength="3" value="${part.armor}" data-action="set-field" data-part="${escapeHtml(partName)}" data-field="armor" ${bodyFieldDisabled}></td>
                        <td>${partName === "Torso" ? "Main" : partName === SPECIAL_PART_NAME ? "Special" : `<select data-action="set-field" data-part="${escapeHtml(partName)}" data-field="slot" ${bodyFieldDisabled}>${BODY_PART_SLOTS.map((slot) => `<option value="${slot}" ${part.slot === slot ? "selected" : ""}>${BODY_PART_SLOT_LABELS[slot]}</option>`).join("")}</select>`}</td>
                        <td><input class="compact-input" type="text" inputmode="numeric" maxlength="3" value="${part.attackPenalty ?? 0}" data-action="set-field" data-part="${escapeHtml(partName)}" data-field="attackPenalty" ${penaltyDisabled}></td>
                      </tr>`;
    }).join("")}</tbody>
                  </table>
                </div>
                ${bodyPartAddOpen ? `<div class="body-part-add">
                  <label class="field-stack"><span class="field-label">Part</span><input type="text" maxlength="40" data-body-field="new-name" placeholder="Extra arm" ${bodyFieldDisabled}></label>
                  <label class="field-stack"><span class="field-label">Position</span><select data-body-field="new-slot" ${bodyFieldDisabled}>${BODY_PART_SLOTS.map((slot) => `<option value="${slot}" ${slot === "other" ? "selected" : ""}>${BODY_PART_SLOT_LABELS[slot]}</option>`).join("")}</select></label>
                  <label class="field-stack"><span class="field-label">HP</span><input type="text" inputmode="numeric" maxlength="3" min="0" max="999" value="1" data-body-field="new-max" ${bodyFieldDisabled}></label>
                  <label class="field-stack"><span class="field-label">Armor</span><input type="text" inputmode="numeric" maxlength="3" min="0" max="999" value="0" data-body-field="new-armor" ${bodyFieldDisabled}></label>
                  <label class="field-stack"><span class="field-label">Other Penalty</span><input type="text" inputmode="numeric" maxlength="3" min="0" max="999" value="0" data-body-field="new-attack-penalty" ${bodyFieldDisabled}></label>
                  <label class="check-label body-part-visible" title="Visible to players"><input type="checkbox" data-body-field="new-visible" ${bodyFieldDisabled}><span></span></label>
                  <div class="body-part-add-actions"><button type="button" class="square-icon-button" data-action="add-body-part" title="Save body part" aria-label="Save body part" ${bodyFieldDisabled}>\u25A3</button><button type="button" class="square-icon-button danger" data-action="close-body-part-add" title="Cancel" aria-label="Cancel" ${bodyFieldDisabled}>\xD7</button></div>
                </div>` : ""}
                <div class="item-toolbar item-toolbar-bottom"><button type="button" class="square-icon-button" data-action="open-body-part-add" title="Add body part" aria-label="Add body part" ${bodyFieldDisabled}>+</button></div>
              `,
    true
  ) : ""}
      ${canViewOverlayPreview(token) ? renderCollapsibleSection(
    "Overlay Preview",
    `<pre class="console-output">${escapeHtml(formatOverlayPreviewText(data))}</pre>`,
    false
  ) : ""}
    </div>`;
  if (characterTools?.isBusy()) ui.selectedTokenPanel.querySelectorAll("[data-od-action]").forEach((e) => {
    if (e instanceof HTMLButtonElement) e.disabled = true;
  });
  arrangeCombatPanels(ui.selectedTokenPanel);
  paintRangeAdjustments();
  refreshTargetDistance();
  restoreSelectedPanelState(panelState);
}
function render() {
  ui.roleBadge.textContent = playerRole === "GM" ? "GM" : "PLAYER";
  ui.clearDebugBtn?.classList.toggle("hidden", playerRole !== "GM");
  renderSelectedToken();
  renderDebugConsole();
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
  setPartyPlayersState(players);
  setSceneItems(items);
  setSelectionState(selection);
  const selectedCharacterId = selectionIds.find(
    (selectionId) => charactersByIdCache.has(selectionId)
  );
  if (selectedCharacterId) {
    activeTokenId = selectedCharacterId;
    const initialized = await initializeCharacterToken(selectedCharacterId);
    if (initialized) {
      setSceneItems(await lib_default.scene.items.getItems());
    }
  } else if (activeTokenId && !charactersByIdCache.has(activeTokenId)) {
    activeTokenId = resolveActiveTokenId();
  }
  await syncLocalOwnedHiddenTokenViews(sceneItems);
  scheduleRender();
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
  scheduleRender();
  await syncTargetHighlight();
}
async function reloadTokenVisuals() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot reload this token.", "error");
    return;
  }
  const currentItems = await lib_default.scene.items.getItems();
  const liveToken = currentItems.find((item) => item.id === token.id) ?? token;
  if (liveToken.visible === false) {
    await syncLocalOwnedHiddenTokenViews(currentItems);
  } else {
    await removeOverlaysForToken(token.id, currentItems);
    await ensureOverlayForToken(token.id);
  }
  setSceneItems(await lib_default.scene.items.getItems());
  await syncLocalOwnedHiddenTokenViews(sceneItems);
  scheduleRender();
  await syncTargetHighlight();
  setStatus(`${getCharacterName(token)} visuals reloaded.`, "success");
}
async function healLimbs() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canEditTokenData(token)) {
    setStatus("Only the GM or assigned player can heal this token.", "error");
    return;
  }
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    for (const partName of getBodyPartNames(next)) {
      const part = next.body?.[partName];
      if (!part) continue;
      setCombinedPartHp(next, partName, part.max);
      part.minor = 0;
      part.serious = 0;
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
  setStatus(`${getCharacterName(token)} healed.`, "success");
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
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const part = next.body[partName];
    if (!part) return next;
    if (field === "current") {
      setCombinedPartHp(next, partName, part.current + delta);
    } else if (field === "max") {
      setCombinedPartMax(next, partName, part.max + delta);
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      part.armor = clamp(part.armor + delta, 0, 999);
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
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
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const part = next.body[partName];
    if (!part) return next;
    const numericValue = clamp(Number(value) || 0, 0, 999);
    if (field === "current") {
      setCombinedPartHp(next, partName, numericValue);
    } else if (field === "max") {
      setCombinedPartMax(next, partName, numericValue);
      part.current = clamp(part.current, 0, part.max);
    } else if (field === "armor") {
      const inv = next.odyssey.inventory;
      if (Object.hasOwn(inv.baseArmor || {}, partName)) inv.baseArmor[partName] = Math.max(0, numericValue - (part.armor - inv.baseArmor[partName]));
      part.armor = numericValue;
    } else if (field === "attackPenalty") {
      part.attackPenalty = numericValue;
    } else if (field === "slot" && BODY_PART_SLOTS.includes(value)) {
      part.slot = value;
      part.attackPenalty = getBodyPartAttackPenalty(next.body, partName);
    }
    return next;
  });
  await ensureOverlayForToken(token.id);
}
async function addBodyPart() {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!canEditTokenData(token)) throw new Error("Only the GM or assigned player can edit this token.");
  const name = getActionFieldValue('[data-body-field="new-name"]').trim();
  const slot = getActionFieldValue('[data-body-field="new-slot"]');
  const max = clamp(Number(getActionFieldValue('[data-body-field="new-max"]')) || 0, 0, 999);
  const armor = clamp(Number(getActionFieldValue('[data-body-field="new-armor"]')) || 0, 0, 999);
  const attackPenalty = clamp(Number(getActionFieldValue('[data-body-field="new-attack-penalty"]')) || 0, 0, 999);
  const hidden = document.querySelector('[data-body-field="new-visible"]')?.checked !== true;
  if (!name) throw new Error("Enter a body part name first.");
  if (!BODY_PART_SLOTS.includes(slot)) throw new Error("Choose a body part position first.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    if (next.body[name]) throw new Error("A body part with this name already exists.");
    next.body[name] = {
      current: max,
      max,
      armor,
      minor: 0,
      serious: 0,
      slot,
      attackPenalty,
      hidden,
      label: name
    };
    return next;
  });
  await ensureOverlayForToken(token.id);
}
async function removeBodyPart(partName) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!canEditTokenData(token)) throw new Error("Only the GM or assigned player can edit this token.");
  if (partName === "Torso") throw new Error("Torso cannot be removed.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    delete next.body[partName];
    return next;
  });
  await ensureOverlayForToken(token.id);
}
async function renameBodyPart(partName, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!canEditTokenData(token)) throw new Error("Only the GM or assigned player can edit this token.");
  const nextName = String(value ?? "").trim().slice(0, 40);
  if (!nextName) throw new Error("Enter a body part name first.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    if (!next.body[partName] || nextName === partName) return next;
    next.body[partName].label = nextName;
    return next;
  });
  await ensureOverlayForToken(token.id);
}
async function toggleBodyPartHidden(partName, hidden) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!canEditTokenData(token)) throw new Error("Only the GM or assigned player can edit this token.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    if (next.body[partName] && partName !== "Torso") next.body[partName].hidden = hidden;
    return next;
  });
  await ensureOverlayForToken(token.id);
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
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    next.odyssey ?? (next.odyssey = structuredClone(getTrackerData(token).odyssey));
    (_a = next.odyssey).owner ?? (_a.owner = { playerId: "", playerName: "" });
    next.odyssey.owner.playerId = selectedPlayer?.id ?? "";
    next.odyssey.owner.playerName = selectedPlayer?.name ?? "";
    return next;
  });
  await ensureOverlayForToken(token.id);
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
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    next.odyssey.skills[skill] = clamp(Number(value) || 0, 0, 10);
    return next;
  });
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
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey).skillStrengthBonuses ?? (_a.skillStrengthBonuses = {});
    next.odyssey.skillStrengthBonuses[skill] = Boolean(enabled);
    return next;
  });
}
async function renameOdysseySkill(skill, label) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!isEditable()) throw new Error("Only the GM can edit Odyssey skills.");
  const nextLabel = String(label ?? "").trim().slice(0, 40);
  if (!nextLabel) throw new Error("Enter a skill name first.");
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    if (!next.odyssey.skills[skill] && next.odyssey.skills[skill] !== 0) return next;
    (_a = next.odyssey).skillLabels ?? (_a.skillLabels = {});
    next.odyssey.skillLabels[skill] = nextLabel;
    return next;
  });
}
async function addOdysseySkillPerk(skill, perk) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!isEditable()) throw new Error("Only the GM can edit skill perks.");
  const name = String(perk ?? "").trim().slice(0, 40);
  if (!name) throw new Error("Enter a perk name first.");
  await updateTrackerData2(token.id, (current2) => {
    var _a, _b;
    const next = structuredClone(current2);
    (_a = next.odyssey).skillPerks ?? (_a.skillPerks = {});
    const perks = (_b = next.odyssey.skillPerks)[skill] ?? (_b[skill] = []);
    if (!perks.includes(name) && perks.length < 12) perks.push(name);
    return next;
  });
}
async function removeOdysseySkillPerk(skill, perkIndex) {
  const token = getCharacterById(activeTokenId);
  if (!token) throw new Error("Select a character first.");
  if (!isEditable()) throw new Error("Only the GM can edit skill perks.");
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const perks = next.odyssey.skillPerks?.[skill];
    if (Array.isArray(perks)) perks.splice(perkIndex, 1);
    return next;
  });
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
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    next.odyssey.attributes[attribute] = clamp(Number(value) || 0, 0, 20);
    return next;
  });
}
async function addWeapon() {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }
  const name = getActionFieldValue('[data-weapon-field="new-name"]').trim() || "New Weapon";
  const damage = clamp(Number(getActionFieldValue('[data-weapon-field="new-damage"]')) || 0, -999, 999);
  const accuracy = clamp(Number(getActionFieldValue('[data-weapon-field="new-accuracy"]')) || 0, -999, 999);
  const rangedField = ui.selectedTokenPanel.querySelector(
    '[data-weapon-field="new-ranged"]'
  );
  const ranged = rangedField instanceof HTMLInputElement && rangedField.checked;
  await updateTrackerData2(token.id, (current2) => {
    var _a, _b;
    const next = structuredClone(current2);
    if (ranged) {
      (_a = next.odyssey.weapons).ranged ?? (_a.ranged = []);
      next.odyssey.weapons.ranged.push({
        id: crypto.randomUUID(),
        name,
        damage: 0,
        accuracy,
        capacity: 0,
        loaded: 0,
        loadedAmmoId: "",
        ammoIds: []
      });
    } else {
      (_b = next.odyssey.weapons).melee ?? (_b.melee = []);
      next.odyssey.weapons.melee.push({
        name,
        damage,
        accuracy
      });
    }
    return next;
  });
  if (!ranged) {
    syncAttackWeaponInputs(token.id, name, damage, accuracy);
  }
  setStatus(`Weapon "${name}" saved for ${getCharacterName(token)}.`, "success");
}
async function setWeaponDamage(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }
  const currentWeapon = getOdysseyData(token).weapons?.melee?.[index] ?? { name: "Default", damage: 0, accuracy: 0 };
  const currentWeaponName = currentWeapon.name ?? "Default";
  const nextDamage = clamp(Number(value) || 0, -999, 999);
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0, accuracy: 0 };
    }
    next.odyssey.weapons.melee[index].damage = nextDamage;
    return next;
  });
  const currentDraft = attackFormDrafts.get(token.id) ?? {};
  if (currentDraft.weaponName === currentWeaponName) {
    syncAttackWeaponInputs(token.id, currentWeaponName, nextDamage, currentWeapon.accuracy ?? 0);
  }
}
async function setWeaponAccuracy(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }
  const currentWeapon = getOdysseyData(token).weapons?.melee?.[index] ?? { name: "Default", damage: 0, accuracy: 0 };
  const currentWeaponName = currentWeapon.name ?? "Default";
  const nextAccuracy = clamp(Number(value) || 0, -999, 999);
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0, accuracy: 0 };
    }
    next.odyssey.weapons.melee[index].accuracy = nextAccuracy;
    return next;
  });
  const currentDraft = attackFormDrafts.get(token.id) ?? {};
  if (currentDraft.weaponName === currentWeaponName) {
    syncAttackWeaponInputs(token.id, currentWeaponName, currentWeapon.damage ?? 0, nextAccuracy);
  }
}
async function setWeaponName(index, value) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }
  const previousWeapon = getOdysseyData(token).weapons?.melee?.[index] ?? { name: "Default", damage: 0, accuracy: 0 };
  const nextWeaponName = String(value || "").trim() || "Default";
  await updateTrackerData2(token.id, (current2) => {
    var _a;
    const next = structuredClone(current2);
    (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
    if (!next.odyssey.weapons.melee[index]) {
      next.odyssey.weapons.melee[index] = { name: "Default", damage: 0, accuracy: 0 };
    }
    next.odyssey.weapons.melee[index].name = nextWeaponName;
    return next;
  });
  const currentDraft = attackFormDrafts.get(token.id) ?? {};
  if (currentDraft.weaponName === previousWeapon.name) {
    syncAttackWeaponInputs(token.id, nextWeaponName, previousWeapon.damage, previousWeapon.accuracy ?? 0);
  }
}
async function removeWeapon(type = "melee", index) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!isEditable()) {
    setStatus("Only the GM can edit weapons.", "error");
    return;
  }
  const currentWeapons = getOdysseyData(token).weapons?.[type] ?? [];
  const removedWeapon = currentWeapons[index];
  if (!removedWeapon) {
    setStatus("Weapon not found.", "error");
    return;
  }
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    const weapons = next.odyssey.weapons[type] ?? [];
    const weapon = weapons[index];
    if (type === "ranged" && weapon?.magazineId) swapMagazine(next.odyssey, index, "");
    if (type === "ranged" && weapon?.loaded > 0) {
      const ammo = next.odyssey.ammunition.find(
        (item) => item.id === weapon.loadedAmmoId
      );
      if (ammo) {
        ammo.quantity += weapon.loaded;
      }
    }
    next.odyssey.weapons[type] = weapons.filter(
      (_weapon, weaponIndex) => weaponIndex !== index
    );
    return next;
  });
  const refreshedToken = getCharacterById(token.id);
  const fallbackWeapon = refreshedToken ? getDefaultAttackWeapon(refreshedToken) : { name: UNARMED_WEAPON_NAME, damage: 0, accuracy: 0 };
  syncAttackWeaponInputs(token.id, fallbackWeapon.name, fallbackWeapon.damage, fallbackWeapon.accuracy ?? 0);
  setStatus(`Weapon "${removedWeapon.name}" removed.`, "success");
}
async function autosaveDraftField(draft2) {
  const token = getCharacterById(draft2.tokenId);
  if (!token) return;
  await updateTrackerData2(token.id, (current2) => {
    var _a, _b, _c;
    const next = structuredClone(current2);
    if (draft2.action === "set-odyssey-skill") {
      if (!isEditable()) return next;
      next.odyssey.skills[draft2.skill] = clamp(Number(draft2.value) || 0, 0, 10);
      return next;
    }
    if (draft2.action === "set-odyssey-attribute") {
      if (!isEditable()) return next;
      next.odyssey.attributes[draft2.attribute] = clamp(Number(draft2.value) || 0, 0, 20);
      return next;
    }
    if (draft2.action === "set-weapon-damage") {
      (_a = next.odyssey.weapons).melee ?? (_a.melee = []);
      if (!next.odyssey.weapons.melee[draft2.weaponIndex]) {
        next.odyssey.weapons.melee[draft2.weaponIndex] = { name: "Default", damage: 0, accuracy: 0 };
      }
      next.odyssey.weapons.melee[draft2.weaponIndex].damage = clamp(Number(draft2.value) || 0, -999, 999);
      return next;
    }
    if (draft2.action === "set-weapon-accuracy") {
      (_b = next.odyssey.weapons).melee ?? (_b.melee = []);
      if (!next.odyssey.weapons.melee[draft2.weaponIndex]) {
        next.odyssey.weapons.melee[draft2.weaponIndex] = { name: "Default", damage: 0, accuracy: 0 };
      }
      next.odyssey.weapons.melee[draft2.weaponIndex].accuracy = clamp(Number(draft2.value) || 0, -999, 999);
      return next;
    }
    if (draft2.action === "set-weapon-name") {
      (_c = next.odyssey.weapons).melee ?? (_c.melee = []);
      if (!next.odyssey.weapons.melee[draft2.weaponIndex]) {
        next.odyssey.weapons.melee[draft2.weaponIndex] = { name: "Default", damage: 0, accuracy: 0 };
      }
      next.odyssey.weapons.melee[draft2.weaponIndex].name = String(draft2.value || "").trim() || "Default";
      return next;
    }
    if (draft2.action === "set-field") {
      if (!canEditTokenData(token)) return next;
      const part = next.body[draft2.partName];
      if (!part) return next;
      const numericValue = clamp(Number(draft2.value) || 0, 0, 999);
      if (draft2.field === "current") {
        setCombinedPartHp(next, draft2.partName, numericValue);
      } else if (draft2.field === "max") {
        setCombinedPartMax(next, draft2.partName, numericValue);
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
var attackPending = false;
async function performAttack(options = {}) {
  if (attackPending) return;
  attackPending = true;
  try {
    await performAttackOnce(options);
  } finally {
    attackPending = false;
  }
}
async function performAttackOnce({ manualDefense = false } = {}) {
  const attacker = getCharacterById(activeTokenId);
  if (!attacker) {
    setStatus("Select an attacker token first.", "error");
    return;
  }
  if (!canUseToken(attacker)) {
    setStatus("You cannot roll for this attacker token.", "error");
    return;
  }
  const savedTargetTokenId = getActionFieldValue('[data-attack-field="targetTokenId"]') || "";
  const targetTokenId = manualDefense ? "" : savedTargetTokenId;
  const target = targetTokenId ? getCharacterById(targetTokenId) : null;
  if (!manualDefense && !targetTokenId) {
    setStatus("Pick a target on the map first.", "error");
    return;
  }
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
  const attackerOdyssey = getOdysseyData(attacker);
  const targetData = target ? getTrackerData(target) : null;
  const targetOdyssey = target ? getOdysseyData(target) : null;
  const skillName = manualDefense ? getActionFieldValue('[data-manual-attack-field="skill"]') || getActionFieldValue('[data-attack-field="skill"]') : getActionFieldValue('[data-attack-field="skill"]');
  const weaponName = manualDefense ? getActionFieldValue('[data-manual-attack-field="weaponName"]') || getActionFieldValue('[data-attack-field="weaponName"]') : getActionFieldValue('[data-attack-field="weaponName"]');
  const selectedWeapon = getWeaponByName(attacker, weaponName);
  const shotPrefix = manualDefense ? "data-manual-attack-field" : "data-attack-field";
  const shot = selectedWeapon.rangedIndex == null ? null : getAmmoShot(
    attackerOdyssey,
    selectedWeapon.rangedIndex,
    getActionFieldValue(`[${shotPrefix}="ammoId"]`),
    Number(getActionFieldValue(`[${shotPrefix}="rounds"]`))
  );
  const requestedTargetPart = manualDefense ? "Torso" : getActionFieldValue('[data-attack-field="targetPart"]');
  const weaponDamage = Number(
    manualDefense ? getActionFieldValue('[data-manual-attack-field="weaponDamage"]') || getActionFieldValue('[data-attack-field="weaponDamage"]') : getActionFieldValue('[data-attack-field="weaponDamage"]')
  ) || 0;
  const manualAttackBonuses = Number(
    manualDefense ? getActionFieldValue('[data-manual-attack-field="attackBonuses"]') || getActionFieldValue('[data-attack-field="attackBonuses"]') : getActionFieldValue('[data-attack-field="attackBonuses"]')
  ) || 0;
  const weaponAccuracy = Number(selectedWeapon?.accuracy) || 0;
  const attackDistance = target && targetDistance ? await targetDistance.measure(attacker, target) : null;
  const enteredDistance = manualDefense ? getActionFieldValue('[data-manual-attack-field="distanceMeters"]') : null;
  const meters = manualDefense ? String(enteredDistance ?? "").trim() === "" ? null : Number(enteredDistance) : attackDistance?.meters;
  const range = rangeAdjustment(selectedWeapon, meters, attackerOdyssey, skillName);
  if (range.required) throw new Error(manualDefense ? "Enter distance in metres for this weapon." : "Set scene ruler units to metres or feet to calculate weapon range.");
  const attackBonuses = manualAttackBonuses + weaponAccuracy + Math.max(0, range.modifier);
  const manualAttackPenalties = Number(
    manualDefense ? getActionFieldValue('[data-manual-attack-field="attackPenalties"]') || getActionFieldValue('[data-attack-field="attackPenalties"]') : getActionFieldValue('[data-attack-field="attackPenalties"]')
  ) || 0;
  const availableTargetParts = getTargetableBodyParts(targetData, isEditable());
  const targetPart = manualDefense ? "Torso" : availableTargetParts.includes(requestedTargetPart) ? requestedTargetPart : "Torso";
  const automaticTargetPenalty = manualDefense ? 0 : getBodyPartAttackPenalty(targetData, targetPart);
  const totalAttackPenalties = manualAttackPenalties + automaticTargetPenalty + Math.max(0, -range.modifier);
  const defenseBonuses = Number(getActionFieldValue('[data-attack-field="defenseBonuses"]')) || 0;
  const defensePenalties = Number(getActionFieldValue('[data-attack-field="defensePenalties"]')) || 0;
  const manualArmor = clamp(
    Number(
      manualDefense ? getActionFieldValue('[data-manual-attack-field="manualArmor"]') || getActionFieldValue('[data-attack-field="manualArmor"]') : getActionFieldValue('[data-attack-field="manualArmor"]')
    ) || 0,
    0,
    999
  );
  const manualParry = clamp(
    Number(
      manualDefense ? getActionFieldValue('[data-manual-attack-field="manualParry"]') || getActionFieldValue('[data-attack-field="manualParry"]') : getActionFieldValue('[data-attack-field="manualParry"]')
    ) || 0,
    0,
    10
  );
  const parryMode = getActionFieldValue('[data-attack-field="parryMode"]') || "off";
  const parryDivisor = getParryDivisor(parryMode);
  saveAttackDraftValue(attacker.id, "skill", skillName);
  saveAttackDraftValue(attacker.id, "weaponName", weaponName);
  if (!manualDefense) {
    saveAttackDraftValue(attacker.id, "targetTokenId", targetTokenId);
    saveAttackDraftValue(attacker.id, "targetTokenName", target ? getCharacterName(target) : "");
    await persistAttackTargetToken(attacker.id, targetTokenId);
  }
  saveAttackDraftValue(attacker.id, "targetPart", targetPart);
  saveAttackDraftValue(attacker.id, "weaponDamage", String(weaponDamage));
  saveAttackDraftValue(
    attacker.id,
    manualDefense ? "manualAttackBonuses" : "attackBonuses",
    manualDefense ? getActionFieldValue('[data-manual-attack-field="attackBonuses"]') : getActionFieldValue('[data-attack-field="attackBonuses"]')
  );
  saveAttackDraftValue(
    attacker.id,
    manualDefense ? "manualAttackPenalties" : "attackPenalties",
    manualDefense ? getActionFieldValue('[data-manual-attack-field="attackPenalties"]') : getActionFieldValue('[data-attack-field="attackPenalties"]')
  );
  saveAttackDraftValue(attacker.id, "defenseBonuses", getActionFieldValue('[data-attack-field="defenseBonuses"]'));
  saveAttackDraftValue(attacker.id, "defensePenalties", getActionFieldValue('[data-attack-field="defensePenalties"]'));
  saveAttackDraftValue(
    attacker.id,
    "manualArmor",
    manualDefense ? getActionFieldValue('[data-manual-attack-field="manualArmor"]') : getActionFieldValue('[data-attack-field="manualArmor"]')
  );
  saveAttackDraftValue(
    attacker.id,
    "manualParry",
    manualDefense ? getActionFieldValue('[data-manual-attack-field="manualParry"]') : getActionFieldValue('[data-attack-field="manualParry"]')
  );
  saveAttackDraftValue(attacker.id, "parryMode", parryMode);
  const specialPartState = targetData?.body?.[SPECIAL_PART_NAME] ?? null;
  const specialWasActive = Boolean(target) && hasConfiguredSpecial(targetData) && (Number(specialPartState?.max) || 0) > 0 && (Number(specialPartState?.current) || 0) > 0;
  const baseTargetArmor = target ? (Number(targetData?.body?.[targetPart]?.armor) || 0) + (specialWasActive ? Number(specialPartState?.armor) || 0 : 0) : manualArmor;
  const targetArmor = effectiveArmor(baseTargetArmor, shot?.penetration ?? 0);
  const targetPartState = targetData?.body?.[targetPart] ?? { current: 0, max: 0, armor: 0, minor: 0, serious: 0 };
  const beforeHp = target ? targetPartState.current ?? 0 : null;
  const beforeMinor = target ? targetPartState.minor ?? 0 : null;
  const beforeSerious = target ? targetPartState.serious ?? 0 : null;
  const specialBeforeHp = specialWasActive ? Number(specialPartState?.current) || 0 : null;
  const strengthBonus = !shot && getSkillStrengthBonusFlag(attackerOdyssey, skillName) ? Math.max((attackerOdyssey.attributes.Strength ?? 0) - 10, 0) : 0;
  const finalWeaponDamage = shot ? shot.damage : weaponDamage + strengthBonus;
  const baseTargetParry = target ? targetOdyssey?.skills?.[PARRY_SKILL_NAME] ?? 0 : manualParry;
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
  const specialResolution = target && result.hit && result.damage ? projectDamageWithSpecialProtection({
    specialPart: specialPartState,
    targetPart: targetPartState,
    damage: result.damage,
    targetPartName: targetPart
  }) : {
    specialProjectedState: specialPartState ? getNormalizedPartState(specialPartState) : null,
    projectedTargetState: {
      ...target ? getNormalizedPartState(targetPartState) : {},
      critApplied: 0
    },
    specialActive: false,
    specialArmor: 0,
    damageAppliedLabel: "No Damage"
  };
  const projectedPartState = specialResolution.projectedTargetState;
  const equipmentCrit = projectedPartState.critApplied || 0;
  if (target && result.hit && equipmentCrit) {
    const preview = getTrackerData(target), equipment = absorbEquipmentCritical(preview, result.targetPart, equipmentCrit);
    projectedPartState.current = Math.max(0, (targetPartState.current || 0) - equipmentCrit);
  }
  const projectedSpecialState = specialResolution.specialProjectedState;
  const afterHp = target ? projectedPartState.current ?? beforeHp : null;
  const afterMinor = target ? projectedPartState.minor ?? beforeMinor : null;
  const afterSerious = target ? projectedPartState.serious ?? beforeSerious : null;
  const specialAfterHp = specialWasActive ? projectedSpecialState?.current ?? specialBeforeHp : null;
  const resolvedTargetName = target ? getCharacterName(target) : "Manual Defense";
  const damageSummary = specialResolution.specialActive && result.hit && specialResolution.damageAppliedLabel !== "No Damage" ? `${result.summary} Applied: ${specialResolution.damageAppliedLabel}.` : result.summary;
  const resolvedAttackSummary = shot ? `${damageSummary} ${shot.name} x${shot.count}; ammo damage ${shot.damage}; armor ${baseTargetArmor} -> ${targetArmor} (AP ${shot.penetration}).` : damageSummary;
  await updateTrackerData2(attacker.id, (current2) => {
    const next = structuredClone(current2);
    if (shot) {
      if (next.odyssey.weapons.ranged[selectedWeapon.rangedIndex]?.id !== selectedWeapon.id) throw new Error("Weapon changed. Please retry the attack.");
      consumeAmmo(next.odyssey, selectedWeapon.rangedIndex, shot);
    }
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary: `${getCharacterName(attacker)} -> ${resolvedTargetName}: ${resolvedAttackSummary}`,
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
    await updateTrackerData2(target.id, (current2) => {
      const next = structuredClone(current2);
      if (specialResolution.specialActive && next.body[SPECIAL_PART_NAME]) {
        setCombinedPartHp(next, SPECIAL_PART_NAME, projectedSpecialState.current);
        next.body[SPECIAL_PART_NAME].minor = projectedSpecialState.minor;
        next.body[SPECIAL_PART_NAME].serious = projectedSpecialState.serious;
      }
      if (result.hit && next.body[result.targetPart]) {
        absorbEquipmentCritical(next, result.targetPart, equipmentCrit);
        next.body[result.targetPart].current = projectedPartState.current;
        next.body[result.targetPart].minor = projectedPartState.minor;
        next.body[result.targetPart].serious = projectedPartState.serious;
      }
      next.lastRoll = {
        eventId: 0,
        actorName: getCharacterName(attacker),
        summary: resolvedAttackSummary,
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
  if (target) {
    await ensureOverlayForToken(target.id);
  }
  await pushDebugEntry(
    `${getAttackOutcomeIcon(result.outcome)} ${getCharacterName(attacker)} attacks ${resolvedTargetName}`,
    formatAttackDebug({
      distance: attackDistance || (Number.isFinite(meters) ? { label: meters + " m (manual)" } : null),
      range,
      attackerName: getCharacterName(attacker),
      targetName: resolvedTargetName,
      targetPart,
      attackSkillName: skillName,
      attackSkillValue: attackerOdyssey.skills[skillName] ?? 0,
      weaponDamage: finalWeaponDamage,
      weaponAccuracy,
      strengthBonus,
      manualAttackBonuses,
      totalAttackBonuses: attackBonuses,
      manualAttackPenalties,
      automaticTargetPenalty,
      totalAttackPenalties,
      defenseBonuses,
      defensePenalties,
      baseTargetParry,
      targetParry,
      parryMode,
      targetArmor,
      specialArmor: specialResolution.specialArmor,
      specialActive: specialResolution.specialActive,
      result,
      beforeHp,
      afterHp,
      specialBeforeHp,
      specialAfterHp,
      beforeMinor,
      afterMinor,
      beforeSerious,
      afterSerious,
      critApplied: projectedPartState.critApplied ?? 0,
      damageAppliedLabel: specialResolution.damageAppliedLabel
    }),
    result.hit ? "success" : "info"
  );
  setStatus(`${getCharacterName(attacker)} -> ${resolvedTargetName}: ${resolvedAttackSummary}`, result.hit ? "success" : "info");
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
  const modifier2 = Number(getActionFieldValue('[data-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier2, count);
  const diceLabel = `${result.count}d${result.sides}`;
  const summary = buildDiceRollSummary(diceLabel, result);
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    next.lastRoll = {
      eventId: 0,
      actorName: playerName || "Owlbear Player",
      summary,
      outcome: "roll",
      total: result.total,
      targetPart: "",
      timestamp: (/* @__PURE__ */ new Date()).toISOString(),
      source: "owlbear-extension"
    };
    next.history = [next.lastRoll, ...next.history ?? []].slice(0, 12);
    return next;
  });
  await pushDebugEntry(`${getCharacterName(token)} rolls dice`, formatDiceDebug({
    tokenName: getCharacterName(token),
    result
  }), "success");
  setStatus(summary, "success");
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
  const category = name === MELEE_SKILL_NAME || name === PARRY_SKILL_NAME ? COMBAT_SKILL_CATEGORY : getActionFieldValue('[data-skill-field="new-category"]') === COMBAT_SKILL_CATEGORY ? COMBAT_SKILL_CATEGORY : getActionFieldValue('[data-skill-field="new-category"]') === ABILITIES_SKILL_CATEGORY ? ABILITIES_SKILL_CATEGORY : APPLIED_SKILL_CATEGORY;
  await updateTrackerData2(token.id, (current2) => {
    var _a, _b, _c, _d;
    const next = structuredClone(current2);
    next.odyssey.skills[name] = value;
    (_a = next.odyssey).skillCategories ?? (_a.skillCategories = {});
    (_b = next.odyssey).skillStrengthBonuses ?? (_b.skillStrengthBonuses = {});
    (_c = next.odyssey).skillLabels ?? (_c.skillLabels = {});
    (_d = next.odyssey).skillPerks ?? (_d.skillPerks = {});
    next.odyssey.skillCategories[name] = category;
    next.odyssey.skillStrengthBonuses[name] = name === MELEE_SKILL_NAME ? true : name === PARRY_SKILL_NAME ? false : Boolean(addStrengthBonus);
    next.odyssey.skillLabels[name] = name;
    next.odyssey.skillPerks[name] = [];
    return next;
  });
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
  await updateTrackerData2(token.id, (current2) => {
    const next = structuredClone(current2);
    delete next.odyssey.skills[skillName];
    delete next.odyssey.skillCategories?.[skillName];
    delete next.odyssey.skillStrengthBonuses?.[skillName];
    delete next.odyssey.skillLabels?.[skillName];
    delete next.odyssey.skillPerks?.[skillName];
    return next;
  });
  setStatus(`Skill "${skillName}" removed.`, "success");
}
async function performRollCharOnce(options = {}) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }
  const attribute = options.attribute ?? (getActionFieldValue('[data-roll-char-field="attribute"]') || "Strength");
  if (!ATTRIBUTE_UI_FIELDS.some(([key]) => key === attribute)) throw new Error("Unknown characteristic.");
  const modifier2 = options.modifier ?? (Number(getActionFieldValue('[data-roll-char-field="modifier"]')) || 0);
  const odyssey = getOdysseyData(token);
  const result = rollCharacterCheck(odyssey.attributes[attribute] ?? 0, modifier2);
  const attributeLabel = ATTRIBUTE_UI_FIELDS.find(([key]) => key === attribute)?.[1] ?? attribute;
  const summary = `${getResolvedCheckResultIcon(result.result)} Characteristic ${attributeLabel}: ${result.roll} vs ${result.finalAttribute} (${result.result})`;
  await updateTrackerData2(token.id, (current2) => {
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
    `${getResolvedCheckResultIcon(result.result)} ${getCharacterName(token)} rolls characteristic`,
    formatRollCharDebug({
      tokenName: getCharacterName(token),
      attributeLabel,
      result
    }),
    isResolvedCheckResultSuccess(result.result) ? "success" : result.result === "Critical Failure" ? "error" : "info"
  );
  setStatus(summary, isResolvedCheckResultSuccess(result.result) ? "success" : "error");
}
async function useCharacterItem(sourceId, itemId, targetId, partName) {
  const effectId = crypto.randomUUID();
  await lib_default.scene.items.updateItems([.../* @__PURE__ */ new Set([sourceId, targetId])], (items) => {
    const source = items.find((t) => t.id === sourceId), target = items.find((t) => t.id === targetId);
    if (!source || !target || !canUseToken(source)) throw new Error("Character is unavailable.");
    if (playerRole !== "GM" && (target.visible === false || getTrackerData(target).body[partName]?.hidden)) throw new Error("Target is unavailable.");
    const sourceData = getTrackerData(source), targetData = source === target ? sourceData : getTrackerData(target);
    const item = sourceData.odyssey.inventory.items.find((i) => i.id === itemId);
    if (!item || !["consumable", "medical"].includes(item.kind) || item.quantity < 1) throw new Error("No usable item available.");
    useItemOnTracker(targetData, item, partName, effectId);
    item.quantity -= 1;
    source.metadata[META_KEY] = sanitizeTrackerData(sourceData);
    target.metadata[META_KEY] = sanitizeTrackerData(targetData);
  });
  await syncState();
  await Promise.all([.../* @__PURE__ */ new Set([sourceId, targetId])].map(ensureOverlayForToken));
  scheduleRender();
}
async function performRollSkillOnce(options = {}) {
  const token = getCharacterById(activeTokenId);
  if (!token) {
    setStatus("Select a character first.", "error");
    return;
  }
  if (!canUseToken(token)) {
    setStatus("You cannot roll for this token.", "error");
    return;
  }
  const skillName = options.raw ? "Random d100" : options.skill ?? getActionFieldValue('[data-roll-skill-field="skill"]');
  if (!skillName) {
    setStatus("Choose a skill first.", "error");
    return;
  }
  const modifier2 = options.modifier ?? (Number(getActionFieldValue('[data-roll-skill-field="modifier"]')) || 0);
  const odyssey = getOdysseyData(token);
  if (!options.raw && !Object.hasOwn(odyssey.skills, skillName)) throw new Error("Skill no longer exists.");
  const result = rollSkillCheck(options.raw ? 0 : odyssey.skills[skillName] ?? 0, modifier2);
  const summary = `${getResolvedCheckResultIcon(result.result)} Skill ${skillName}: ${result.totalPrimary} vs ${result.totalSecondary} (${result.result})`;
  await updateTrackerData2(token.id, (current2) => {
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
    `${getResolvedCheckResultIcon(result.result)} ${getCharacterName(token)} checks skill`,
    formatRollSkillDebug({
      tokenName: getCharacterName(token),
      skillName,
      result
    }),
    isResolvedCheckResultSuccess(result.result) ? "success" : result.result === "Critical Failure" ? "error" : "info"
  );
  setStatus(summary, isResolvedCheckResultSuccess(result.result) ? "success" : "error");
}
async function performRollChar(options) {
  if (profileRollPending) return;
  profileRollPending = true;
  try {
    await performRollCharOnce(options);
  } finally {
    profileRollPending = false;
  }
}
async function performRollSkill(options) {
  if (profileRollPending) return;
  profileRollPending = true;
  try {
    await performRollSkillOnce(options);
  } finally {
    profileRollPending = false;
  }
}
async function performPrivateGmRoll() {
  if (!isEditable()) {
    setStatus("Only the GM can use private rolls.", "error");
    return;
  }
  const dice = Number(getActionFieldValue('[data-gm-roll-field="dice"]')) || 20;
  const count = Number(getActionFieldValue('[data-gm-roll-field="count"]')) || 1;
  const modifier2 = Number(getActionFieldValue('[data-gm-roll-field="modifier"]')) || 0;
  const result = rollDice(dice, modifier2, count);
  const diceLabel = `${result.count}d${result.sides}`;
  const summary = buildDiceRollSummary(diceLabel, result);
  pushPrivateGmEntry(
    `GM private ${diceLabel}`,
    formatDiceDebug({
      tokenName: "GM private roll",
      result
    })
  );
  scheduleRender();
  setStatus(`Private GM roll. ${summary}`, "success");
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
      if (!target.isConnected || !target.dataset.sectionKey) return;
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
    if (actionNode.disabled) return;
    if (action === "close-character-hud") {
      void lib_default.popover.close("com.codex.body-hp/character-hud").catch((error) => setStatus(error.message, "error"));
      return;
    }
    if (action === "open-character-hud") {
      const id = "com.codex.body-hp/character-hud";
      void lib_default.popover.close(id).catch(() => {
      }).then(() => lib_default.popover.open({
        id,
        url: new URL("./index.html?hud=1", window.location.href).href,
        width: 920,
        height: 420,
        anchorReference: "POSITION",
        anchorPosition: { left: 20, top: 80 },
        anchorOrigin: { horizontal: "LEFT", vertical: "TOP" },
        transformOrigin: { horizontal: "LEFT", vertical: "TOP" },
        disableClickAway: true
      })).catch((error) => setStatus(error.message, "error"));
      return;
    }
    if (action === "profile-roll-char" || action === "profile-roll-skill" || action === "profile-random-skill") {
      const kind = action === "profile-roll-char" ? "attribute" : "skill";
      const modifier2 = clamp(Number(ui.selectedTokenPanel.querySelector(`[data-profile-modifier="${kind}"]`)?.value) || 0, -999, 999);
      actionNode.disabled = true;
      const operation = kind === "attribute" ? performRollChar({ attribute: actionNode.dataset.attribute, modifier: modifier2 }) : performRollSkill({ skill: actionNode.dataset.skill, raw: action === "profile-random-skill", modifier: modifier2 });
      void operation.catch((error) => setStatus(error.message, "error")).finally(() => {
        if (actionNode.isConnected) actionNode.disabled = false;
      });
      return;
    }
    if (action === "open-ammo-add") {
      ammunitionAddOpen = true;
      scheduleRender();
      return;
    }
    if (action === "close-ammo-add") {
      ammunitionAddOpen = false;
      scheduleRender();
      return;
    }
    if (action === "open-weapon-add") {
      weaponAddOpen = true;
      scheduleRender();
      return;
    }
    if (action === "close-weapon-add") {
      weaponAddOpen = false;
      scheduleRender();
      return;
    }
    if (action === "open-body-part-add") {
      bodyPartAddOpen = true;
      scheduleRender();
      return;
    }
    if (action === "close-body-part-add") {
      bodyPartAddOpen = false;
      scheduleRender();
      return;
    }
    if (action === "open-skill-add") {
      skillAddOpen = true;
      scheduleRender();
      return;
    }
    if (action === "close-skill-add") {
      skillAddOpen = false;
      scheduleRender();
      return;
    }
    if (action === "open-skill-perk") {
      skillPerkAddSkill = actionNode.dataset.skill ?? "";
      scheduleRender();
      return;
    }
    if (action === "cancel-skill-perk") {
      skillPerkAddSkill = "";
      scheduleRender();
      return;
    }
    if (action === "save-skill-perk") {
      const row = actionNode.closest(".skill-row");
      const field2 = row?.querySelector('[data-skill-perk-field="name"]');
      void addOdysseySkillPerk(actionNode.dataset.skill ?? "", field2?.value ?? "").then(() => {
        skillPerkAddSkill = "";
        scheduleRender();
      }).catch((error) => setStatus(error.message, "error"));
      return;
    }
    if (action === "remove-skill-perk") {
      void removeOdysseySkillPerk(actionNode.dataset.skill ?? "", Number(actionNode.dataset.perkIndex ?? -1)).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (action === "weapon-save") {
      void saveWeapon(actionNode).then(() => {
        if (actionNode.dataset.weaponType == null) {
          weaponAddOpen = false;
          scheduleRender();
        }
      }).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (action === "ammo-save") {
      void saveAmmunition(actionNode).then(() => {
        if (!actionNode.dataset.ammoId) {
          ammunitionAddOpen = false;
          scheduleRender();
        }
      }).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (action === "weapon-remove") {
      const type = actionNode.dataset.weaponType;
      const weaponIndex = Number(actionNode.dataset.weaponIndex ?? -1);
      if (weaponIndex < 0) return;
      void removeWeapon(type, weaponIndex).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (["ammo-remove", "ammo-reload", "ammo-unload"].includes(action)) {
      void editAmmunition(action, actionNode).catch((error) => {
        setStatus(error.message, "error");
        scheduleRender();
      });
      return;
    }
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
    if (action === "reload-token-visuals") {
      void reloadTokenVisuals().catch((error) => {
        setStatus(error?.message ?? "Unable to reload token visuals.", "error");
      });
      return;
    }
    if (action === "heal-limbs") {
      void healLimbs().catch((error) => {
        setStatus(error?.message ?? "Unable to heal limbs.", "error");
      });
      return;
    }
    if (action === "add-body-part") {
      void addBodyPart().then(() => {
        bodyPartAddOpen = false;
        scheduleRender();
      }).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (action === "remove-body-part") {
      void removeBodyPart(actionNode.dataset.part ?? "").catch((error) => {
        setStatus(error.message, "error");
      });
      return;
    }
    if (action === "toggle-part-hidden" && actionNode instanceof HTMLInputElement) {
      void toggleBodyPartHidden(actionNode.dataset.part ?? "", !actionNode.checked).catch((error) => {
        setStatus(error.message, "error");
      });
      return;
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
      return;
    }
    if (action === "perform-manual-attack") {
      void performAttack({ manualDefense: true }).catch((error) => {
        setStatus(error?.message ?? "Unable to resolve no target attack.", "error");
      });
      return;
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
      void addOdysseySkill().then(() => {
        skillAddOpen = false;
        scheduleRender();
      }).catch((error) => {
        setStatus(error?.message ?? "Unable to add skill.", "error");
      });
      return;
    }
    if (action === "add-weapon") {
      void addWeapon().catch((error) => {
        setStatus(error?.message ?? "Unable to add weapon.", "error");
      });
      return;
    }
    if (action === "export-token-data") {
      void exportSelectedTokenData().catch((error) => {
        setStatus(error?.message ?? "Unable to export token data.", "error");
      });
      return;
    }
    if (action === "open-token-import") {
      openTokenImportPicker();
      return;
    }
    if (action === "remove-skill" && skill) {
      void removeOdysseySkill(skill).catch((error) => {
        setStatus(error?.message ?? "Unable to remove skill.", "error");
      });
      return;
    }
    if (action === "remove-weapon") {
      const weaponIndex = Number(actionNode.dataset.weaponIndex ?? -1);
      if (weaponIndex < 0) return;
      void removeWeapon(weaponIndex).catch((error) => {
        setStatus(error?.message ?? "Unable to remove weapon.", "error");
      });
    }
  });
  document.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;
    if (target.dataset.attackField || target.dataset.manualAttackField || target.dataset.hudAttackField) queueMicrotask(paintRangeAdjustments);
    if (target.dataset.hudAttackField) {
      const field2 = target.dataset.hudAttackField;
      const original = ui.selectedTokenPanel.querySelector(`[data-attack-field="${field2}"]`);
      if (original) {
        original.value = target.value;
        original.dispatchEvent(new Event("change", { bubbles: true }));
      }
      scheduleRender();
      return;
    }
    if (target.dataset.weaponDraft === "type") {
      toggleWeaponDraftFields(target);
      return;
    }
    if (target.dataset.rangedEdit) {
      void editAmmunition("ranged-edit", target).catch((error) => {
        setStatus(error.message, "error");
        scheduleRender();
      });
      return;
    }
    if (target.dataset.attackField && activeTokenId) {
      if (target.dataset.attackField === "weaponName") {
        const token = getCharacterById(activeTokenId);
        const selectedWeapon = token ? getWeaponByName(token, target.value) : null;
        if (selectedWeapon) {
          syncAttackWeaponInputs(
            activeTokenId,
            selectedWeapon.selectionKey ?? selectedWeapon.name,
            selectedWeapon.damage,
            selectedWeapon.accuracy ?? 0
          );
        }
        return;
      }
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
      if (target.dataset.attackField === "parryMode") void updateTrackerData2(activeTokenId, (current2) => {
        const next = structuredClone(current2);
        next.odyssey.preferences.parryMode = target.value;
        return next;
      }).catch((e) => setStatus(e.message, "error"));
      if (target.dataset.attackField === "targetTokenId") {
        const selectedTarget = target.value ? getCharacterById(target.value) : null;
        saveAttackDraftValue(
          activeTokenId,
          "targetTokenName",
          selectedTarget ? getCharacterName(selectedTarget) : ""
        );
        void persistAttackTargetToken(activeTokenId, target.value).catch((error) => {
          console.warn("[Body HP] Unable to persist attack target", error);
        });
        void syncTargetHighlight().catch((error) => {
          console.warn("[Body HP] Unable to sync target highlight", error);
        });
      }
    }
    if (target.dataset.manualAttackField && activeTokenId) {
      if (target.dataset.manualAttackField === "weaponName") {
        const token = getCharacterById(activeTokenId);
        const selectedWeapon = token ? getWeaponByName(token, target.value) : null;
        if (selectedWeapon) {
          syncAttackWeaponInputs(
            activeTokenId,
            selectedWeapon.selectionKey ?? selectedWeapon.name,
            selectedWeapon.damage,
            selectedWeapon.accuracy ?? 0
          );
        }
        return;
      }
      const draftField = getSharedAttackDraftField(target.dataset.manualAttackField);
      if (draftField) {
        saveAttackDraftValue(activeTokenId, draftField, target.value);
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
    if (target.dataset.action === "rename-odyssey-skill") {
      const skill = target.dataset.skill;
      if (!skill) return;
      void renameOdysseySkill(skill, target.value).catch((error) => {
        setStatus(error.message, "error");
        scheduleRender();
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
    if (target.dataset.action === "set-weapon-accuracy") {
      const index = Number(target.dataset.weaponIndex ?? 0);
      void setWeaponAccuracy(index, target.value).catch((error) => {
        setStatus(error?.message ?? "Unable to save weapon accuracy.", "error");
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
    if (target.dataset.tokenImportInput != null) {
      const [file] = target.files ?? [];
      target.value = "";
      void importSelectedTokenData(file).catch((error) => {
        setStatus(error?.message ?? "Unable to import token data.", "error");
      });
      return;
    }
    if (target.dataset.action === "rename-body-part") {
      const partName2 = target.dataset.part;
      if (!partName2) return;
      void renameBodyPart(partName2, target.value).catch((error) => {
        setStatus(error.message, "error");
        scheduleRender();
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
    if (target.dataset.attackField || target.dataset.manualAttackField || target.dataset.hudAttackField) queueMicrotask(paintRangeAdjustments);
    if (target.dataset.hudAttackField) {
      const field = target.dataset.hudAttackField;
      const original = ui.selectedTokenPanel.querySelector(`[data-attack-field="${field}"]`);
      if (original) original.value = target.value;
      saveAttackDraftValue(activeTokenId, field, target.value);
    }
    if (target.dataset.attackField) {
      saveAttackDraftValue(activeTokenId, target.dataset.attackField, target.value);
    }
    if (target.dataset.manualAttackField) {
      const draftField = getSharedAttackDraftField(target.dataset.manualAttackField);
      if (draftField) {
        saveAttackDraftValue(activeTokenId, draftField, target.value);
      }
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
    if (target.dataset.action === "set-weapon-accuracy") {
      queueInputAutosave({
        tokenId: activeTokenId,
        action: "set-weapon-accuracy",
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
  installCompactButtons();
  try {
    enableActionResize(document.querySelector("[data-action-resize]"), {
      storageKey: "com.codex.body-hp/action-size",
      defaultWidth: 560,
      defaultHeight: 860
    });
    characterTools = installCharacterTools({
      OBR: lib_default,
      sanitize: sanitizeTrackerData,
      tracker: getTrackerData,
      getTokens: () => characterListCache,
      useItem: useCharacterItem,
      getToken: (id) => getCharacterById(id ?? activeTokenId),
      getRole: () => playerRole,
      canControl: canUseToken,
      update: updateTrackerData2,
      refreshOverlay: ensureOverlayForToken,
      render: scheduleRender,
      status: setStatus,
      download: downloadTextFile,
      resetDraft: (id) => {
        attackFormDrafts.delete(id);
        for (const [key, timer] of inputAutosaveTimers) {
          if (key.startsWith(id + "|")) {
            clearTimeout(timer);
            inputAutosaveTimers.delete(key);
          }
        }
      }
    });
    targetDistance = createTargetDistance({ OBR: lib_default, buildRuler, onRead: paintTargetDistance, context: HUD_MODE ? "hud" : "main" });
    lib_default.scene.grid.onChange(() => {
      targetDistance.invalidate();
      refreshTargetDistance();
    });
    window.addEventListener("beforeunload", () => {
      void targetDistance.clear();
    });
    bindUiEvents();
    await removeAttackTargetContextMenu();
    await Promise.all([
      loadSharedDebugConsole(),
      lib_default.scene.local.getItems().then((items) => {
        setLocalSceneItems(items);
      })
    ]);
    await syncState(true);
    startSelectionPolling();
    setStatus(
      "Ready. Select a character token on the map to edit it here.",
      "info"
    );
    lib_default.scene.items.onChange((items) => {
      setSceneItems(items);
      scheduleRender();
      scheduleSceneSideEffects(items);
    });
    lib_default.scene.local.onChange((items) => {
      setLocalSceneItems(items);
    });
    lib_default.player.onChange((player) => {
      playerRole = player.role;
      playerId = player.id ?? playerId;
      playerName = player.name ?? playerName;
      playerColor = player.color ?? playerColor;
      void handleSelectionChange(player.selection ?? []).then(() => {
        scheduleRender();
        return syncLocalOwnedHiddenTokenViews(sceneItems);
      }).catch((error) => {
        console.warn("[Body HP] Player state sync failed", error);
        scheduleRender();
      });
    });
    lib_default.party.onChange((players) => {
      setPartyPlayersState(players);
      scheduleRender();
    });
    lib_default.broadcast.onMessage(DEBUG_BROADCAST_CHANNEL, (event) => {
      const payload = event?.data;
      if (!payload || typeof payload !== "object") return;
      if (payload.type === "debug-clear") {
        if (pendingLocalDebugClear) {
          pendingLocalDebugClear = false;
          return;
        }
        debugEntries = [];
        renderDebugConsole();
        return;
      }
      if (payload.type !== "debug-entry") return;
      if (pendingLocalDebugEntryIds.has(payload.entry?.id)) {
        pendingLocalDebugEntryIds.delete(payload.entry.id);
        return;
      }
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
