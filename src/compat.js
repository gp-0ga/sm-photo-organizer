// 古いiPhone/Androidのブラウザに無い機能を補う（既にある場合は何もしない）。

function define(target, name, value) {
  if (target && typeof target[name] !== "function") {
    Object.defineProperty(target, name, { value, configurable: true, writable: true });
  }
}

define(Blob.prototype, "arrayBuffer", function arrayBuffer() {
  const blob = this;
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
});

define(Blob.prototype, "text", function text() {
  const blob = this;
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(blob);
  });
});

define(Element.prototype, "replaceChildren", function replaceChildren(...nodes) {
  while (this.firstChild) this.removeChild(this.firstChild);
  nodes.forEach((node) => this.appendChild(typeof node === "string" ? document.createTextNode(node) : node));
});

define(Promise, "withResolvers", function withResolvers() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
});

define(Object, "hasOwn", (object, key) => Object.prototype.hasOwnProperty.call(object, key));

for (const type of [Array, Uint8Array, Int8Array, Uint16Array, Int16Array, Uint32Array, Int32Array, Float32Array, Float64Array, String]) {
  define(type.prototype, "at", function at(index) {
    const length = this.length;
    const position = index < 0 ? length + index : index;
    return position < 0 || position >= length ? undefined : this[position];
  });
}

for (const type of [Array, Uint8Array]) {
  define(type.prototype, "findLast", function findLast(callback, thisArg) {
    for (let index = this.length - 1; index >= 0; index -= 1) {
      if (callback.call(thisArg, this[index], index, this)) return this[index];
    }
    return undefined;
  });
  define(type.prototype, "findLastIndex", function findLastIndex(callback, thisArg) {
    for (let index = this.length - 1; index >= 0; index -= 1) {
      if (callback.call(thisArg, this[index], index, this)) return index;
    }
    return -1;
  });
}

define(String.prototype, "replaceAll", function replaceAll(search, replacement) {
  if (search instanceof RegExp) {
    if (!search.global) throw new TypeError("replaceAll must be called with a global RegExp");
    return this.replace(search, replacement);
  }
  return this.split(String(search)).join(replacement);
});

define(Array.prototype, "flat", function flat(depth = 1) {
  const result = [];
  const walk = (list, level) => list.forEach((item) => (Array.isArray(item) && level > 0 ? walk(item, level - 1) : result.push(item)));
  walk(this, depth);
  return result;
});

define(Array.prototype, "flatMap", function flatMap(callback, thisArg) {
  return this.map(callback, thisArg).flat(1);
});

define(globalThis, "structuredClone", (value) => JSON.parse(JSON.stringify(value)));
