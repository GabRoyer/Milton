const deniedGlobalNames = [
  "setTimeout",
  "setInterval",
  "queueMicrotask",
  "requestAnimationFrame",
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "Worker",
  "SharedWorker",
  "importScripts",
  "indexedDB",
  "caches",
] as const;

const baselineGlobalProperties = new Set(Object.getOwnPropertyNames(globalThis));

/** Replaces ambient worker capabilities that generated workbook automation does not need. */
export function installDeniedGlobals(): void {
  for (const name of deniedGlobalNames) {
    Object.defineProperty(globalThis, name, {
      configurable: true,
      writable: true,
      value: () => {
        throw new Error(`Sandbox denied access to ${name}.`);
      },
    });
  }
}

/** Removes global properties added by one run before the worker is reused. */
export function resetGlobalProperties(): void {
  for (const name of Object.getOwnPropertyNames(globalThis)) {
    if (!baselineGlobalProperties.has(name)) {
      Reflect.deleteProperty(globalThis, name);
    }
  }

  installDeniedGlobals();
}

/** Converts log details into JSON-compatible data before crossing the worker boundary. */
export function cloneJsonSafe(value: unknown): unknown {
  if (value === undefined) {
    return undefined;
  }

  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return String(value);
  }
}
