# Runtime sandboxing

Status: draft
Document branch: `effort/runtime-sandboxing-phase-0`

This effort focuses strictly on sandboxing generated code execution in Milton's Office runtime. It does not try to solve model behavior, prompt policy, workbook-level permissioning, or server-side isolation.

## Current state

Generated TypeScript is compiled in a browser worker, but emitted JavaScript is evaluated on the taskpane thread through `unsafeEvaluateOfficeCode()`. The evaluated module receives an `ExcelRuntimeContext` containing the raw `Excel.RequestContext`, a `workbook` convenience alias, `sync()`, `log()`, and an optional `AbortSignal`.

The current prompt asks generated code not to import packages, touch browser globals, use the DOM, or make network requests. The compiler rejects imports, but the runtime still executes with access to the taskpane global object.

## OfficeJS transport findings

The Office.js implementation source is not maintained as a normal open-source package in the public `OfficeDev/office-js` repo. That repo says it is primarily for issues, and the Office CDN is the officially supported source for add-ins. The debug CDN bundles are readable enough to inspect the runtime shape.

The Excel debug bundles confirm that OfficeJS client objects are mostly proxy command builders:

- `ClientRequestBase.buildRequestMessageBodyAndRequestFlags()` serializes queued `Actions` and `ObjectPaths`.
- `ClientRequestContext.syncPrivate()` sends that serialized body through a request executor.
- local workbook requests go to the `http://document.localhost/_api/ProcessQuery` path.
- the default local-document sender calls `OSF.DDA.RichApi.executeRichApiRequestAsync(...)`.
- when a `_richApiNativeBridge` object exists, OfficeJS initializes `HostBridge`, installs `HttpUtility.setCustomSendLocalDocumentRequestFunc(...)`, and sends serialized bridge messages through `bridge.sendMessageToHost(JSON.stringify(message))`.
- `HostBridgeSession` supplies a custom request executor so `ClientRequestContext` can route local document requests over that bridge.

This means the worker sandbox should recreate the OfficeJS command-builder behavior rather than inventing a hand-authored workbook API. The main challenge is that the shipped OfficeJS host bundles are browser-window scripts: they reference `window`, `document`, script loading, storage, dialogs, telemetry, and host initialization surfaces. They should not be assumed to load cleanly inside a worker.

## Goals

- Prevent generated code from accessing the DOM, taskpane globals, local storage, cookies, network APIs, worker spawning, and other ambient browser capabilities unless explicitly exposed.
- Preserve a practical workbook automation API for generated code.
- Keep the taskpane responsive while generated code runs for a long time, loops forever, or schedules excessive async work.
- Return clear structured errors for sandbox denials, timeouts, aborts, compile failures, and OfficeJS failures.
- Keep the implementation browser-compatible for the Office add-in host.

## Non-goals

- Full adversarial JavaScript containment on the same thread.
- A general-purpose package/runtime loader for generated code.
- Server-side code execution.
- Trusting TypeScript types or prompts as a security boundary.
- Replacing OfficeJS with a small Milton-specific workbook API. Milton depends on models already knowing the OfficeJS object model.
- Preventing valid generated code from making destructive workbook edits after the host has explicitly exposed an edit capability.

## Threat model

Generated code should be treated as untrusted input. The primary risks are:

- Ambient taskpane access: `window`, `document`, DOM mutation, cookies, local storage, extension globals, and app state.
- Network and persistence access: `fetch`, `XMLHttpRequest`, `WebSocket`, `indexedDB`, caches, file APIs, and related worker-available APIs.
- Escape through JavaScript constructors or prototypes after simple global shadowing.
- Denial of service through infinite loops, promise recursion, microtask loops, many timers, or large allocations.
- Capability confusion where generated code can reach raw OfficeJS objects or callback functions that expose more than intended.

## Chosen approach

Run generated code only in a dedicated web worker. The worker owns an OfficeJS-compatible proxy runtime that creates `Actions` and `ObjectPaths`. The taskpane owns the actual Office host bridge.

Generated code should continue to look like normal OfficeJS:

```ts
await Excel.run(async (context) => {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:C10");
  range.load(["address", "values"]);
  await context.sync();
  return range.values;
});
```

The worker-owned `context.sync()` serializes a `ProcessQuery` payload and sends it to the taskpane. The taskpane validates envelope-level policy and budget, forwards the payload through the real OfficeJS local-document executor, then returns the serialized response. The worker applies returned action results to its proxy objects.

The worker should not receive raw OfficeJS objects. It should receive:

- the compiled JavaScript for one run
- a run id
- a declarative set of allowed capabilities
- serializable runtime options, including timeout and budget policy
- a host capability snapshot for requirement-set gating

The worker should expose OfficeJS-compatible proxy objects backed by a sandbox transport. Calls should be serialized to the taskpane host, validated, executed through real OfficeJS, and resolved back to the worker.

The taskpane should not try to reinterpret `Actions` and `ObjectPaths` into public OfficeJS calls. That would require reimplementing OfficeJS and Excel semantics in the taskpane. The cleaner boundary is to make the taskpane a narrow relay to the real `ProcessQuery` execution path.

## DoS prevention

DoS prevention is worth doing. In the current same-thread evaluator, a single generated `while (true) {}` can hang the taskpane and force the user to reload the add-in. In the worker design, runaway code should not hang the taskpane, but it can still monopolize a run slot, keep issuing host work, consume memory, or require explicit termination.

Use layered controls:

### Hard wall-clock timeout

The taskpane host starts a timer for each execution. The initial timeout should be 30 seconds. If the worker does not finish before the deadline, or if the user explicitly kills the run, the host calls `worker.terminate()`, rejects the run as timed out or killed, and creates a fresh worker for future runs.

This is the most important control because it handles non-cooperative infinite loops. The timeout should remain configurable later; long-running workbook scripts can have a longer deadline while still keeping the taskpane responsive and cancellable.

### Cooperative cancellation

Keep passing an `AbortSignal` concept through the public API, but treat it as a convenience rather than the primary safety mechanism. The worker should check cancellation before and after host RPC calls, and host-side OfficeJS request handling should stop accepting new requests after cancellation.

### Timer and microtask controls

Generated code should not need raw timers for workbook automation. The concern is not that timers can freeze the taskpane, since the code runs in a worker. The concern is that unrestricted timers let a script schedule unbounded background work, continue producing side effects after the main run appears done, or keep issuing host RPCs in small async chunks to avoid simple CPU-style budgets.

In the worker sandbox, shadow or replace:

- `setTimeout`
- `setInterval`
- `queueMicrotask`
- `requestAnimationFrame`
- string-eval timer forms

The first policy should deny them all. If a later workflow needs waits or polling, expose a metered `ctx.sleep(ms)` helper through Milton's runtime API with a max delay, max call count, and cancellation checks.

### Async operation budget

Host RPC requests should have per-run budgets:

- max OfficeJS operations
- max `sync()` calls
- max log entries and log payload bytes
- max returned JSON bytes
- max outstanding RPC requests

This prevents code from avoiding the CPU timeout by scheduling unbounded host work.

### Memory and payload limits

Browser workers do not give us precise portable heap limits. Use practical payload caps instead:

- max source length
- max emitted JavaScript length
- max message size per RPC
- max final result size
- terminate the worker on structured-clone failures or over-budget payloads

### Loop instrumentation

The compiler could eventually inject budget checks into loops, function entries, and awaited boundaries. This would improve error messages and stop some runaway code before worker termination.

Do not include loop instrumentation in the initial implementation. It is brittle around emitted JavaScript patterns and cannot replace worker termination.

## Generated command-builder requirements

The worker-side OfficeJS surface should be generated at build time, but the generic proxy machinery should be Milton-owned handwritten code. In practice, the generator should produce a table-driven command-builder layer rather than copying OfficeJS runtime implementation code.

The handwritten core should provide:

- `Excel.run(...)` and a worker-owned `RequestContext`
- root objects such as `context.workbook` and `context.application`
- `ClientObject`, `ClientResult`, collection, tracked-object, `load(...)`, and `sync()` primitives
- action and object-path factories
- request serialization into the OfficeJS `ProcessQuery` body shape
- response application back onto worker proxy objects
- requirement-set checks, sandbox errors, and budget accounting
- a request executor that relays only serialized host requests to the taskpane

```ts
class SandboxRequestExecutor {
  executeAsync(customData, requestFlags, requestMessage) {
    return postToTaskpane({
      type: "officejs-process-query",
      customData,
      requestFlags,
      requestMessage,
    });
  }
}
```

The generated layer should provide the Excel API classes, properties, methods, enums, and load-option types that generated scripts expect. For each generated member, the generator needs enough metadata to decide how a call mutates the pending request:

- the public namespace, class, member name, TypeScript signature, and return type
- the Office host dispatch name when it differs from the JavaScript name
- whether the member is a scalar loadable property, settable property, navigation property, collection indexer, method that returns a proxy object, method that queues an action, or method that returns a `ClientResult`
- argument serialization rules, including optional arguments and enum/string conversions
- return handling, including proxy object construction, collection item paths, nullable object patterns, and scalar result assignment
- requirement-set metadata such as `ExcelApi 1.N`
- preview, deprecated, platform-specific, or unsupported markers

The current `@types/office-js` input is useful for TypeScript signatures and requirement comments, but it is not sufficient by itself. It does not reliably encode every command-builder distinction needed to build exact `Actions` and `ObjectPaths`. For example, the generator can often infer that `worksheet.getRange(...)` returns a child object path and `range.clear(...)` queues an action, but special cases, host dispatch names, result-shaping behavior, collection semantics, and object-path roots need explicit metadata.

The `Id` fields inside `Actions` and `ObjectPaths` are request-local handles, not stable Office API member ids. OfficeJS allocates them monotonically from the request context while building a batch, then cross-references actions to object paths through those generated ids. Milton can reproduce that allocator directly.

Current OfficeJS Rich API payloads also do not appear to send stable per-method ids for members such as `Worksheet.getRange(...)`. At the JavaScript-to-host `ProcessQuery` boundary, the bundle builds method and property entries with string `Name` fields. For example, `Worksheet.getRange(...)` creates a method object path with `Name: "GetRange"`, and `ClientRequestBase.buildRequestMessageBodyAndRequestFlags()` serializes queued `Actions` and `ObjectPaths` into JSON. The outer Office bridge still has numeric dispatch ids, such as the dispatch id for `ExecuteRichApiRequestAsync`, and the native host may map names to an internal IDL/member table after it receives the request. That lower-level mapping is not something the worker command builder needs if it forwards the same `ProcessQuery` JSON shape that OfficeJS sends today.

The harder metadata is not those per-request ids. It is the mapping from public OfficeJS members to the right command shape:

- dispatch names such as `GetRange`, `Worksheets`, or `Values`
- action and object-path type constants
- operation type, invalidation, collection, and result-processing flags
- property original names and compressed loaded-property result names
- helper methods such as `getByIdMethodName` for reconstructed collection items
- requirement-set checks and platform gates

Some of this appears in public declarations and docs, but the `ActionType` and `ObjectPathType` numeric constants, exact compact command metadata, and flag choices appear to be generated into the OfficeJS host bundles rather than published as a supported standalone metadata artifact. Treat the bundles and captured payloads as parity oracles, not as code to copy into the worker runtime.

The most practical input stack is:

- parse `@types/office-js` for the public TypeScript surface Milton already compiles against
- parse Office API documentation metadata for API set, deprecation, preview, and member ownership details
- maintain a small override table for command-builder facts that declarations and docs do not express
- generate parity fixtures from real OfficeJS debug bundles for representative calls and compare normalized payloads

This should not require a full OfficeJS runtime clone. The generic runtime is small: allocate request-local ids, create object paths, queue actions, serialize `ProcessQuery`, and apply results. The generated API layer is mostly class/member scaffolding plus metadata descriptors. A CDN-derived generator can scrape the generated OfficeJS class/prototype code or compact metadata where available to recover host names, action/object-path choices, flags, and result handling. A types-only generator can produce a useful first pass for signatures and many obvious command shapes, but it will need overrides and parity tests before it is trustworthy for broad Excel coverage.

Concrete types-only gaps from the current Excel bundle include:

- `Worksheet.getRange(address?: string): Excel.Range` does not say the host name is `GetRange`, the returned object path is invalid after the request, the operation type is read, or the flags value is `4`.
- `Range.values: any[][]` does not say large 2D writes are split into row chunks through `getRow(...).getBoundingRect(...)` before enqueueing `SetProperty("Values", ...)`.
- `Range.getColumnsAfter(count?: number): Excel.Range` does not say older hosts polyfill the call by composing adjacent range methods instead of sending `GetColumnsAfter`.
- `WorksheetCollection.add(name?: string): Excel.Worksheet` and `WorksheetCollection.getActiveWorksheet(): Excel.Worksheet` have the same return type shape in declarations, but the bundle uses different operation types, invalidation behavior, and undo flags.

The generated output should be split into stable pieces:

- `excel-types.generated.ts`: generated public classes, enums, method signatures, and load-option types
- `excel-command-metadata.generated.ts`: member descriptors used by the generic action/object-path factories
- `excel-requirements.generated.ts`: API set requirements for runtime gating
- `excel-overrides.ts`: reviewed handwritten metadata for special cases and unsupported members

### Source and package strategy

Do not make the normal application build download Office CDN bundles or rewrite generated files. That would make builds network-dependent, hard to review, and vulnerable to silent upstream changes.

Use a dedicated workspace package for the runtime and generated proxy surface:

- package: `@milton/office-js-proxy`
- owns: worker-safe proxy runtime, generated Excel proxy classes, generated command metadata, generated requirement metadata, and reviewed overrides
- consumed by: `@milton/office-runtime`
- normal `build`: typecheck and bundle committed sources only
- normal tests: run unit, response-replay, and payload parity fixtures against committed snapshots

Use a separate generator entrypoint, either in `packages/office-js-proxy-codegen` or `packages/office-js-proxy/scripts`, that is run intentionally:

- `generate:sources`: download pinned OfficeJS debug bundle inputs into a local cache, verify URL and SHA-256 from a checked-in source manifest, and never accept floating `latest` inputs
- `generate:metadata`: extract normalized command metadata from the pinned bundle plus `@types/office-js` and docs metadata
- `generate`: emit the proxy files into `@milton/office-js-proxy/src/generated`
- `generate:check`: regenerate into a temp directory and fail if committed generated files differ

Commit the generated proxy files and the small normalized metadata snapshots needed for review. Do not commit the full CDN bundle unless legal review explicitly approves it; keep the source manifest with exact URLs, hashes, capture date, and OfficeJS bundle names so regeneration remains auditable. Upgrading OfficeJS then becomes a deliberate PR: update the manifest, regenerate, inspect generated diffs, update overrides, and refresh parity fixtures.

The taskpane side receives worker executor messages, checks run id and budgets, forwards the serialized request through the real OfficeJS executor, and returns the serialized response. The worker applies the response to its proxy objects using OfficeJS-compatible result handling.

The taskpane should not try to reinterpret `Actions` and `ObjectPaths` into public OfficeJS calls. That would require reimplementing OfficeJS and Excel semantics in the taskpane. The cleaner boundary is:

- worker owns the OfficeJS-compatible proxy objects and creates serialized `ProcessQuery` payloads
- taskpane validates envelope-level policy and budget
- taskpane forwards the payload through the real OfficeJS local-document executor
- worker receives the serialized response and applies action results

This still relies on a private or semi-private taskpane execution surface, such as the local-document `ProcessQuery` sender used by OfficeJS, but it confines that risk to one small host adapter instead of using private OfficeJS internals throughout generated code execution.

The generator is only credible if it has parity tests. For each supported API pattern, tests should run equivalent code through real OfficeJS, capture the serialized `ProcessQuery` request, run the generated worker proxy, and compare normalized `Actions` and `ObjectPaths`. Response replay tests should then feed captured OfficeJS responses back into the worker runtime and verify loaded properties, `ClientResult` values, thrown errors, and null-object behavior.

## API availability and requirement sets

Different Excel clients support different OfficeJS requirement sets. The generated worker proxy layer should include requirement metadata for each generated class member and method.

At execution start, the taskpane should capture the host's supported requirement sets through `Office.context.requirements.isSetSupported(...)` checks and send a capability snapshot to the worker. The worker should use that snapshot to gate generated APIs before they enqueue actions:

- if a method or property requires `ExcelApi 1.N` and the host does not support it, throw a clear `ApiNotSupported`-style error at the call site
- if a member is setless or has no reliable requirement metadata, use a conservative allowlist or an explicit runtime probe
- include the required set/version and detected host support in structured error details
- keep the full OfficeJS-compatible surface in TypeScript so models can write familiar code, but fail fast when the current host cannot execute it

The taskpane should still treat host errors as authoritative. Even with worker-side gating, the real OfficeJS executor can reject requests for platform-specific, preview, or partially rolled-out APIs. Those errors should be relayed back without masking them.

Build-time generation should preserve requirement metadata from the agreed source inputs: pinned OfficeJS debug bundles, `@types/office-js`, Office API docs metadata, and Office requirement-set metadata. Parity tests should cover both supported and unsupported API members.

## Phased implementation plan

### Phase 1: Worker execution harness

- Branch: `effort/runtime-sandboxing-phase-1`
- Target: `main`, after the design doc is accepted.
- PR scope: introduce a reusable execution worker for sequential runs, run compiled JavaScript inside it, add host/worker request-response plumbing, hard timeout termination, user kill, cancellation propagation, structured error mapping, and worker recreation after termination.
- Dependencies: accepted design doc.
- Validation: browser/Vitest worker tests for successful execution, thrown errors, timeout termination for infinite loops, user kill, cancellation, worker reset between sequential runs, worker recreation after termination, and no taskpane-thread blocking for runaway code.

The first version can support a minimal runtime API such as `log()` and pure return values before workbook capabilities are added.

### Phase 2: Generated command-builder foundation

- Branch: `effort/runtime-sandboxing-phase-2`
- Target: `effort/runtime-sandboxing-phase-1`.
- PR scope: add the build-time generator, generated Excel API surface, generated command metadata, generated requirement metadata, handwritten proxy primitives, and mocked `ProcessQuery` serialization tests for core workbook, worksheet, range, collection, load, set, and `ClientResult` patterns.
- Dependencies: Phase 1.
- Validation: unit tests for generated signatures, descriptor coverage, request serialization, response replay, unsupported API gating, and payload budget accounting.

### Phase 3: Taskpane host adapter

- Branch: `effort/runtime-sandboxing-phase-3`
- Target: `effort/runtime-sandboxing-phase-2`.
- PR scope: add the taskpane adapter that validates worker messages and forwards serialized `ProcessQuery` payloads through the real OfficeJS local-document executor.
- Dependencies: Phase 2.
- Validation: Office-host validation for `Excel.run`, range load, range write, `context.sync()`, OfficeJS errors, timeout behavior during pending sync, taskpane cancellation, and rejection when unsupported ambient APIs are used.

This phase is the first end-to-end workbook execution path.

### Phase 4: Coverage, budgets, and UI feedback

- Branch: `effort/runtime-sandboxing-phase-4`
- Target: `effort/runtime-sandboxing-phase-3`.
- PR scope: expand generated API coverage, add parity fixtures for additional Excel patterns, configure execution limits, add structured denial/timeout messages, log/result truncation, taskpane-visible timeout state, and update the model-facing Office code tool instructions to mention sandbox restrictions such as no imports, no DOM/network access, denied timers, timeout, kill, and supported OfficeJS-compatible APIs.
- Dependencies: Phase 3.
- Validation: tests for each budget, parity fixtures for new API patterns, user-facing error text, transcript details, and prompt/tool description alignment.

## Decisions and remaining details

- Use reusable workers, but never run two scripts concurrently inside the same worker. A reusable worker is a run slot. If Milton needs parallel script execution, allocate one worker per active run from a small worker pool. Reuse only idle workers that completed cleanly; terminate and replace workers after timeout, user kill, failed reset, or suspicious protocol errors.
- Generator inputs are pinned OfficeJS debug CDN bundles, `@types/office-js`, Office API docs metadata from `OfficeDev/office-js-docs-reference`, and requirement-set metadata from `OfficeDev/Office-Js-Requirement-Sets`. The remaining implementation detail is the exact source manifest format, pinned versions, and precedence rules when sources disagree.
- Generated metadata should include facts that can be mechanically recovered and reviewed in diffs: public member names, host dispatch names, action/object-path category, operation type, flags, return proxy type, load/result field names, requirement sets, and platform gates. Handwritten overrides should be reserved for behavior that is not a simple descriptor: compatibility polyfills, bulk-write splitting, custom validation, known source conflicts, and intentionally unsupported members.
- The taskpane host adapter should be a narrow `ProcessQuery` relay that calls the same local-document Rich API path OfficeJS uses today. The preferred implementation is a small adapter over `OSF.DDA.RichApi.executeRichApiRequestAsync(...)` using the same safe-array request shape as OfficeJS. If a stable OfficeExtension local-document request helper is accessible in the loaded bundle, it can be wrapped behind the same adapter. The taskpane must not reinterpret the payload into public OfficeJS calls.
- Start with a 30-second hard timeout per script plus an explicit user kill path. Treat CPU/runtime timeout and pending Office host-call timeout as the same deadline at first; split them later only if real workbook workflows need longer host calls.
- Deny raw timer APIs initially. Add a metered `ctx.sleep(ms)` only when a concrete workflow needs waits or polling.
- Do not include loop instrumentation initially. Worker termination is the primary non-cooperative stop mechanism.

## Alternatives considered

### Main-thread shadowing

Lexical shadowing and frozen intrinsics can reduce accidental DOM/network use, but same-thread JavaScript is not a sandbox. Infinite loops still freeze the taskpane before timeout or cancellation logic can run.

### Runtime monkey-patching OfficeJS

Patching OfficeJS internals after load avoids shipping a modified file, but it still depends on private implementation names and may still count as modifying or working around the shipped library. It is useful for spikes, not as the design foundation.

### Loading OfficeJS directly in the worker

The OfficeJS bundle already has bridge-shaped paths, but the shipped host bundles are browser-window scripts that reference `window`, `document`, script loading, storage, dialogs, telemetry, and host initialization. Making that worker-safe would be fragile and still depend on private hooks such as `_richApiNativeBridge`.

### Sandboxed iframe

An iframe can isolate DOM access, but it is not a reliable long-running or infinite-loop boundary across browsers and Office WebView hosts because it may share an event loop with the taskpane.

### Milton-specific workbook API

A small custom API would be easier to secure, but it discards the main advantage that models already know OfficeJS. Milton should preserve OfficeJS names and call shapes.

### Third-party JS isolates

SES, QuickJS/WASM, or an interpreter could add language-level controls, but they add runtime cost and still need an OfficeJS-compatible command-builder layer plus taskpane host bridge.

## Reference notes

- Web workers have their own global scope and event loop, and do not expose the normal `Window` object: https://developer.mozilla.org/en-US/docs/Web/API/WorkerGlobalScope
- Workers can be terminated immediately from the owning thread: https://developer.mozilla.org/en-US/docs/Web/API/Worker/terminate
- Workers still expose many APIs, including `fetch`, timers, storage-related APIs, `WebSocket`, and the ability to spawn workers in supporting browsers, so the worker global still needs capability hardening: https://developer.mozilla.org/en-US/docs/Web/API/Web_Workers_API/Functions_and_classes_available_to_workers
- CSP `worker-src` can restrict where worker scripts load from, but it does not replace runtime capability controls: https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Content-Security-Policy/worker-src
- The public `OfficeDev/office-js` repo says the CDN is the official supported Office.js source and the repo is primarily for issues: https://github.com/OfficeDev/office-js
- Office API reference source and generated docs metadata live in `OfficeDev/office-js-docs-reference`: https://github.com/OfficeDev/office-js-docs-reference
- Office requirement-set metadata is published in `OfficeDev/Office-Js-Requirement-Sets`: https://github.com/OfficeDev/Office-Js-Requirement-Sets
- Inspected CDN debug bundles: https://appsforoffice.microsoft.com/lib/1/hosted/excel-web-16.00.debug.js and https://appsforoffice.microsoft.com/lib/1/hosted/excel-win32-16.01.debug.js
