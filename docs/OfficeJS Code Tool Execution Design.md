# OfficeJS Code Tool Execution Design

## Overview

This document scopes the first implementation milestone for Milton's OfficeJS code execution tool.

The goal is to let the agent call one tool with TypeScript source, typecheck and compile that source to JavaScript in the Office taskpane runtime, and execute it against a controlled OfficeJS runtime context. This milestone intentionally does not attempt to solve sandboxing, persisted macro libraries, compilation caching, or document-packaged artifacts.

The general architecture document describes the long-term pipeline as:

```text
TS source
  -> parse
  -> static analysis
  -> AST transforms
  -> capability validation
  -> instrumentation
  -> compilation
  -> sandboxed execution
```

This milestone implements the narrow vertical slice:

```text
agent tool call
  -> TS source
  -> typecheck and compile to JS in a web worker
  -> evaluate module
  -> call run(ctx)
  -> return structured result
```

## Goals

- Add a single agent tool for OfficeJS code execution.
- Accept TypeScript source from the model.
- Typecheck and compile TypeScript in a browser worker.
- Evaluate the compiled JavaScript and call an exported `run(ctx)` function.
- Provide a small typed runtime context for Excel-first workbook operations.
- Return text plus structured details suitable for transcript replay and UI display.
- Keep execution sequential so workbook mutations happen in a predictable order.
- Make the unsafe nature of this milestone explicit in code, docs, and prompts.

## Non-Goals

- Sandboxing or security isolation.
- Security or capability-policy static analysis.
- AST instrumentation.
- Compilation artifact caching.
- Durable macro storage.
- External package imports inside generated code.
- Multi-host support beyond Excel.
- Third-party package type resolution for model-generated programs.

Sandboxing is critical, but it is intentionally deferred. This milestone should be treated as a functional prototype that establishes the agent/runtime/tool boundary, not as a safe execution boundary.

## Current Context

The taskpane currently runs the Pi agent loop from `packages/ui/src/ExcelTaskpaneApp.tsx` with no tools:

```ts
tools: []
```

The repo already has:

- `apps/office`: the Vite Office add-in app.
- `packages/ui`: shared React taskpane UI.
- `packages/office-host`: thin typed Office host integrations.
- `packages/pi-agent-core`: agent loop and `AgentTool` execution support.
- `packages/pi-ai`: model abstraction, TypeBox exports, and tool schemas.

`AgentTool` already supports:

- TypeBox parameter validation.
- partial tool updates.
- sequential or parallel execution mode.
- structured result details.
- error conversion into tool results.

This makes the OfficeJS executor a natural Pi agent tool rather than a separate orchestration loop.

## Proposed Architecture

Add a new workspace package:

```text
packages/office-runtime
```

This package owns the browser-compatible code execution runtime:

```text
@milton/office-runtime
  - executeOfficeCode(source, options)
  - compileOfficeCode(source, options)
  - createOfficeCodeCompilerWorkerClient(options)
  - createExcelRuntimeContext(options)
  - createExecuteOfficeJsCodeTool(options)
```

The package should keep implementation areas separated by ownership:

```text
packages/office-runtime/src/
  execution.ts        # core compile/evaluate/Excel.run orchestration
  compiler/
    compile.ts        # TypeScript virtual compiler host and compile result types
    raw-modules.d.ts  # ambient typing for Vite raw declaration imports
    worker.ts         # web worker entrypoint
    worker-client.ts  # lazy worker client and request tracking
    built-in-types/   # Milton declarations injected into generated code typechecking
      milton-runtime.d.ts
  runtime/
    context.ts        # ExcelRuntimeContext and ctx construction
  evaluation/
    unsafe-evaluator.ts
  index.ts            # public package export surface
```

`packages/office-host` should remain focused on typed Office host helpers. The runtime package can depend on it later, but the execution pipeline is a separate concern.

The initial runtime flow:

```text
Model
  -> execute_officejs_code({ code })
  -> AgentTool.execute()
  -> typecheck and compile TypeScript source in compiler worker
  -> evaluate generated JS
  -> validate module exports run(ctx)
  -> Excel.run(async context => ...)
  -> return tool result
```

## Tool Contract

Tool name:

```text
execute_officejs_code
```

Description:

```text
Compile and run TypeScript code against the current Excel workbook using the Excel OfficeJS API.
```

Parameters:

```ts
Type.Object({
  code: Type.String({
    description:
      "TypeScript source that uses the Excel OfficeJS API through ctx. It must export or return an async function named run(ctx: ExcelRuntimeContext). Return a JSON-serializable object from run(ctx) for any workbook data or results the model should see.",
  }),
})
```

The tool should use `executionMode: "sequential"` because OfficeJS workbook mutation and `Excel.run` batching should not run concurrently until the runtime has explicit coordination.

## Program Shape

The first supported program shape should be:

```ts
export async function run(ctx: ExcelRuntimeContext) {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.values = [["Hello from Milton"]];
  await ctx.sync();

  return {
    message: "Wrote A1.",
    address: "A1",
    value: "Hello from Milton",
  };
}
```

For the milestone, generated programs should not import modules. The compiler can allow TypeScript syntax, but module resolution should not be supported.

Generated code should return a JSON-serializable object when it needs to report workbook data back to the model. This is especially important for read operations:

```ts
export async function run(ctx: ExcelRuntimeContext) {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:C10");
  range.load(["address", "values"]);
  await ctx.sync();

  return {
    message: "Read A1:C10.",
    address: range.address,
    values: range.values,
  };
}
```

The runtime should normalize exports so the model has one obvious target:

- Preferred: `export async function run(ctx)`.
- Acceptable fallback: `async function run(ctx)` with a returned object from the evaluator wrapper.

If `run` is missing or not a function, the tool should fail with a clear tool error.

## Runtime Context

The initial context should be intentionally small:

```ts
export interface ExcelRuntimeContext {
  context: Excel.RequestContext;
  workbook: Excel.Workbook;
  sync(): Promise<void>;
  log(message: string, details?: unknown): void;
  signal?: AbortSignal;
}
```

The runtime executes user code inside:

```ts
await Excel.run(async (context) => {
  const runtimeContext = createExcelRuntimeContext(context, {
    signal,
    onLog: (entry) => logs.push(entry),
  });
  result = await run(runtimeContext);
});
```

This preserves OfficeJS batching semantics and keeps the model's mental model close to normal OfficeJS code.

`ctx.context` exposes the raw `Excel.RequestContext`. `ctx.workbook` and `ctx.sync()` are convenience aliases for the common path. The raw context is useful because OfficeJS APIs are scoped to the request context, and hiding it would force Milton to recreate OfficeJS surface area too early.

The context can grow later with higher-level helpers, but the first milestone should avoid inventing a large abstraction over OfficeJS.

## Worker Strategy

Compilation and typechecking should happen in a web worker so large generated programs or TypeScript initialization do not freeze the taskpane UI.

OfficeJS execution still happens on the taskpane main thread:

```text
main thread
  -> send TS source to compiler worker
worker
  -> typecheck and emit JS
main thread
  -> unsafe evaluate JS
  -> call run(ctx) inside Excel.run(...)
```

The worker should be lazy-loaded on the first `execute_officejs_code` call. That keeps the normal taskpane startup path from paying the TypeScript compiler cost before workbook automation is used.

The default worker client should spawn the worker with the bundler-native module-worker pattern:

```ts
new Worker(new URL("./worker.ts", import.meta.url), {
  type: "module",
});
```

That lets Vite discover the worker entrypoint, build it as a separate chunk, and keep the TypeScript compiler plus raw `@types/office-js` declaration text out of the normal taskpane startup bundle. Non-browser tests can inject a compiler or fall back to a dynamic direct compiler import.

The worker is a responsiveness boundary, not a security boundary. It does not replace sandboxing.

## Compilation Strategy

The add-in should package the TypeScript compiler as a runtime dependency of `@milton/office-runtime`.

Initial compiler behavior:

- Use a TypeScript `Program` through a virtual compiler host in the compiler worker.
- Emit modern JavaScript compatible with the Office WebView targets already used by the app.
- Typecheck generated code against TypeScript's packaged standard library declarations, Milton runtime declarations, and the installed `@types/office-js` declarations.
- Report syntactic and semantic diagnostics.
- Strip TypeScript types and emit evaluator-compatible CommonJS-style output.
- Reject import declarations before evaluation with a clear unsupported-imports diagnostic.
- Do not execute emitted JavaScript when there are TypeScript error diagnostics.

Recommended initial compiler options:

```ts
{
  target: ts.ScriptTarget.ES2020,
  module: ts.ModuleKind.CommonJS,
  jsx: ts.JsxEmit.None,
  noEmitOnError: true,
  strict: true,
  noImplicitAny: false,
  skipLibCheck: true,
}
```

`noImplicitAny` should be disabled for generated workbook code because models
often omit helper parameter annotations. This keeps the tool focused on
actionable OfficeJS and runtime-shape diagnostics without rejecting otherwise
valid generated code solely for missing annotations.

The virtual compiler host should provide at least:

- the generated source as an in-memory entry file,
- the standard library declaration files from the installed `typescript` package,
- `ExcelRuntimeContext` declarations,
- the installed `@types/office-js` declaration file instead of a hand-maintained Excel API subset,
- a no-op module resolver that rejects model-authored imports.

Typechecking is worth including in the milestone because it gives the model actionable feedback about nonexistent OfficeJS APIs, wrong property names, and mismatched value shapes before any workbook mutation runs.

Bundling the TypeScript compiler will add weight. The worker should be split into its own lazy-loaded bundle so the normal chat path and taskpane startup remain lighter.

## Evaluation Strategy

Because sandboxing is out of scope, the first evaluator can use a direct browser JavaScript evaluator. It must be named and documented as unsafe.

The runtime should still avoid exposing browser globals intentionally through the API. The first implementation cannot prevent access to globals, but the tool contract should guide generated code through `run(ctx)`.

One practical evaluator approach:

```text
TS source
  -> compile to CommonJS-style JS
  -> new AsyncFunction("exports", compiledJs)
  -> validate exports.run
```

For non-exported fallback support, the runtime can append a small trailer to the compiled JavaScript before evaluation:

```js
if (typeof run === "function" && !exports.run) {
  exports.run = run;
}
```

The evaluator should produce:

- compile diagnostics,
- execution logs,
- returned value from `run(ctx)`,
- elapsed time,
- clear error messages with stack traces where available.

Direct `eval`, `Function`, or `AsyncFunction` use should be isolated in one file, for example:

```text
packages/office-runtime/src/evaluation/unsafe-evaluator.ts
```

That file should include a short comment naming the deferred sandboxing work.

## Result Contract

Tool results should be concise text for the model plus details for UI/debugging:

```ts
interface OfficeCodeExecutionDetails {
  status: "success" | "error";
  diagnostics: OfficeCodeDiagnostic[];
  logs: OfficeCodeLogEntry[];
  returnValue?: unknown;
  elapsedMs: number;
}
```

On success without returned data, content should summarize:

```text
OfficeJS code executed successfully.
```

If `run(ctx)` returns JSON-serializable data, include that JSON in the text result so the model can directly see workbook data returned by the script. The same JSON-compatible value should also be stored in `details.returnValue` for UI/debugging.

On compile or runtime failure, throw from `execute()` and let the Pi agent loop convert it to an error tool result. The thrown message should be actionable:

```text
TypeScript compilation failed: Line 4, Column 12: ...
```

## Prompting Changes

The taskpane system prompt must change from "no workbook tools yet" to a narrow tool-use policy.

It should instruct the model to:

- Use `execute_officejs_code` for workbook inspection and mutation.
- Use the Excel OfficeJS API through `ctx`.
- Write TypeScript code with `export async function run(ctx: ExcelRuntimeContext)`.
- Prefer normal OfficeJS batching with `ctx.sync()`.
- Return a JSON-serializable object containing any workbook data or results it needs to see.
- Avoid imports, DOM access, network access, timers, or browser globals.
- Inspect and mutate the workbook directly when the user asks for workbook changes.
- Treat execution as unsandboxed and avoid secrets or unrelated browser state.

These are prompt-level guardrails only. They are not security controls.

## Error Handling

The tool should distinguish:

- invalid tool arguments,
- TypeScript compile diagnostics,
- missing `run` export,
- Office readiness failure,
- OfficeJS runtime failure,
- user cancellation through `AbortSignal`.

The first implementation should favor clear messages over recovery logic. The model can retry with corrected code after receiving a tool error.

## Security Posture

This milestone is intentionally unsafe.

Known risks:

- Evaluated code can access browser globals.
- Evaluated code can call network APIs exposed by the WebView.
- Evaluated code can inspect taskpane application state if reachable.
- Evaluated code can mutate the workbook through OfficeJS.
- Evaluated code can run expensive loops unless interrupted by the host.

Required mitigations for this milestone:

- Keep the evaluator isolated and visibly named unsafe.
- Keep the tool prompt honest about the lack of sandboxing.
- Do not persist generated code as trusted macros.
- Keep the API surface small.

Deferred sandboxing work should include a real execution compartment, policy validation, cancellation instrumentation, capability restrictions, and import control.

## Resolved Decisions

- The tool should be available as product functionality, not hidden behind a development flag.
- The model may inspect and mutate the workbook directly after the user asks for workbook work.
- `run(ctx)` receives both `ctx.context` and convenience aliases such as `ctx.workbook` and `ctx.sync()`.
- Typechecking is in scope for the first runtime milestone.
- Typechecking and compilation should run in a web worker.
- Compile diagnostics can be returned to the model transcript only for the first pass.

## Phased Implementation Plan

### Phase 1: Runtime Package, Compiler Worker, and Unsafe Evaluator

Branch: `codex/officejs-code-tool-phase-1`

PR scope:

- Add `packages/office-runtime`.
- Add TypeScript as a runtime dependency where the browser bundle can include it.
- Keep `execution.ts` at the package root as the primary runtime entrypoint, with supporting implementation under `compiler/`, `runtime/`, and `evaluation/`.
- Implement a lazy compiler worker client.
- Implement `compileOfficeCode(source)` using a virtual TypeScript compiler host.
- Include Milton runtime declarations and the installed `@types/office-js` declarations in the worker.
- Implement unsafe evaluator support for `export async function run(ctx)`.
- Implement `executeOfficeCode(source, options)` with logs, diagnostics, elapsed time, and result normalization.
- Add unit tests for compilation, semantic diagnostics, missing `run`, thrown runtime errors, returned summaries, and log capture.

Dependencies:

- Starts from this design-doc branch after review.

Validation:

- `pnpm --filter @milton/office-runtime build`
- targeted runtime and compiler tests
- `pnpm build` if package wiring touches workspace build behavior

### Phase 2: Agent Tool Wiring in the Taskpane

Branch: `codex/officejs-code-tool-phase-2`

PR scope:

- Add `createExecuteOfficeJsCodeTool()` using the Pi `AgentTool` interface.
- Register the tool in `ExcelTaskpaneApp`.
- Set tool execution to sequential.
- Update the Milton system prompt for OfficeJS tool use.
- Display basic tool execution status and errors in the taskpane conversation.

Dependencies:

- Stacked on Phase 1.

Validation:

- `pnpm --filter @milton/ui build`
- `pnpm --filter @milton/office build`
- Manual Excel taskpane smoke test:
  - ask Milton to write a value to `A1`,
  - ask Milton to read it back,
  - confirm errors are surfaced when generated code is invalid.

### Phase 3: Diagnostics and Developer UX

Branch: `codex/officejs-code-tool-phase-3`

PR scope:

- Improve compile diagnostic formatting with line and column mapping.
- Stream tool logs through `tool_execution_update`.
- Add UI affordances for viewing code, logs, and returned details.
- Add transcript details that preserve executed code and result metadata for debugging.

Dependencies:

- Stacked on Phase 2.

Validation:

- `pnpm --filter @milton/ui build`
- targeted UI tests if the repo has a test harness by then
- manual smoke tests for compile error, runtime error, successful mutation, and log output

### Phase 4: Type Surface Hardening and Pre-Sandbox Policy Prep

Branch: `codex/officejs-code-tool-phase-4`

PR scope:

- Expand and harden the maintained runtime declaration surface based on Phase 1-3 usage.
- Refine prompt examples around `ctx.context`, `ctx.workbook`, and `ctx.sync()`.
- Add lightweight source preflight checks that produce warnings, not security decisions.
- Document the future sandbox interface and the replacement point for the unsafe evaluator.

Dependencies:

- Stacked on Phase 3.

Validation:

- package builds
- tests for preflight warning behavior
- manual verification that normal OfficeJS snippets still execute

## Later Work

Sandboxing should be the next major effort after the code tool proves useful. It should not be treated as a patch on top of the unsafe evaluator. The runtime should define an evaluator interface early so the unsafe evaluator can be replaced by a sandboxed implementation without changing the agent tool contract.

Future efforts:

- sandboxed execution compartment,
- static policy validation,
- AST instrumentation,
- cancellation injection,
- import/package control,
- compiled artifact caching,
- persisted macro storage,
- document-packaged macro artifacts,
- reusable Office helper libraries,
- multi-host runtime contexts for Word and PowerPoint.
