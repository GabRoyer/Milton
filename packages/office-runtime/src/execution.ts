import { createOfficeCodeCompilerWorkerClient } from "./compiler/worker-client";
import type { OfficeCodeCompileResult, OfficeCodeDiagnostic } from "./compiler/compile";
import { createExcelRuntimeContext } from "./runtime/context";
import type { OfficeCodeLogEntry } from "./runtime/context";
import { unsafeEvaluateOfficeCode } from "./evaluation/unsafe-evaluator";

/** Host adapter for running callbacks inside Excel.run. */
export type ExcelRunner = (callback: (context: Excel.RequestContext) => Promise<void>) => Promise<void>;

/** Structured execution metadata returned by the Office code tool. */
export interface OfficeCodeExecutionDetails {
  /** Final execution status. */
  status: "success" | "error";
  /** Compile diagnostics associated with the run. */
  diagnostics: OfficeCodeDiagnostic[];
  /** Logs emitted by generated code. */
  logs: OfficeCodeLogEntry[];
  /** JSON-serializable value returned from run(ctx). */
  returnValue?: unknown;
  /** Total elapsed execution time in milliseconds. */
  elapsedMs: number;
}

/** User/model-facing content plus structured execution metadata. */
export interface OfficeCodeExecutionResult {
  /** Concise text result sent back through the tool transcript. */
  content: string;
  /** Structured details for UI and debugging. */
  details: OfficeCodeExecutionDetails;
}

/** Dependency injection points for executing generated Office code. */
export interface ExecuteOfficeCodeOptions {
  /** Optional compiler override, typically for tests. */
  compile?: (source: string) => OfficeCodeCompileResult | Promise<OfficeCodeCompileResult>;
  /** Optional evaluator override, typically for future sandboxing or tests. */
  evaluate?: typeof unsafeEvaluateOfficeCode;
  /** Optional Excel.run adapter override, typically for tests. */
  excelRunner?: ExcelRunner;
  /** Optional cancellation signal checked around cooperative execution boundaries. */
  signal?: AbortSignal;
  /** Optional clock override used for deterministic elapsed-time tests. */
  now?: () => number;
}

/** Error wrapper that carries structured execution details. */
export class OfficeCodeExecutionError extends Error {
  /** Structured details available to UI and tool-result handling. */
  public readonly details: OfficeCodeExecutionDetails;

  /** Creates an execution error with model-facing text and structured details. */
  constructor(message: string, details: OfficeCodeExecutionDetails) {
    super(message);
    this.name = "OfficeCodeExecutionError";
    this.details = details;
  }
}

/** Compiles, evaluates, and executes generated TypeScript against the Excel host. */
export async function executeOfficeCode(
  source: string,
  options: ExecuteOfficeCodeOptions = {},
): Promise<OfficeCodeExecutionResult> {
  const now = options.now ?? Date.now;
  const startedAt = now();
  const logs: OfficeCodeLogEntry[] = [];
  const compile = options.compile ?? compileOfficeCodeWithDefaultHost;
  const evaluate = options.evaluate ?? unsafeEvaluateOfficeCode;
  let diagnostics: OfficeCodeDiagnostic[] = [];

  try {
    throwIfAborted(options.signal);

    const compileResult = await compile(source);
    diagnostics = compileResult.diagnostics;
    const errorDiagnostics = compileResult.diagnostics.filter((diagnostic) => diagnostic.severity === "error");

    if (errorDiagnostics.length > 0) {
      throw new OfficeCodeExecutionError(formatCompileError(errorDiagnostics[0]), {
        status: "error",
        diagnostics: compileResult.diagnostics,
        logs,
        elapsedMs: now() - startedAt,
      });
    }

    throwIfAborted(options.signal);

    const module = await evaluate(compileResult.javascript);
    let returnValue: unknown;

    throwIfAborted(options.signal);

    await (options.excelRunner ?? defaultExcelRunner)(async (context) => {
      const runtimeContext = createExcelRuntimeContext(context, {
        signal: options.signal,
        onLog: (entry) => logs.push(entry),
      });

      throwIfAborted(options.signal);
      returnValue = await module.run(runtimeContext);
      throwIfAborted(options.signal);
    });

    const jsonReturnValue = toJsonSerializable(returnValue);
    const details: OfficeCodeExecutionDetails = {
      status: "success",
      diagnostics: compileResult.diagnostics,
      logs,
      returnValue: jsonReturnValue,
      elapsedMs: now() - startedAt,
    };

    return {
      content: getSuccessContent(jsonReturnValue),
      details,
    };
  } catch (error) {
    if (error instanceof OfficeCodeExecutionError) {
      throw error;
    }

    throw new OfficeCodeExecutionError(error instanceof Error ? error.message : String(error), {
      status: "error",
      diagnostics,
      logs,
      elapsedMs: now() - startedAt,
    });
  }
}

let defaultCompilerClient: ReturnType<typeof createOfficeCodeCompilerWorkerClient> | undefined;

/** Compiles with a lazy worker in browsers and falls back to direct compilation elsewhere. */
async function compileOfficeCodeWithDefaultHost(source: string): Promise<OfficeCodeCompileResult> {
  if (typeof Worker === "undefined") {
    const { compileOfficeCode } = await import("./compiler/compile");
    return compileOfficeCode(source);
  }

  defaultCompilerClient ??= createOfficeCodeCompilerWorkerClient();
  return defaultCompilerClient.compile(source);
}

/** Runs an OfficeJS callback through the ambient Excel host. */
async function defaultExcelRunner(callback: Parameters<ExcelRunner>[0]): Promise<void> {
  const excel = (globalThis as typeof globalThis & { Excel?: { run?: ExcelRunner } }).Excel;

  if (!excel?.run) {
    throw new Error("Excel APIs are not available in this host.");
  }

  await excel.run(callback);
}

/** Formats the first TypeScript diagnostic as a concise tool error. */
function formatCompileError(diagnostic: OfficeCodeExecutionDetails["diagnostics"][number]): string {
  const location =
    diagnostic.line !== undefined && diagnostic.column !== undefined
      ? `Line ${diagnostic.line}, Column ${diagnostic.column}: `
      : "";

  return `TypeScript compilation failed: ${location}${diagnostic.message}`;
}

/** Throws when the cooperative execution signal has already been aborted. */
function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw new Error(signal.reason instanceof Error ? signal.reason.message : "OfficeJS code execution was cancelled.");
  }
}

/** Normalizes returned data to the JSON-compatible contract promised to the model. */
function toJsonSerializable(value: unknown): unknown {
  if (value === undefined) {
    return undefined;
  }

  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    throw new Error("run(ctx) returned a value that could not be JSON-serialized.");
  }
}

/** Builds concise tool-result text from the generated code return value. */
function getSuccessContent(returnValue: unknown): string {
  if (isSummaryObject(returnValue)) {
    return returnValue.summary;
  }

  return "OfficeJS code executed successfully.";
}

/** Checks whether a returned value includes a model-facing summary string. */
function isSummaryObject(value: unknown): value is { summary: string } {
  return (
    typeof value === "object" &&
    value !== null &&
    "summary" in value &&
    typeof (value as { summary: unknown }).summary === "string"
  );
}
