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
  /** TypeScript source that was compiled and executed. */
  code: string;
  /** Compile diagnostics associated with the run. */
  diagnostics: OfficeCodeDiagnostic[];
  /** Logs emitted by generated code. */
  logs: OfficeCodeLogEntry[];
  /** JSON-serializable value returned from run(ctx). */
  returnValue?: unknown;
  /** Total elapsed execution time in milliseconds. */
  elapsedMs: number;
}

/** JSON-safe data returned from generated Office code plus model-facing text. */
interface SerializedReturnValue {
  /** JSON-compatible value stored in tool execution details. */
  value?: unknown;
  /** Pretty-printed JSON sent back through the tool transcript. */
  text?: string;
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
  /** Optional callback invoked whenever generated code emits a log entry. */
  onLog?: (entry: OfficeCodeLogEntry) => void;
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
      throw new OfficeCodeExecutionError(formatCompileErrors(errorDiagnostics), {
        status: "error",
        code: source,
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
        onLog: (entry) => {
          logs.push(entry);
          options.onLog?.(entry);
        },
      });

      throwIfAborted(options.signal);
      returnValue = await module.run(runtimeContext);
      throwIfAborted(options.signal);
    });

    const serializedReturnValue = serializeReturnValue(returnValue);
    const details: OfficeCodeExecutionDetails = {
      status: "success",
      code: source,
      diagnostics: compileResult.diagnostics,
      logs,
      returnValue: serializedReturnValue.value,
      elapsedMs: now() - startedAt,
    };

    return {
      content: getSuccessContent(serializedReturnValue),
      details,
    };
  } catch (error) {
    if (error instanceof OfficeCodeExecutionError) {
      throw error;
    }

    throw new OfficeCodeExecutionError(error instanceof Error ? error.message : String(error), {
      status: "error",
      code: source,
      diagnostics,
      logs,
      elapsedMs: now() - startedAt,
    });
  }
}

let defaultCompilerClient: ReturnType<typeof createOfficeCodeCompilerWorkerClient> | undefined;

/** Compiles with the lazy browser worker used by the Office taskpane. */
async function compileOfficeCodeWithDefaultHost(source: string): Promise<OfficeCodeCompileResult> {
  if (typeof Worker === "undefined") {
    throw new Error("Office code compiler workers are not available in this runtime.");
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

/** Formats TypeScript error diagnostics as concise model-facing text. */
function formatCompileErrors(diagnostics: OfficeCodeDiagnostic[]): string {
  return `TypeScript compilation failed:\n${diagnostics.map(formatCompileDiagnostic).join("\n")}`;
}

/** Formats one TypeScript diagnostic with source location and diagnostic code. */
function formatCompileDiagnostic(diagnostic: OfficeCodeDiagnostic): string {
  const location =
    diagnostic.line !== undefined && diagnostic.column !== undefined
      ? `Line ${diagnostic.line}, Column ${diagnostic.column}: `
      : "";
  const code = diagnostic.code === undefined ? "" : ` [${diagnostic.code}]`;

  return `${location}${diagnostic.message}${code}`;
}

/** Throws when the cooperative execution signal has already been aborted. */
function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw new Error(signal.reason instanceof Error ? signal.reason.message : "OfficeJS code execution was cancelled.");
  }
}

/** Serializes returned data for the model transcript and structured details. */
function serializeReturnValue(value: unknown): SerializedReturnValue {
  if (value === undefined) {
    return {};
  }

  try {
    const text = JSON.stringify(value, null, 2);

    if (text === undefined) {
      throw new Error("run(ctx) returned a value that could not be JSON-serialized.");
    }

    return {
      value: JSON.parse(text),
      text,
    };
  } catch {
    throw new Error("run(ctx) returned a value that could not be JSON-serialized.");
  }
}

/** Builds model-facing tool-result text from execution status and returned data. */
function getSuccessContent(serializedReturnValue: SerializedReturnValue): string {
  if (!serializedReturnValue.text) {
    return "OfficeJS code executed successfully.";
  }

  return `OfficeJS code executed successfully.\n\nReturned data:\n${serializedReturnValue.text}`;
}
