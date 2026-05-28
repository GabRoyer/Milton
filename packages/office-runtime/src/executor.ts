import { compileOfficeCode } from "./compiler";
import { createOfficeCodeCompilerWorkerClient } from "./compiler-worker-client";
import { createExcelRuntimeContext } from "./runtime-context";
import type {
  ExcelRunner,
  OfficeCodeCompileResult,
  OfficeCodeDiagnostic,
  OfficeCodeExecutionDetails,
  OfficeCodeExecutionResult,
  OfficeCodeLogEntry,
} from "./types";
import { unsafeEvaluateOfficeCode } from "./unsafe-evaluator";

export interface ExecuteOfficeCodeOptions {
  compile?: (source: string) => OfficeCodeCompileResult | Promise<OfficeCodeCompileResult>;
  evaluate?: typeof unsafeEvaluateOfficeCode;
  excelRunner?: ExcelRunner;
  signal?: AbortSignal;
  now?: () => number;
}

export class OfficeCodeExecutionError extends Error {
  public readonly details: OfficeCodeExecutionDetails;

  constructor(message: string, details: OfficeCodeExecutionDetails) {
    super(message);
    this.name = "OfficeCodeExecutionError";
    this.details = details;
  }
}

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
        logs,
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

function compileOfficeCodeWithDefaultHost(source: string): OfficeCodeCompileResult | Promise<OfficeCodeCompileResult> {
  if (typeof Worker === "undefined") {
    return compileOfficeCode(source);
  }

  defaultCompilerClient ??= createOfficeCodeCompilerWorkerClient();
  return defaultCompilerClient.compile(source);
}

async function defaultExcelRunner(callback: Parameters<ExcelRunner>[0]): Promise<void> {
  const excel = (globalThis as typeof globalThis & { Excel?: { run?: ExcelRunner } }).Excel;

  if (!excel?.run) {
    throw new Error("Excel APIs are not available in this host.");
  }

  await excel.run(callback);
}

function formatCompileError(diagnostic: OfficeCodeExecutionDetails["diagnostics"][number]): string {
  const location =
    diagnostic.line !== undefined && diagnostic.column !== undefined
      ? `Line ${diagnostic.line}, Column ${diagnostic.column}: `
      : "";

  return `TypeScript compilation failed: ${location}${diagnostic.message}`;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw new Error(signal.reason instanceof Error ? signal.reason.message : "OfficeJS code execution was cancelled.");
  }
}

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

function getSuccessContent(returnValue: unknown): string {
  if (isSummaryObject(returnValue)) {
    return returnValue.summary;
  }

  return "OfficeJS code executed successfully.";
}

function isSummaryObject(value: unknown): value is { summary: string } {
  return (
    typeof value === "object" &&
    value !== null &&
    "summary" in value &&
    typeof (value as { summary: unknown }).summary === "string"
  );
}
