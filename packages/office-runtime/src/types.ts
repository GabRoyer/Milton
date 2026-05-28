/** Runtime context passed to generated Excel OfficeJS code. */
export interface ExcelRuntimeContext {
  /** Raw Excel request context that owns all OfficeJS objects used during the run. */
  context: Excel.RequestContext;
  /** Current workbook convenience alias from the active request context. */
  workbook: Excel.Workbook;
  /** Synchronizes queued OfficeJS loads and mutations with Excel. */
  sync(): Promise<void>;
  /** Captures structured execution logs for the tool result. */
  log(message: string, details?: unknown): void;
  /** Optional cancellation signal provided by the agent runtime. */
  signal?: AbortSignal;
}

/** Severity levels reported by the Office code compiler. */
export type OfficeCodeDiagnosticSeverity = "error" | "warning" | "suggestion" | "message";

/** Compiler diagnostic formatted for model and UI consumption. */
export interface OfficeCodeDiagnostic {
  /** Diagnostic severity. */
  severity: OfficeCodeDiagnosticSeverity;
  /** Human-readable diagnostic message. */
  message: string;
  /** One-based source line when available. */
  line?: number;
  /** One-based source column when available. */
  column?: number;
  /** TypeScript or Milton-specific diagnostic code. */
  code?: number | string;
}

/** JavaScript output and diagnostics from compiling generated Office code. */
export interface OfficeCodeCompileResult {
  /** Evaluator-compatible JavaScript, empty when compilation fails. */
  javascript: string;
  /** Diagnostics collected during compile and emit. */
  diagnostics: OfficeCodeDiagnostic[];
}

/** Structured log entry emitted by generated Office code. */
export interface OfficeCodeLogEntry {
  /** Log message intended for debugging or model feedback. */
  message: string;
  /** Optional JSON-serializable details associated with the log. */
  details?: unknown;
  /** Unix timestamp recorded when the log is captured. */
  timestamp: number;
}

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

/** Callable shape expected from evaluated generated Office code. */
export type OfficeCodeRunFunction = (ctx: ExcelRuntimeContext) => Promise<unknown> | unknown;
/** Host adapter for running callbacks inside Excel.run. */
export type ExcelRunner = (callback: (context: Excel.RequestContext) => Promise<void>) => Promise<void>;
