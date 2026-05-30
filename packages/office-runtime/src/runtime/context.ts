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

/** Structured log entry emitted by generated Office code. */
export interface OfficeCodeLogEntry {
  /** Log message intended for debugging or model feedback. */
  message: string;
  /** Optional JSON-serializable details associated with the log. */
  details?: unknown;
  /** Unix timestamp recorded when the log is captured. */
  timestamp: number;
}

/** Receives structured log entries emitted by generated Office code. */
export type OfficeCodeLogCallback = (entry: OfficeCodeLogEntry) => void;

/** Options used to build the generated-code runtime context. */
export interface CreateExcelRuntimeContextOptions {
  /** Optional cancellation signal passed through to generated code. */
  signal?: AbortSignal;
  /** Optional callback invoked for each ctx.log call. */
  onLog?: OfficeCodeLogCallback;
}

/** Creates the minimal Excel runtime context exposed to generated code. */
export function createExcelRuntimeContext(
  context: Excel.RequestContext,
  options: CreateExcelRuntimeContextOptions = {},
): ExcelRuntimeContext {
  return {
    context,
    workbook: context.workbook,
    signal: options.signal,
    sync: () => context.sync(),
    log: (message, details) => {
      options.onLog?.({
        message,
        details,
        timestamp: Date.now(),
      });
    },
  };
}
