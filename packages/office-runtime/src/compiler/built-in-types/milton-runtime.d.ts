/** Runtime context passed to generated Excel OfficeJS code. */
interface ExcelRuntimeContext {
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
