export interface ExcelRuntimeContext {
  context: Excel.RequestContext;
  workbook: Excel.Workbook;
  sync(): Promise<void>;
  log(message: string, details?: unknown): void;
  signal?: AbortSignal;
}

export type OfficeCodeDiagnosticSeverity = "error" | "warning" | "suggestion" | "message";

export interface OfficeCodeDiagnostic {
  severity: OfficeCodeDiagnosticSeverity;
  message: string;
  line?: number;
  column?: number;
  code?: number | string;
}

export interface OfficeCodeCompileResult {
  javascript: string;
  diagnostics: OfficeCodeDiagnostic[];
}

export interface OfficeCodeLogEntry {
  message: string;
  details?: unknown;
  timestamp: number;
}

export interface OfficeCodeExecutionDetails {
  status: "success" | "error";
  diagnostics: OfficeCodeDiagnostic[];
  logs: OfficeCodeLogEntry[];
  returnValue?: unknown;
  elapsedMs: number;
}

export interface OfficeCodeExecutionResult {
  content: string;
  details: OfficeCodeExecutionDetails;
}

export type OfficeCodeRunFunction = (ctx: ExcelRuntimeContext) => Promise<unknown> | unknown;
export type ExcelRunner = (callback: (context: Excel.RequestContext) => Promise<void>) => Promise<void>;
