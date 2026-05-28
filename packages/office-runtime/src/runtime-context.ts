import type { ExcelRuntimeContext, OfficeCodeLogEntry } from "./types";

export interface CreateExcelRuntimeContextOptions {
  signal?: AbortSignal;
  logs?: OfficeCodeLogEntry[];
}

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
      options.logs?.push({
        message,
        details,
        timestamp: Date.now(),
      });
    },
  };
}
