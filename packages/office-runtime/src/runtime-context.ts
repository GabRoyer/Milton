import type { ExcelRuntimeContext, OfficeCodeLogEntry } from "./types";

/** Options used to build the generated-code runtime context. */
export interface CreateExcelRuntimeContextOptions {
  /** Optional cancellation signal passed through to generated code. */
  signal?: AbortSignal;
  /** Mutable log sink populated by ctx.log calls. */
  logs?: OfficeCodeLogEntry[];
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
      options.logs?.push({
        message,
        details,
        timestamp: Date.now(),
      });
    },
  };
}
