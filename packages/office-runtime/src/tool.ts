import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core/types";
import { Type, type Static } from "typebox";
import { executeOfficeCode, type ExecuteOfficeCodeOptions, type OfficeCodeExecutionDetails } from "./execution";
import type { OfficeCodeLogEntry } from "./runtime/context";

/** Agent tool name used by the model to execute Excel OfficeJS code. */
export const EXECUTE_OFFICEJS_CODE_TOOL_NAME = "execute_officejs_code";

/** TypeBox parameters accepted by the OfficeJS execution tool. */
const executeOfficeJsCodeParameters = Type.Object({
  code: Type.String({
    description:
      "TypeScript source that uses the Excel OfficeJS API through ctx. It must define an async function named run(ctx: ExcelRuntimeContext). Return a JSON-serializable object from run(ctx) for any workbook data or results the model should see.",
  }),
});

/** Validated parameters passed into the OfficeJS execution tool. */
type ExecuteOfficeJsCodeParameters = Static<typeof executeOfficeJsCodeParameters>;

/** Options forwarded to the Office code executor when creating the agent tool. */
export type CreateExecuteOfficeJsCodeToolOptions = ExecuteOfficeCodeOptions;

/** Agent tool result produced by the OfficeJS execution tool. */
export type ExecuteOfficeJsCodeToolResult = AgentToolResult<ExecuteOfficeJsCodeToolDetails>;

/** Final details emitted by the OfficeJS tool, including the submitted source for transcript debugging. */
export interface ExecuteOfficeJsCodeToolFinalDetails extends OfficeCodeExecutionDetails {
  /** TypeScript source submitted to the OfficeJS tool. */
  code: string;
}

/** Details emitted while OfficeJS code is still running. */
export interface ExecuteOfficeJsCodeToolUpdateDetails {
  /** Running status for streamed tool updates. */
  status: "running";
  /** TypeScript source currently being executed. */
  code: string;
  /** Log entry that triggered this update. */
  latestLog: OfficeCodeLogEntry;
}

/** Details emitted by the OfficeJS execution tool. */
export type ExecuteOfficeJsCodeToolDetails = ExecuteOfficeJsCodeToolFinalDetails | ExecuteOfficeJsCodeToolUpdateDetails;

/** Creates the Pi agent tool that compiles and runs generated Excel OfficeJS code. */
export function createExecuteOfficeJsCodeTool(
  options: CreateExecuteOfficeJsCodeToolOptions = {},
): AgentTool<typeof executeOfficeJsCodeParameters, ExecuteOfficeJsCodeToolDetails> {
  return {
    label: "Run OfficeJS Code",
    name: EXECUTE_OFFICEJS_CODE_TOOL_NAME,
    description: "Compile and run TypeScript code against the current Excel workbook using the Excel OfficeJS API.",
    parameters: executeOfficeJsCodeParameters,
    executionMode: "sequential",
    /** Executes validated OfficeJS code from the model. */
    async execute(_toolCallId, params: ExecuteOfficeJsCodeParameters, signal, onUpdate): Promise<ExecuteOfficeJsCodeToolResult> {
      try {
        const result = await executeOfficeCode(params.code, {
          ...options,
          onLog: (entry) => {
            options.onLog?.(entry);
            onUpdate?.({
              content: [{ type: "text", text: `OfficeJS log: ${entry.message}` }],
              details: {
                status: "running",
                code: params.code,
                latestLog: entry,
              },
            });
          },
          signal,
        });

        return {
          content: [{ type: "text", text: result.content }],
          details: {
            ...result.details,
            code: params.code,
          },
        };
      } catch (error) {
        throw attachSubmittedCodeToError(error, params.code);
      }
    },
  };
}

/** Adds submitted source to structured thrown-error details for transcript/UI consumers. */
function attachSubmittedCodeToError(error: unknown, code: string): unknown {
  if (typeof error !== "object" || error === null || !("details" in error)) {
    return error;
  }

  const details = (error as { details?: unknown }).details;

  if (typeof details !== "object" || details === null) {
    return error;
  }

  (error as { details: unknown }).details = {
    ...details,
    code,
  };

  return error;
}
