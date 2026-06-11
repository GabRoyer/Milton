import { cloneJsonSafe, installDeniedGlobals, resetGlobalProperties } from "./environment";
import type { OfficeCodeLogEntry } from "../runtime/context";

/** Request sent from the host thread to execute compiled JavaScript. */
interface ExecutionWorkerRequest {
  /** Request id echoed back in worker responses. */
  id: number;
  /** Discriminator for execution requests. */
  type: "execute";
  /** Evaluator-compatible JavaScript emitted by the Office code compiler. */
  javascript: string;
}

/** Minimal runtime context available in the Phase 1 execution worker. */
interface WorkerRuntimeContext {
  /** Placeholder for raw Excel context until the worker command builder exists. */
  context: never;
  /** Placeholder for workbook proxy until the worker command builder exists. */
  workbook: never;
  /** Placeholder sync method until workbook RPC is implemented. */
  sync(): Promise<void>;
  /** Captures structured execution logs for the tool result. */
  log(message: string, details?: unknown): void;
  /** Cancellation-like state exposed for code that checks ctx.signal?.aborted. */
  signal: {
    /** Whether the run has been cooperatively aborted. */
    aborted: boolean;
  };
}

const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (
  ...args: string[]
) => (...args: unknown[]) => Promise<void>;

installDeniedGlobals();

/** Handles execution requests posted to the worker. */
self.addEventListener("message", (event: MessageEvent<ExecutionWorkerRequest>) => {
  const request = event.data;

  if (request.type !== "execute") {
    return;
  }

  void runRequest(request);
});

/** Evaluates and runs one compiled JavaScript request. */
async function runRequest(request: ExecutionWorkerRequest): Promise<void> {
  try {
    const exports: Record<string, unknown> = {};
    const evaluator = new AsyncFunction(
      "exports",
      `${request.javascript}
if (typeof run === "function" && !exports.run) {
  exports.run = run;
}`,
    );

    await evaluator(exports);

    if (typeof exports.run !== "function") {
      throw new Error("OfficeJS code must export an async function named run(ctx).");
    }

    const returnValue = await exports.run(createRuntimeContext(request.id));

    self.postMessage({
      id: request.id,
      type: "result",
      returnValue,
    });
  } catch (error) {
    self.postMessage({
      id: request.id,
      type: "error",
      message: error instanceof Error ? error.message : String(error),
    });
  } finally {
    resetGlobalProperties();
  }
}

/** Creates the minimal runtime context exposed to generated code in Phase 1. */
function createRuntimeContext(requestId: number): WorkerRuntimeContext {
  const context = {
    signal: {
      aborted: false,
    },
    sync: async () => {
      throw new Error("Workbook sync is not available in the sandbox worker yet.");
    },
    log: (message: string, details?: unknown) => {
      const entry: OfficeCodeLogEntry = {
        message,
        details: cloneJsonSafe(details),
        timestamp: Date.now(),
      };

      self.postMessage({
        id: requestId,
        type: "log",
        entry,
      });
    },
  };

  Object.defineProperties(context, {
    context: {
      get() {
        throw new Error("Raw Excel context is not available in the sandbox worker yet.");
      },
    },
    workbook: {
      get() {
        throw new Error("Workbook APIs are not available in the sandbox worker yet.");
      },
    },
  });

  return context as unknown as WorkerRuntimeContext;
}
