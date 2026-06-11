import type { OfficeCodeLogCallback, OfficeCodeLogEntry } from "../runtime/context";

const DEFAULT_EXECUTION_TIMEOUT_MS = 30_000;

/** Message sent to the execution worker to run compiled JavaScript. */
interface ExecutionWorkerRequest {
  /** Request id used to pair worker responses with callers. */
  id: number;
  /** Discriminator for execution requests. */
  type: "execute";
  /** Evaluator-compatible JavaScript emitted by the Office code compiler. */
  javascript: string;
}

/** Successful execution response from the worker. */
interface ExecutionWorkerResultResponse {
  /** Request id copied from the original request. */
  id: number;
  /** Discriminator for successful responses. */
  type: "result";
  /** JSON-compatible or structured-clone-compatible value returned by run(ctx). */
  returnValue?: unknown;
}

/** Log response streamed from the worker while generated code runs. */
interface ExecutionWorkerLogResponse {
  /** Request id copied from the original request. */
  id: number;
  /** Discriminator for log responses. */
  type: "log";
  /** Structured log entry emitted by ctx.log. */
  entry: OfficeCodeLogEntry;
}

/** Failed execution response from the worker. */
interface ExecutionWorkerErrorResponse {
  /** Request id copied from the original request. */
  id: number;
  /** Discriminator for failed responses. */
  type: "error";
  /** Error message produced while evaluating or running generated code. */
  message: string;
}

/** Union of all execution worker responses. */
type ExecutionWorkerResponse =
  | ExecutionWorkerResultResponse
  | ExecutionWorkerLogResponse
  | ExecutionWorkerErrorResponse;

/** Result returned from one worker execution. */
export interface OfficeCodeExecutionWorkerResult {
  /** Value returned from run(ctx), if any. */
  returnValue?: unknown;
  /** Logs emitted by generated code during the run. */
  logs: OfficeCodeLogEntry[];
}

/** Options for one worker-backed execution. */
export interface OfficeCodeExecutionWorkerRunOptions {
  /** Optional hard wall-clock timeout in milliseconds. */
  timeoutMs?: number;
  /** Optional cancellation signal that kills the active worker run. */
  signal?: AbortSignal;
  /** Optional callback invoked whenever generated code emits a log entry. */
  onLog?: OfficeCodeLogCallback;
}

/** Reusable client for running generated Office code inside a worker. */
export interface OfficeCodeExecutionWorkerClient {
  /** Runs one compiled JavaScript program in the worker. */
  run(javascript: string, options?: OfficeCodeExecutionWorkerRunOptions): Promise<OfficeCodeExecutionWorkerResult>;
  /** Terminates the active run and replaces the worker. */
  kill(reason?: string): void;
  /** Terminates the worker and rejects the active run if present. */
  dispose(): void;
}

/** Construction options for injecting or creating an execution worker. */
export interface CreateOfficeCodeExecutionWorkerClientOptions {
  /** Existing worker instance, primarily for tests or custom hosts. */
  worker?: Worker;
  /** Factory used to create and recreate the worker. */
  createWorker?: () => Worker;
  /** Default hard wall-clock timeout in milliseconds. */
  timeoutMs?: number;
}

/** Creates a reusable request/response client for the Office code execution worker. */
export function createOfficeCodeExecutionWorkerClient(
  options: CreateOfficeCodeExecutionWorkerClientOptions = {},
): OfficeCodeExecutionWorkerClient {
  let injectedWorker = options.worker;
  const createWorker =
    options.createWorker ??
    (() => {
      if (injectedWorker) {
        const worker = injectedWorker;
        injectedWorker = undefined;
        return worker;
      }

      return createDefaultExecutionWorker();
    });
  const defaultTimeoutMs = options.timeoutMs ?? DEFAULT_EXECUTION_TIMEOUT_MS;
  let worker = createWorker();
  let nextId = 1;
  let activeRun: ActiveExecutionRun | undefined;
  let disposed = false;

  attachWorkerListeners(worker);

  return {
    run(javascript, runOptions = {}) {
      if (disposed) {
        return Promise.reject(new Error("Office code execution worker was disposed."));
      }

      if (activeRun) {
        return Promise.reject(new Error("Office code execution worker is already running a script."));
      }

      const id = nextId++;
      const logs: OfficeCodeLogEntry[] = [];
      const timeoutMs = runOptions.timeoutMs ?? defaultTimeoutMs;
      const request: ExecutionWorkerRequest = {
        id,
        type: "execute",
        javascript,
      };

      return new Promise((resolve, reject) => {
        const finish = (error?: Error, result?: OfficeCodeExecutionWorkerResult, replaceWorker = false) => {
          if (!activeRun || activeRun.id !== id) {
            return;
          }

          cleanupActiveRun();

          if (replaceWorker && !disposed) {
            worker.terminate();
            worker = createWorker();
            attachWorkerListeners(worker);
          }

          if (error) {
            reject(error);
            return;
          }

          resolve(result ?? { logs });
        };

        const timeout = setTimeout(() => {
          finish(new Error(`OfficeJS code execution timed out after ${timeoutMs}ms.`), undefined, true);
        }, timeoutMs);

        const abort = () => {
          const message =
            runOptions.signal?.reason instanceof Error
              ? runOptions.signal.reason.message
              : "OfficeJS code execution was cancelled.";

          finish(new Error(message), undefined, true);
        };

        activeRun = {
          id,
          logs,
          onLog: runOptions.onLog,
          finish,
          timeout,
          signal: runOptions.signal,
          abort,
        };

        if (runOptions.signal?.aborted) {
          abort();
          return;
        }

        runOptions.signal?.addEventListener("abort", abort, { once: true });
        worker.postMessage(request);
      });
    },
    kill(reason = "OfficeJS code execution was killed.") {
      if (!activeRun) {
        return;
      }

      activeRun.finish(new Error(reason), undefined, true);
    },
    dispose() {
      disposed = true;

      if (activeRun) {
        activeRun.finish(new Error("Office code execution worker was disposed."));
      }

      worker.terminate();
    },
  };

  /** Registers message and failure listeners on the current worker instance. */
  function attachWorkerListeners(targetWorker: Worker): void {
    targetWorker.addEventListener("message", (event: MessageEvent<ExecutionWorkerResponse>) => {
      if (targetWorker !== worker) {
        return;
      }

      const response = event.data;

      if (!activeRun || response.id !== activeRun.id) {
        return;
      }

      if (response.type === "log") {
        handleLogResponse(activeRun, response.entry);
        return;
      }

      if (response.type === "error") {
        activeRun.finish(new Error(response.message));
        return;
      }

      activeRun.finish(undefined, {
        logs: activeRun.logs,
        returnValue: response.returnValue,
      });
    });

    targetWorker.addEventListener("error", (event) => {
      if (targetWorker !== worker) {
        return;
      }

      if (!activeRun) {
        return;
      }

      activeRun.finish(new Error(event.message || "Office code execution worker failed."), undefined, true);
    });

    targetWorker.addEventListener("messageerror", () => {
      if (targetWorker !== worker) {
        return;
      }

      if (!activeRun) {
        return;
      }

      activeRun.finish(new Error("Office code execution worker sent an unreadable message."), undefined, true);
    });
  }

  /** Handles one streamed log entry from the worker. */
  function handleLogResponse(run: ActiveExecutionRun, entry: OfficeCodeLogEntry): void {
    run.logs.push(entry);

    try {
      run.onLog?.(entry);
    } catch (error) {
      run.finish(error instanceof Error ? error : new Error(String(error)), undefined, true);
    }
  }

  /** Clears timers and abort listeners for the current active run. */
  function cleanupActiveRun(): void {
    if (!activeRun) {
      return;
    }

    clearTimeout(activeRun.timeout);
    activeRun.signal?.removeEventListener("abort", activeRun.abort);
    activeRun = undefined;
  }
}

/** Creates the default module worker that runs generated Office code. */
function createDefaultExecutionWorker(): Worker {
  if (typeof Worker === "undefined") {
    throw new Error("Office code execution workers are not available in this runtime.");
  }

  return new Worker(new URL("./worker.ts", import.meta.url), {
    type: "module",
  });
}

/** State for the one active run owned by a reusable execution worker. */
interface ActiveExecutionRun {
  /** Request id for the active worker run. */
  id: number;
  /** Logs emitted so far by the active run. */
  logs: OfficeCodeLogEntry[];
  /** Optional callback invoked for each log entry. */
  onLog?: OfficeCodeLogCallback;
  /** Completes the run, optionally replacing the worker. */
  finish: (error?: Error, result?: OfficeCodeExecutionWorkerResult, replaceWorker?: boolean) => void;
  /** Timeout handle for the active run. */
  timeout: ReturnType<typeof setTimeout>;
  /** Cancellation signal observed by the active run. */
  signal?: AbortSignal;
  /** Abort listener registered on the cancellation signal. */
  abort: () => void;
}
