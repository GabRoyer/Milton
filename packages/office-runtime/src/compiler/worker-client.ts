import type { OfficeCodeCompileResult } from "./compile";

/** Message sent to the compiler worker to compile one source string. */
interface CompilerWorkerRequest {
  /** Request id used to pair worker responses with callers. */
  id: number;
  /** Discriminator for compile requests. */
  type: "compile";
  /** Generated TypeScript source to compile. */
  source: string;
}

/** Successful compiler worker response. */
interface CompilerWorkerResult {
  /** Request id copied from the original request. */
  id: number;
  /** Discriminator for successful responses. */
  type: "result";
  /** Compile result produced by the worker. */
  result: OfficeCodeCompileResult;
}

/** Failed compiler worker response. */
interface CompilerWorkerError {
  /** Request id copied from the original request. */
  id: number;
  /** Discriminator for failed responses. */
  type: "error";
  /** Error message produced while compiling. */
  message: string;
}

/** Union of all compiler worker responses. */
type CompilerWorkerResponse = CompilerWorkerResult | CompilerWorkerError;

/** Client wrapper that tracks pending compiler worker requests. */
export interface OfficeCodeCompilerWorkerClient {
  /** Compiles generated TypeScript source in the worker. */
  compile(source: string): Promise<OfficeCodeCompileResult>;
  /** Terminates the worker and rejects pending compile requests. */
  dispose(): void;
}

/** Construction options for injecting or creating a compiler worker. */
export interface CreateOfficeCodeCompilerWorkerClientOptions {
  /** Existing worker instance, primarily for tests or custom hosts. */
  worker?: Worker;
  /** Factory used to lazily create the worker. */
  createWorker?: () => Worker;
}

/** Creates a request/response client for the Office code compiler worker. */
export function createOfficeCodeCompilerWorkerClient(
  options: CreateOfficeCodeCompilerWorkerClientOptions = {},
): OfficeCodeCompilerWorkerClient {
  const worker = options.worker ?? options.createWorker?.() ?? createDefaultCompilerWorker();
  const pending = new Map<number, PendingCompileRequest>();
  let nextId = 1;

  worker.addEventListener("message", (event: MessageEvent<CompilerWorkerResponse>) => {
    const response = event.data;
    const request = pending.get(response.id);

    if (!request) {
      return;
    }

    pending.delete(response.id);

    if (response.type === "error") {
      request.reject(new Error(response.message));
      return;
    }

    request.resolve(response.result);
  });

  worker.addEventListener("error", (event) => {
    for (const request of pending.values()) {
      request.reject(new Error(event.message || "Office code compiler worker failed."));
    }

    pending.clear();
  });

  return {
    compile(source) {
      const id = nextId++;
      const request: CompilerWorkerRequest = {
        id,
        type: "compile",
        source,
      };

      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject });
        worker.postMessage(request);
      });
    },
    dispose() {
      for (const request of pending.values()) {
        request.reject(new Error("Office code compiler worker was disposed."));
      }

      pending.clear();
      worker.terminate();
    },
  };
}

/** Creates the default module worker that hosts the TypeScript compiler. */
function createDefaultCompilerWorker(): Worker {
  if (typeof Worker === "undefined") {
    throw new Error("Office code compiler workers are not available in this runtime.");
  }

  return new Worker(new URL("./worker.ts", import.meta.url), {
    type: "module",
  });
}

/** Pending compile request callbacks stored until the worker responds. */
interface PendingCompileRequest {
  /** Resolves the compile promise with a worker result. */
  resolve: (result: OfficeCodeCompileResult) => void;
  /** Rejects the compile promise when worker execution fails. */
  reject: (error: Error) => void;
}
