import type { OfficeCodeCompileResult } from "./types";

interface CompilerWorkerRequest {
  id: number;
  type: "compile";
  source: string;
}

interface CompilerWorkerResult {
  id: number;
  type: "result";
  result: OfficeCodeCompileResult;
}

interface CompilerWorkerError {
  id: number;
  type: "error";
  message: string;
}

type CompilerWorkerResponse = CompilerWorkerResult | CompilerWorkerError;

export interface OfficeCodeCompilerWorkerClient {
  compile(source: string): Promise<OfficeCodeCompileResult>;
  dispose(): void;
}

export interface CreateOfficeCodeCompilerWorkerClientOptions {
  worker?: Worker;
  createWorker?: () => Worker;
}

export function createOfficeCodeCompilerWorkerClient(
  options: CreateOfficeCodeCompilerWorkerClientOptions = {},
): OfficeCodeCompilerWorkerClient {
  const worker = options.worker ?? options.createWorker?.() ?? createDefaultCompilerWorker();
  const pending = new Map<number, PendingCompile>();
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

function createDefaultCompilerWorker(): Worker {
  if (typeof Worker === "undefined") {
    throw new Error("Office code compiler workers are not available in this runtime.");
  }

  return new Worker(new URL("./compiler-worker.ts", import.meta.url), {
    type: "module",
  });
}

interface PendingCompile {
  resolve: (result: OfficeCodeCompileResult) => void;
  reject: (error: Error) => void;
}
