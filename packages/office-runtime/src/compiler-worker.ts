import { compileOfficeCode } from "./compiler";

interface CompilerWorkerRequest {
  id: number;
  type: "compile";
  source: string;
}

self.addEventListener("message", (event: MessageEvent<CompilerWorkerRequest>) => {
  const request = event.data;

  if (request.type !== "compile") {
    return;
  }

  try {
    self.postMessage({
      id: request.id,
      type: "result",
      result: compileOfficeCode(request.source),
    });
  } catch (error) {
    self.postMessage({
      id: request.id,
      type: "error",
      message: error instanceof Error ? error.message : String(error),
    });
  }
});
