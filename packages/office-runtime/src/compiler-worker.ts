import { compileOfficeCode } from "./compiler";

/** Message sent from the host thread to the compiler worker. */
interface CompilerWorkerRequest {
  /** Request id echoed back in the worker response. */
  id: number;
  /** Discriminator for compile requests. */
  type: "compile";
  /** Generated TypeScript source to compile. */
  source: string;
}

/** Handles compile requests posted to the worker and returns serialized results. */
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
