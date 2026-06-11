import { describe, expect, it } from "vitest";
import { compileOfficeCode } from "../src/compiler/compile";
import { executeOfficeCode } from "../src/execution";
import { createOfficeCodeExecutionWorkerClient } from "../src/execution/worker-client";

/** Callback invoked whenever the fake worker receives a host message. */
type FakeWorkerPostMessageHandler = (worker: FakeWorker, message: unknown) => void;

/** Minimal worker test double used by execution-worker client tests. */
class FakeWorker {
  /** Messages posted from the host into the fake worker. */
  public readonly messages: unknown[] = [];
  /** Whether terminate has been called. */
  public terminated = false;

  private readonly listeners = new Map<string, Set<EventListenerOrEventListenerObject>>();

  /** Creates a fake worker with an optional postMessage observer. */
  constructor(private readonly onPostMessage?: FakeWorkerPostMessageHandler) {}

  /** Registers a worker event listener. */
  public addEventListener(type: string, listener: EventListenerOrEventListenerObject): void {
    let listeners = this.listeners.get(type);

    if (!listeners) {
      listeners = new Set();
      this.listeners.set(type, listeners);
    }

    listeners.add(listener);
  }

  /** Removes a worker event listener. */
  public removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void {
    this.listeners.get(type)?.delete(listener);
  }

  /** Captures a host message sent to the worker. */
  public postMessage(message: unknown): void {
    this.messages.push(message);
    this.onPostMessage?.(this, message);
  }

  /** Marks the fake worker as terminated. */
  public terminate(): void {
    this.terminated = true;
  }

  /** Emits a worker message event to the host. */
  public emitMessage(data: unknown): void {
    this.dispatch("message", { data } as MessageEvent);
  }

  /** Emits a worker error event to the host. */
  public emitError(message: string): void {
    this.dispatch("error", { message } as ErrorEvent);
  }

  /** Casts the fake to the DOM Worker shape expected by production code. */
  public asWorker(): Worker {
    return this as unknown as Worker;
  }

  /** Dispatches an event to registered listeners. */
  private dispatch(type: string, event: Event): void {
    for (const listener of this.listeners.get(type) ?? []) {
      if (typeof listener === "function") {
        listener(event);
      } else {
        listener.handleEvent(event);
      }
    }
  }
}

/** Reads the id from the latest message posted to a fake worker. */
function getLatestRequestId(worker: FakeWorker): number {
  const message = worker.messages.at(-1);

  if (typeof message !== "object" || message === null || !("id" in message)) {
    throw new Error("Expected the fake worker to receive a request with an id.");
  }

  return Number((message as { id: unknown }).id);
}

describe("createOfficeCodeExecutionWorkerClient", () => {
  it("runs compiled JavaScript through executeOfficeCode when injected", async () => {
    const worker = new FakeWorker((fakeWorker, message) => {
      const id = Number((message as { id: unknown }).id);

      fakeWorker.emitMessage({
        id,
        type: "log",
        entry: {
          message: "worker log",
          details: { source: "worker" },
          timestamp: 1,
        },
      });
      fakeWorker.emitMessage({
        id,
        type: "result",
        returnValue: { message: "Worker ran." },
      });
    });
    const executionWorker = createOfficeCodeExecutionWorkerClient({
      createWorker: () => worker.asWorker(),
      timeoutMs: 1_000,
    });

    const result = await executeOfficeCode(
      `
export async function run(ctx: ExcelRuntimeContext) {
  ctx.log("unused");
  return { message: "unused" };
}
`,
      {
        compile: compileOfficeCode,
        executionWorker,
      },
    );

    expect(worker.messages[0]).toMatchObject({
      type: "execute",
      javascript: expect.stringContaining("exports.run = run"),
    });
    expect(result.content).toContain("Worker ran.");
    expect(result.details).toMatchObject({
      status: "success",
      logs: [
        {
          message: "worker log",
          details: { source: "worker" },
        },
      ],
      returnValue: { message: "Worker ran." },
    });
  });

  it("rejects concurrent scripts on the same reusable worker", async () => {
    const worker = new FakeWorker();
    const executionWorker = createOfficeCodeExecutionWorkerClient({
      createWorker: () => worker.asWorker(),
      timeoutMs: 1_000,
    });
    const firstRun = executionWorker.run("exports.run = async () => undefined;").catch((error: unknown) => error);

    await expect(executionWorker.run("exports.run = async () => undefined;")).rejects.toThrow(
      "already running a script",
    );

    executionWorker.kill("stopped");
    await expect(firstRun).resolves.toMatchObject({ message: "stopped" });
  });

  it("terminates and recreates the worker after a timeout", async () => {
    const workers: FakeWorker[] = [];
    const executionWorker = createOfficeCodeExecutionWorkerClient({
      createWorker: () => {
        const worker = new FakeWorker();
        workers.push(worker);
        return worker.asWorker();
      },
      timeoutMs: 1_000,
    });

    await expect(executionWorker.run("exports.run = async () => undefined;", { timeoutMs: 1 })).rejects.toThrow(
      "timed out",
    );
    expect(workers[0]?.terminated).toBe(true);
    expect(workers).toHaveLength(2);

    const secondRun = executionWorker.run("exports.run = async () => 'ok';", { timeoutMs: 1_000 });
    workers[1]?.emitMessage({
      id: getLatestRequestId(workers[1]),
      type: "result",
      returnValue: "ok",
    });

    await expect(secondRun).resolves.toEqual({
      logs: [],
      returnValue: "ok",
    });
  });

  it("terminates and recreates the worker when the signal aborts", async () => {
    const workers: FakeWorker[] = [];
    const executionWorker = createOfficeCodeExecutionWorkerClient({
      createWorker: () => {
        const worker = new FakeWorker();
        workers.push(worker);
        return worker.asWorker();
      },
      timeoutMs: 1_000,
    });
    const controller = new AbortController();
    const run = executionWorker.run("exports.run = async () => undefined;", {
      signal: controller.signal,
      timeoutMs: 1_000,
    });

    controller.abort(new Error("user killed"));

    await expect(run).rejects.toThrow("user killed");
    expect(workers[0]?.terminated).toBe(true);
    expect(workers).toHaveLength(2);
  });
});
