import type { ExcelRuntimeContext } from "../runtime/context";

/** Callable shape expected from evaluated generated Office code. */
export type OfficeCodeRunFunction = (ctx: ExcelRuntimeContext) => Promise<unknown> | unknown;

/** Evaluated module shape expected from generated Office code. */
export interface OfficeCodeModule {
  /** Entrypoint invoked with Milton's Excel runtime context. */
  run: OfficeCodeRunFunction;
}

/** Evaluates compiled JavaScript into a runnable Office code module. */
export type OfficeCodeEvaluator = (javascript: string) => Promise<OfficeCodeModule>;

// Direct evaluation is the deliberate temporary implementation; a sandboxed evaluator should replace this interface boundary.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (
  ...args: string[]
) => (...args: unknown[]) => Promise<void>;

/** Evaluates compiler output and returns the generated Office code module. */
export async function unsafeEvaluateOfficeCode(javascript: string): Promise<OfficeCodeModule> {
  const exports: Record<string, unknown> = {};
  const evaluator = new AsyncFunction(
    "exports",
    `${javascript}
if (typeof run === "function" && !exports.run) {
  exports.run = run;
}`,
  );

  await evaluator(exports);

  if (typeof exports.run !== "function") {
    throw new Error("OfficeJS code must export an async function named run(ctx).");
  }

  return {
    run: exports.run as OfficeCodeRunFunction,
  };
}
