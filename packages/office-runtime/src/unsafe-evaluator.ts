import type { OfficeCodeRunFunction } from "./types";

export interface OfficeCodeModule {
  run: OfficeCodeRunFunction;
}

// Phase 1 intentionally uses direct evaluation; the sandboxed evaluator replaces this module later.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (
  ...args: string[]
) => (...args: unknown[]) => Promise<void>;

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
