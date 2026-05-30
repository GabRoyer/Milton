import { describe, expect, it } from "vitest";
import { compileOfficeCode } from "../src/compiler/compile";
import { executeOfficeCode, OfficeCodeExecutionError, type ExcelRunner } from "../src/execution";
import { unsafeEvaluateOfficeCode } from "../src/evaluation/unsafe-evaluator";

/** Minimal Excel.run adapter used by runtime tests. */
const mockExcelRunner: ExcelRunner = async (callback) => {
  await callback({
    workbook: {},
    sync: async () => undefined,
  } as unknown as Excel.RequestContext);
};

describe("compileOfficeCode", () => {
  it("typechecks and emits Excel OfficeJS code", () => {
    const result = compileOfficeCode(`
export async function run(ctx: ExcelRuntimeContext) {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:C10");
  range.load(["address", "values"]);
  await ctx.sync();
  return {
    message: "Read A1:C10.",
    address: range.address,
    values: range.values,
  };
}
`);

    expect(result.diagnostics).toEqual([]);
    expect(result.javascript).toContain("exports.run = run");
  });

  it("typechecks generated code against TypeScript standard library declarations", () => {
    const result = compileOfficeCode(`
export async function run(_ctx: ExcelRuntimeContext) {
  const entries: Array<[string, number]> = [["value", 1]];
  const values = new Map(entries);
  const object: Record<string, number> = Object.fromEntries(values);
  return {
    message: "Used standard library APIs.",
    hasValue: [1, 2, 3].includes(object.value),
  };
}
`);

    expect(result.diagnostics).toEqual([]);
    expect(result.javascript).toContain("exports.run = run");
  });

  it("allows implicit any in generated helper parameters", () => {
    const result = compileOfficeCode(`
export async function run(ctx: ExcelRuntimeContext) {
  function summarize(value) {
    return String(value);
  }

  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  return {
    message: summarize("Read active worksheet."),
    sheet,
  };
}
`);

    expect(result.diagnostics).toEqual([]);
    expect(result.javascript).toContain("exports.run = run");
  });

  it("returns semantic diagnostics for nonexistent OfficeJS APIs", () => {
    const result = compileOfficeCode(`
export async function run(ctx: ExcelRuntimeContext) {
  ctx.workbook.worksheets.notARealApi();
}
`);

    expect(result.javascript).toBe("");
    expect(result.diagnostics.some((diagnostic) => diagnostic.message.includes("notARealApi"))).toBe(true);
  });

  it("rejects imports before emit", () => {
    const result = compileOfficeCode(`
import { something } from "somewhere";
export async function run(ctx: ExcelRuntimeContext) {
  return { message: "unused", something };
}
`);

    expect(result.javascript).toBe("");
    expect(result.diagnostics).toMatchObject([
      {
        severity: "error",
        code: "unsupported-import",
      },
    ]);
  });
});

describe("unsafeEvaluateOfficeCode", () => {
  it("loads an exported run function from compiled JavaScript", async () => {
    const compiled = compileOfficeCode(`
export async function run(ctx: ExcelRuntimeContext) {
  ctx.log("called");
  return { message: "Ran." };
}
`);

    const module = await unsafeEvaluateOfficeCode(compiled.javascript);

    expect(module.run).toEqual(expect.any(Function));
  });
});

describe("executeOfficeCode", () => {
  it("runs code, captures logs, and returns JSON data", async () => {
    const result = await executeOfficeCode(
      `
export async function run(ctx: ExcelRuntimeContext) {
  ctx.log("read range", { address: "A1" });
  return {
    message: "Read A1.",
    values: [[1]],
  };
}
`,
      { excelRunner: mockExcelRunner },
    );

    expect(result.content).toBe(`OfficeJS code executed successfully.

Returned data:
{
  "message": "Read A1.",
  "values": [
    [
      1
    ]
  ]
}`);
    expect(result.details).toMatchObject({
      status: "success",
      diagnostics: [],
      logs: [
        {
          message: "read range",
          details: { address: "A1" },
        },
      ],
      returnValue: {
        message: "Read A1.",
        values: [[1]],
      },
    });
  });

  it("throws a clear error when run is missing", async () => {
    await expect(
      executeOfficeCode("export const value = 1;", {
        excelRunner: mockExcelRunner,
      }),
    ).rejects.toThrow(OfficeCodeExecutionError);
    await expect(
      executeOfficeCode("export const value = 1;", {
        excelRunner: mockExcelRunner,
      }),
    ).rejects.toThrow("run(ctx)");
  });

  it("wraps runtime errors with execution details", async () => {
    await expect(
      executeOfficeCode(
        `
export async function run(_ctx: ExcelRuntimeContext) {
  throw new Error("boom");
}
`,
        { excelRunner: mockExcelRunner },
      ),
    ).rejects.toMatchObject({
      name: "OfficeCodeExecutionError",
      details: {
        status: "error",
      },
    });
  });
});
