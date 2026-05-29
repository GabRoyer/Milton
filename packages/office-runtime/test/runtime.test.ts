import { describe, expect, it } from "vitest";
import { compileOfficeCode } from "../src/compiler/compile";
import { executeOfficeCode, OfficeCodeExecutionError, type ExcelRunner } from "../src/execution";
import { unsafeEvaluateOfficeCode } from "../src/evaluation/unsafe-evaluator";
import { createExecuteOfficeJsCodeTool, EXECUTE_OFFICEJS_CODE_TOOL_NAME } from "../src/tool";

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
      { compile: compileOfficeCode, excelRunner: mockExcelRunner },
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
        compile: compileOfficeCode,
        excelRunner: mockExcelRunner,
      }),
    ).rejects.toThrow(OfficeCodeExecutionError);
    await expect(
      executeOfficeCode("export const value = 1;", {
        compile: compileOfficeCode,
        excelRunner: mockExcelRunner,
      }),
    ).rejects.toThrow("run(ctx)");
  });

  it("formats compile errors with source location and structured details", async () => {
    try {
      await executeOfficeCode(
        `
export async function run(ctx: ExcelRuntimeContext) {
  ctx.workbook.worksheets.notARealApi();
}
`,
        { compile: compileOfficeCode, excelRunner: mockExcelRunner },
      );
      throw new Error("Expected executeOfficeCode to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(OfficeCodeExecutionError);
      expect(error).toMatchObject({
        message: expect.stringContaining("Line "),
        details: {
          status: "error",
          code: expect.stringContaining("notARealApi"),
        },
      });
      expect((error as OfficeCodeExecutionError).message).toContain("Column ");
      expect((error as OfficeCodeExecutionError).message).toMatch(/\[\d+\]/);
      expect((error as OfficeCodeExecutionError).details.diagnostics).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            line: expect.any(Number),
            column: expect.any(Number),
          }),
        ]),
      );
    }
  });

  it("wraps runtime errors with execution details", async () => {
    await expect(
      executeOfficeCode(
        `
export async function run(_ctx: ExcelRuntimeContext) {
  throw new Error("boom");
}
`,
        { compile: compileOfficeCode, excelRunner: mockExcelRunner },
      ),
    ).rejects.toMatchObject({
      name: "OfficeCodeExecutionError",
      details: {
        status: "error",
      },
    });
  });
});

describe("createExecuteOfficeJsCodeTool", () => {
  it("creates a sequential agent tool that returns execution content and details", async () => {
    const tool = createExecuteOfficeJsCodeTool({
      compile: compileOfficeCode,
      excelRunner: mockExcelRunner,
    });

    const result = await tool.execute("tool-call-1", {
      code: `
export async function run(_ctx: ExcelRuntimeContext) {
  return { message: "Tool ran." };
}
`,
    });

    expect(tool.name).toBe(EXECUTE_OFFICEJS_CODE_TOOL_NAME);
    expect(tool.executionMode).toBe("sequential");
    expect(result.content).toEqual([
      {
        type: "text",
        text: `OfficeJS code executed successfully.

Returned data:
{
  "message": "Tool ran."
}`,
      },
    ]);
    expect(result.details).toMatchObject({
      status: "success",
      code: expect.stringContaining("Tool ran"),
      returnValue: {
        message: "Tool ran.",
      },
    });
  });

  it("streams OfficeJS logs through tool updates", async () => {
    const tool = createExecuteOfficeJsCodeTool({
      compile: compileOfficeCode,
      excelRunner: mockExcelRunner,
    });
    const updates: unknown[] = [];

    await tool.execute(
      "tool-call-1",
      {
        code: `
export async function run(ctx: ExcelRuntimeContext) {
  ctx.log("read range", { address: "A1" });
  return { message: "Tool ran." };
}
`,
      },
      undefined,
      (update) => updates.push(update),
    );

    expect(updates).toMatchObject([
      {
        content: [{ type: "text", text: "OfficeJS log: read range" }],
        details: {
          status: "running",
          code: expect.stringContaining("ctx.log"),
          latestLog: {
            message: "read range",
            details: { address: "A1" },
          },
        },
      },
    ]);
  });
});
