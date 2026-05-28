import ts from "typescript";
import type { OfficeCodeCompileResult, OfficeCodeDiagnostic } from "./types";

const ENTRY_FILE = "/entry.ts";
const LIB_FILE = "/milton-lib.d.ts";
const RUNTIME_FILE = "/milton-office-runtime.d.ts";

const compilerOptions: ts.CompilerOptions = {
  target: ts.ScriptTarget.ES2020,
  module: ts.ModuleKind.CommonJS,
  jsx: ts.JsxEmit.None,
  noEmitOnError: true,
  strict: true,
  skipLibCheck: true,
  noLib: true,
};

export function compileOfficeCode(source: string): OfficeCodeCompileResult {
  const importDiagnostics = findUnsupportedImportDiagnostics(source);

  if (importDiagnostics.length > 0) {
    return {
      javascript: "",
      diagnostics: importDiagnostics,
    };
  }

  const files = new Map<string, string>([
    [ENTRY_FILE, source],
    [LIB_FILE, STANDARD_DECLARATIONS],
    [RUNTIME_FILE, OFFICE_RUNTIME_DECLARATIONS],
  ]);
  let javascript = "";

  const host = createCompilerHost(files, (fileName, text) => {
    if (fileName.endsWith(".js")) {
      javascript = text;
    }
  });
  const program = ts.createProgram([ENTRY_FILE, LIB_FILE, RUNTIME_FILE], compilerOptions, host);
  const diagnostics = ts.getPreEmitDiagnostics(program).map(toOfficeDiagnostic);

  if (diagnostics.some((diagnostic) => diagnostic.severity === "error")) {
    return {
      javascript: "",
      diagnostics,
    };
  }

  const emitResult = program.emit();
  const emitDiagnostics = emitResult.diagnostics.map(toOfficeDiagnostic);

  if (emitResult.emitSkipped || emitDiagnostics.some((diagnostic) => diagnostic.severity === "error")) {
    return {
      javascript: "",
      diagnostics: diagnostics.concat(emitDiagnostics),
    };
  }

  return {
    javascript,
    diagnostics: diagnostics.concat(emitDiagnostics),
  };
}

function createCompilerHost(files: Map<string, string>, writeFile: ts.WriteFileCallback): ts.CompilerHost {
  return {
    getSourceFile(fileName, languageVersion) {
      const source = files.get(fileName);

      if (source === undefined) {
        return undefined;
      }

      return ts.createSourceFile(fileName, source, languageVersion, true);
    },
    writeFile,
    getDefaultLibFileName: () => LIB_FILE,
    useCaseSensitiveFileNames: () => true,
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: () => "/",
    getNewLine: () => "\n",
    fileExists: (fileName) => files.has(fileName),
    readFile: (fileName) => files.get(fileName),
    directoryExists: () => true,
    getDirectories: () => [],
    resolveModuleNames: (moduleNames) => moduleNames.map(() => undefined),
  };
}

function findUnsupportedImportDiagnostics(source: string): OfficeCodeDiagnostic[] {
  const sourceFile = ts.createSourceFile(ENTRY_FILE, source, ts.ScriptTarget.Latest, true);
  const diagnostics: OfficeCodeDiagnostic[] = [];

  function visit(node: ts.Node) {
    if (
      ts.isImportDeclaration(node) ||
      ts.isImportEqualsDeclaration(node) ||
      (ts.isExportDeclaration(node) && node.moduleSpecifier) ||
      (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword)
    ) {
      diagnostics.push({
        severity: "error",
        message: "Imports are not supported by execute_officejs_code in this milestone.",
        ...getLocation(sourceFile, node.getStart(sourceFile)),
        code: "unsupported-import",
      });
      return;
    }

    ts.forEachChild(node, visit);
  }

  visit(sourceFile);
  return diagnostics;
}

function toOfficeDiagnostic(diagnostic: ts.Diagnostic): OfficeCodeDiagnostic {
  return {
    severity: toSeverity(diagnostic.category),
    message: ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n"),
    code: diagnostic.code,
    ...(diagnostic.file && diagnostic.start !== undefined ? getLocation(diagnostic.file, diagnostic.start) : {}),
  };
}

function getLocation(sourceFile: ts.SourceFile, position: number): Pick<OfficeCodeDiagnostic, "line" | "column"> {
  const location = sourceFile.getLineAndCharacterOfPosition(position);

  return {
    line: location.line + 1,
    column: location.character + 1,
  };
}

function toSeverity(category: ts.DiagnosticCategory): OfficeCodeDiagnostic["severity"] {
  switch (category) {
    case ts.DiagnosticCategory.Error:
      return "error";
    case ts.DiagnosticCategory.Warning:
      return "warning";
    case ts.DiagnosticCategory.Suggestion:
      return "suggestion";
    case ts.DiagnosticCategory.Message:
      return "message";
  }
}

const STANDARD_DECLARATIONS = `
interface Array<T> {
  readonly length: number;
  [n: number]: T;
  push(...items: T[]): number;
  map<U>(callbackfn: (value: T, index: number, array: T[]) => U): U[];
  filter(callbackfn: (value: T, index: number, array: T[]) => unknown): T[];
  forEach(callbackfn: (value: T, index: number, array: T[]) => void): void;
}
interface ReadonlyArray<T> {
  readonly length: number;
  readonly [n: number]: T;
}
interface Boolean {}
interface CallableFunction extends Function {}
interface Function {}
interface IArguments {}
interface NewableFunction extends Function {}
interface Number {}
interface Object {}
interface RegExp {}
interface String {}
interface Error {
  message: string;
}
interface ErrorConstructor {
  new (message?: string): Error;
}
declare var Error: ErrorConstructor;
interface PromiseLike<T> {
  then<TResult1 = T, TResult2 = never>(
    onfulfilled?: ((value: T) => TResult1 | PromiseLike<TResult1>) | null,
    onrejected?: ((reason: any) => TResult2 | PromiseLike<TResult2>) | null
  ): PromiseLike<TResult1 | TResult2>;
}
interface Promise<T> extends PromiseLike<T> {
  catch<TResult = never>(
    onrejected?: ((reason: any) => TResult | PromiseLike<TResult>) | null
  ): Promise<T | TResult>;
}
interface PromiseConstructor {
  resolve<T>(value: T | PromiseLike<T>): Promise<T>;
  reject<T = never>(reason?: any): Promise<T>;
}
declare var Promise: PromiseConstructor;
interface AbortSignal {
  readonly aborted: boolean;
  readonly reason?: unknown;
}
`;

const OFFICE_RUNTIME_DECLARATIONS = `
type ExcelCellValue = string | number | boolean | null;
type ExcelRangeValues = ExcelCellValue[][];

interface ExcelRuntimeContext {
  context: Excel.RequestContext;
  workbook: Excel.Workbook;
  sync(): Promise<void>;
  log(message: string, details?: unknown): void;
  signal?: AbortSignal;
}

declare namespace Excel {
  function run<T>(batch: (context: RequestContext) => Promise<T>): Promise<T>;

  interface RequestContext {
    workbook: Workbook;
    sync(): Promise<void>;
  }

  interface Workbook {
    worksheets: WorksheetCollection;
    tables: TableCollection;
    application: Application;
  }

  interface Application {
    calculate(calculationType: CalculationType): void;
  }

  type CalculationType = "Recalculate" | "Full" | "FullRebuild";

  interface WorksheetCollection {
    getActiveWorksheet(): Worksheet;
    getItem(key: string): Worksheet;
    load(propertyNames?: string | string[]): void;
  }

  interface Worksheet {
    name: string;
    tables: TableCollection;
    getRange(address?: string): Range;
    getUsedRange(valuesOnly?: boolean): Range;
    load(propertyNames?: string | string[]): void;
  }

  interface Range {
    address: string;
    values: ExcelRangeValues;
    text: string[][];
    formulas: ExcelRangeValues;
    numberFormat: string[][];
    rowCount: number;
    columnCount: number;
    format: RangeFormat;
    getCell(row: number, column: number): Range;
    getColumn(column: number): Range;
    getRow(row: number): Range;
    load(propertyNames?: string | string[]): void;
  }

  interface RangeFormat {
    fill: RangeFill;
    font: RangeFont;
    autofitColumns(): void;
    autofitRows(): void;
  }

  interface RangeFill {
    color: string;
  }

  interface RangeFont {
    bold: boolean;
    color: string;
    italic: boolean;
    name: string;
    size: number;
  }

  interface TableCollection {
    add(address: string | Range, hasHeaders: boolean): Table;
    getItem(key: string): Table;
  }

  interface Table {
    name: string;
    getRange(): Range;
    load(propertyNames?: string | string[]): void;
  }
}
`;
