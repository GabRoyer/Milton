import ts from "typescript";

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

/** Severity levels reported by the Office code compiler. */
export type OfficeCodeDiagnosticSeverity = "error" | "warning" | "suggestion" | "message";

/** Compiler diagnostic formatted for model and UI consumption. */
export interface OfficeCodeDiagnostic {
  /** Diagnostic severity. */
  severity: OfficeCodeDiagnosticSeverity;
  /** Human-readable diagnostic message. */
  message: string;
  /** One-based source line when available. */
  line?: number;
  /** One-based source column when available. */
  column?: number;
  /** TypeScript or Milton-specific diagnostic code. */
  code?: number | string;
}

/** JavaScript output and diagnostics from compiling generated Office code. */
export interface OfficeCodeCompileResult {
  /** Evaluator-compatible JavaScript, empty when compilation fails. */
  javascript: string;
  /** Diagnostics collected during compile and emit. */
  diagnostics: OfficeCodeDiagnostic[];
}

/** Typechecks generated TypeScript and emits evaluator-compatible JavaScript. */
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

/** Creates an in-memory TypeScript compiler host for generated Office code. */
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

/** Finds unsupported imports before TypeScript tries module resolution. */
function findUnsupportedImportDiagnostics(source: string): OfficeCodeDiagnostic[] {
  const sourceFile = ts.createSourceFile(ENTRY_FILE, source, ts.ScriptTarget.Latest, true);
  const diagnostics: OfficeCodeDiagnostic[] = [];

  /** Walks the generated source looking for import-like syntax. */
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

/** Converts a TypeScript diagnostic into Milton's tool-result diagnostic shape. */
function toOfficeDiagnostic(diagnostic: ts.Diagnostic): OfficeCodeDiagnostic {
  return {
    severity: toSeverity(diagnostic.category),
    message: ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n"),
    code: diagnostic.code,
    ...(diagnostic.file && diagnostic.start !== undefined ? getLocation(diagnostic.file, diagnostic.start) : {}),
  };
}

/** Converts a zero-based TypeScript source position into one-based line and column values. */
function getLocation(sourceFile: ts.SourceFile, position: number): Pick<OfficeCodeDiagnostic, "line" | "column"> {
  const location = sourceFile.getLineAndCharacterOfPosition(position);

  return {
    line: location.line + 1,
    column: location.character + 1,
  };
}

/** Maps TypeScript diagnostic categories to serialized severities. */
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

/** Runtime context passed to generated Excel OfficeJS code. */
interface ExcelRuntimeContext {
  /** Raw Excel request context that owns all OfficeJS objects used during the run. */
  context: Excel.RequestContext;
  /** Current workbook convenience alias from the active request context. */
  workbook: Excel.Workbook;
  /** Synchronizes queued OfficeJS loads and mutations with Excel. */
  sync(): Promise<void>;
  /** Captures structured execution logs for the tool result. */
  log(message: string, details?: unknown): void;
  /** Optional cancellation signal provided by the agent runtime. */
  signal?: AbortSignal;
}

declare namespace Excel {
  /** Runs a batch of OfficeJS commands against an Excel request context. */
  function run<T>(batch: (context: RequestContext) => Promise<T>): Promise<T>;

  /** Request-scoped root object for Excel OfficeJS operations. */
  interface RequestContext {
    /** Workbook associated with this request context. */
    workbook: Workbook;
    /** Synchronizes queued OfficeJS loads and mutations with Excel. */
    sync(): Promise<void>;
  }

  /** Excel workbook object exposed by OfficeJS. */
  interface Workbook {
    /** Collection of worksheets in the workbook. */
    worksheets: WorksheetCollection;
    /** Collection of workbook-level tables. */
    tables: TableCollection;
    /** Application-level Excel operations. */
    application: Application;
  }

  /** Excel application-level operations. */
  interface Application {
    /** Recalculates workbook formulas using the requested calculation mode. */
    calculate(calculationType: CalculationType): void;
  }

  type CalculationType = "Recalculate" | "Full" | "FullRebuild";

  /** Collection API for workbook worksheets. */
  interface WorksheetCollection {
    /** Gets the currently active worksheet. */
    getActiveWorksheet(): Worksheet;
    /** Gets a worksheet by name or id. */
    getItem(key: string): Worksheet;
    /** Queues worksheet collection properties for loading. */
    load(propertyNames?: string | string[]): void;
  }

  /** Excel worksheet object exposed by OfficeJS. */
  interface Worksheet {
    /** Worksheet name. */
    name: string;
    /** Tables contained in this worksheet. */
    tables: TableCollection;
    /** Gets a range by address, or the full worksheet range when omitted. */
    getRange(address?: string): Range;
    /** Gets the used range in the worksheet. */
    getUsedRange(valuesOnly?: boolean): Range;
    /** Queues worksheet properties for loading. */
    load(propertyNames?: string | string[]): void;
  }

  /** Excel range object exposed by OfficeJS. */
  interface Range {
    /** Fully-qualified range address. */
    address: string;
    /** Raw cell values in row-major order. */
    values: ExcelRangeValues;
    /** Display text in row-major order. */
    text: string[][];
    /** Formula values in row-major order. */
    formulas: ExcelRangeValues;
    /** Number formats in row-major order. */
    numberFormat: string[][];
    /** Number of rows in the range. */
    rowCount: number;
    /** Number of columns in the range. */
    columnCount: number;
    /** Formatting controls for the range. */
    format: RangeFormat;
    /** Gets a cell within the range by zero-based row and column. */
    getCell(row: number, column: number): Range;
    /** Gets a column within the range by zero-based index. */
    getColumn(column: number): Range;
    /** Gets a row within the range by zero-based index. */
    getRow(row: number): Range;
    /** Queues range properties for loading. */
    load(propertyNames?: string | string[]): void;
  }

  /** Formatting controls for an Excel range. */
  interface RangeFormat {
    /** Fill formatting controls. */
    fill: RangeFill;
    /** Font formatting controls. */
    font: RangeFont;
    /** Autofits range columns. */
    autofitColumns(): void;
    /** Autofits range rows. */
    autofitRows(): void;
  }

  /** Fill formatting controls for an Excel range. */
  interface RangeFill {
    /** Fill color as an HTML color string. */
    color: string;
  }

  /** Font formatting controls for an Excel range. */
  interface RangeFont {
    /** Whether the font is bold. */
    bold: boolean;
    /** Font color as an HTML color string. */
    color: string;
    /** Whether the font is italic. */
    italic: boolean;
    /** Font family name. */
    name: string;
    /** Font size in points. */
    size: number;
  }

  /** Collection API for Excel tables. */
  interface TableCollection {
    /** Adds a table at an address or range. */
    add(address: string | Range, hasHeaders: boolean): Table;
    /** Gets a table by name or id. */
    getItem(key: string): Table;
  }

  /** Excel table object exposed by OfficeJS. */
  interface Table {
    /** Table name. */
    name: string;
    /** Gets the range occupied by the table. */
    getRange(): Range;
    /** Queues table properties for loading. */
    load(propertyNames?: string | string[]): void;
  }
}
`;
