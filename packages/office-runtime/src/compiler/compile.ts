import ts from "typescript";
import miltonRuntimeDeclarations from "./built-in-types/milton-runtime.d.ts?raw";
import officeJsDeclarations from "../../node_modules/@types/office-js/index.d.ts?raw";

const ENTRY_FILE = "/entry.ts";
const RUNTIME_FILE = "/milton-office-runtime.d.ts";
const OFFICE_JS_FILE = "/node_modules/@types/office-js/index.d.ts";
const TYPESCRIPT_LIB_DIR = "/node_modules/typescript/lib";
const TYPESCRIPT_LIB_FILE = "lib.es2020.d.ts";

const typescriptEsLibDeclarationModules = import.meta.glob("../../node_modules/typescript/lib/lib.es*.d.ts", {
  eager: true,
  import: "default",
  query: "?raw",
}) as Record<string, string>;

const typescriptDecoratorLibDeclarationModules = import.meta.glob("../../node_modules/typescript/lib/lib.decorators*.d.ts", {
  eager: true,
  import: "default",
  query: "?raw",
}) as Record<string, string>;

const compilerOptions: ts.CompilerOptions = {
  target: ts.ScriptTarget.ES2020,
  module: ts.ModuleKind.CommonJS,
  jsx: ts.JsxEmit.None,
  noEmitOnError: true,
  strict: true,
  noImplicitAny: false,
  skipLibCheck: true,
  lib: [TYPESCRIPT_LIB_FILE],
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
    [RUNTIME_FILE, miltonRuntimeDeclarations],
    [OFFICE_JS_FILE, officeJsDeclarations],
    ...getTypeScriptLibFiles(),
  ]);
  let javascript = "";

  const host = createCompilerHost(files, (fileName, text) => {
    if (fileName.endsWith(".js")) {
      javascript = text;
    }
  });
  const program = ts.createProgram([ENTRY_FILE, RUNTIME_FILE, OFFICE_JS_FILE], compilerOptions, host);
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

/** Creates a TypeScript compiler host backed by virtual source files. */
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
    getDefaultLibFileName: () => TYPESCRIPT_LIB_FILE,
    getDefaultLibLocation: () => TYPESCRIPT_LIB_DIR,
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

/** Maps TypeScript's packaged standard library declarations into virtual compiler files. */
function getTypeScriptLibFiles(): Array<[string, string]> {
  return Object.entries({
    ...typescriptEsLibDeclarationModules,
    ...typescriptDecoratorLibDeclarationModules,
  }).map(([fileName, source]) => [
    `${TYPESCRIPT_LIB_DIR}/${getBaseName(fileName)}`,
    source,
  ]);
}

/** Returns the last path segment for virtual declaration file names. */
function getBaseName(fileName: string): string {
  return fileName.slice(fileName.lastIndexOf("/") + 1);
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
