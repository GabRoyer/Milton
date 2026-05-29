/**
 * Ambient type support for bundler raw imports used by the compiler worker.
 *
 * Vite turns `?raw` imports into string contents at bundle time. The package
 * build still runs plain `tsc --noEmit`, so it needs this declaration to accept
 * the OfficeJS and built-in type declaration imports in `compile.ts`.
 */
declare module "*?raw" {
  /** Imported file contents emitted by Vite's raw import handling. */
  const content: string;
  export default content;
}
