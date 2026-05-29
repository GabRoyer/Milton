/**
 * Minimal standard library declarations injected into the runtime compiler.
 *
 * The runtime compiler sets `noLib: true` so model-generated programs only see
 * the JavaScript globals Milton intentionally provides for the first milestone.
 */
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
