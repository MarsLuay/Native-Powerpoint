// Ambient types for the generated pure-JS PPTX engine fallback
// (src/vendor/pptx-js-engine.mjs). The module is loaded lazily and only when the
// runtime lacks WebAssembly GC. Its exports mirror the Wasm module's exports, so
// a loose signature is sufficient — the renderer consumes them via its own typed
// `exports` surface.
declare module '*/pptx-js-engine.mjs' {
  export type PptxJsEngineExports = Record<string, (...args: never[]) => unknown>;
  export function createPptxJsEngine(): PptxJsEngineExports;
}
