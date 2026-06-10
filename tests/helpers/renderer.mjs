import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PptxRenderer } from "pptx-svg";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const fixtureDirectory = path.join(projectRoot, "tests/fixtures/decks");
const wasmPath = path.join(projectRoot, "node_modules/pptx-svg/dist/main.wasm");

let wasmPromise;

export function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

export async function readDeck(name) {
  return readFile(path.join(fixtureDirectory, name));
}

export async function createRenderer(bytes) {
  wasmPromise ??= readFile(wasmPath);
  const wasm = await wasmPromise;
  const renderer = new PptxRenderer({
    logLevel: "error",
    measureText: (text, _fontFace, fontSize) => text.length * fontSize * 0.55,
  });
  await renderer.init(wasm);
  await renderer.loadPptx(toArrayBuffer(bytes));
  return renderer;
}
