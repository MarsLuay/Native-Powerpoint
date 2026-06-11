import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import Module, { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
let tempDirectoryPromise;
let packageModulePromise;
let presentationEngineModulePromise;
let shapeClipboardModulePromise;
let viewModulePromise;

globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

const inlinePptxSvgWasmPlugin = {
  name: "inline-pptx-svg-wasm",
  setup(buildContext) {
    buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async ({ path: modulePath }) => {
      const contents = (await readFile(modulePath, "utf8")).replace(
        "const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;",
        "const DEFAULT_WASM_URL = undefined;",
      );
      return { contents, loader: "js" };
    });
  },
};

function getTempDirectory() {
  tempDirectoryPromise ??= mkdtemp(path.join(tmpdir(), "native-powerpoint-tests-"));
  return tempDirectoryPromise;
}

async function bundleSource(entry, outputName, external = []) {
  const outputDirectory = await getTempDirectory();
  const outfile = path.join(outputDirectory, outputName);
  await build({
    entryPoints: [path.join(projectRoot, entry)],
    bundle: true,
    external,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  return outfile;
}

export function loadPowerPointPackageModule() {
  packageModulePromise ??= bundleSource("src/PowerPointPackage.ts", "powerpoint-package.cjs").then(
    (outfile) => require(outfile),
  );
  return packageModulePromise;
}

export function loadPresentationEngineModule() {
  presentationEngineModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const outfile = path.join(outputDirectory, "presentation-engine.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/PresentationEngine.ts")],
      bundle: true,
      format: "cjs",
      loader: { ".wasm": "binary" },
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [inlinePptxSvgWasmPlugin],
      target: "node22",
    });
    return require(outfile);
  })();
  return presentationEngineModulePromise;
}

export function loadShapeClipboardModule() {
  shapeClipboardModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const outfile = path.join(outputDirectory, "shape-clipboard.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/ShapeClipboard.ts")],
      bundle: true,
      format: "cjs",
      loader: { ".wasm": "binary" },
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [inlinePptxSvgWasmPlugin],
      target: "node22",
    });
    return require(outfile);
  })();
  return shapeClipboardModulePromise;
}

export function loadNativePowerPointViewModule() {
  viewModulePromise ??= bundleSource(
    "src/NativePowerPointView.ts",
    "native-powerpoint-view.cjs",
    ["obsidian", "pptx-svg", "pptx-svg/wasm"],
  ).then((outfile) => {
    const notices = [];
    const originalLoad = Module._load;
    globalThis.activeDocument ??= {
      activeElement: null,
      addEventListener() {},
      body: { classList: { add() {}, remove() {}, toggle() {} } },
      removeEventListener() {},
    };

    class Notice {
      constructor(message, duration) {
        notices.push({ message, duration });
      }
    }

    class FileView {
      constructor(leaf = {}) {
        this.leaf = leaf;
        this.app = leaf.app ?? { vault: {} };
        this.file = null;
        this.contentEl = leaf.contentEl ?? createElementStub();
        this.containerEl = leaf.containerEl ?? { isShown: () => true };
      }

      registerDomEvent() {}
    }

    class Menu {
      addItem(callback) {
        callback?.({
          onClick() { return this; },
          setIcon() { return this; },
          setTitle() { return this; },
        });
        return this;
      }

      addSeparator() {
        return this;
      }

      showAtMouseEvent() {}
      showAtPosition() {}
    }

    class Modal {
      constructor(app) {
        this.app = app;
        this.contentEl = createElementStub();
        this.modalEl = createElementStub();
      }

      close() {}
      open() {}
    }

    Module._load = function load(request, parent, isMain) {
      if (request === "obsidian") {
        return {
          activeDocument: {
            activeElement: null,
            addEventListener() {},
            body: { classList: { add() {}, remove() {}, toggle() {} } },
            removeEventListener() {},
          },
          FileView,
          Menu,
          Modal,
          Notice,
          Platform: { isDesktop: true, isMobile: false },
          normalizePath: (value) => value.replace(/\\/g, "/").replace(/\/{2,}/g, "/"),
          setIcon: () => undefined,
        };
      }

      if (request === "pptx-svg") {
        return {
          PptxRenderer: class {},
          degreesToOoxml: (value) => value,
          emuToPx: (value) => value,
          getAllShapes: () => [],
          getShapeTransform: () => ({ x: 0, y: 0, cx: 1, cy: 1, rot: 0 }),
          getSlideScale: () => 1,
          ooxmlToDegrees: (value) => value,
          pxToEmu: (value) => value,
        };
      }

      if (request === "pptx-svg/wasm") {
        return new Uint8Array();
      }

      return originalLoad.call(this, request, parent, isMain);
    };

    try {
      return { ...require(outfile), notices };
    } finally {
      Module._load = originalLoad;
    }
  });

  return viewModulePromise;
}

function createElementStub() {
  return {
    addClass() {},
    empty() {},
    removeClass() {},
  };
}
