import assert from "node:assert/strict";
import { Buffer } from "node:buffer";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildStoredZip, createDeck, createDeckEntries } from "../helpers/fixture-builder.mjs";

const fixtureDirectory = path.join(path.dirname(fileURLToPath(import.meta.url)), "decks");
const mode = process.argv[2] ?? "--check";

const fixtures = new Map([
  ["features.pptx", createDeck({ format: "pptx" })],
  ["features.ppsx", createDeck({ format: "ppsx" })],
  ["features.potx", createDeck({ format: "potx" })],
  ["simple-edit.pptx", createDeck({ format: "pptx", richFirstSlide: false })],
  ["macro-view-only.pptm", createDeck({ format: "pptm" })],
  ["macro-view-only.ppsm", createDeck({ format: "ppsm" })],
  ["macro-view-only.potm", createDeck({ format: "potm" })],
  ["large-deck.pptx", createDeck({ format: "pptx", slideCount: 160 })],
]);

const featureEntries = createDeckEntries({ format: "pptx" });
const featureDeck = fixtures.get("features.pptx");
fixtures.set("malformed-random.pptx", new TextEncoder().encode("This is intentionally not a ZIP archive.\n"));
fixtures.set("malformed-truncated.pptx", featureDeck.slice(0, Math.floor(featureDeck.byteLength / 2)));
fixtures.set(
  "malformed-unsafe-path.pptx",
  buildStoredZip([...featureEntries, { name: "../escape.xml", data: new TextEncoder().encode("<escape/>") }]),
);
fixtures.set(
  "malformed-duplicate-entry.pptx",
  buildStoredZip([
    ...featureEntries,
    {
      name: "ppt/presentation.xml",
      data: new TextEncoder().encode('<?xml version="1.0"?><duplicate/>'),
    },
  ]),
);

async function writeFixtures() {
  await mkdir(fixtureDirectory, { recursive: true });
  for (const [name, bytes] of fixtures) {
    await writeFile(path.join(fixtureDirectory, name), bytes);
  }
  console.log(`Wrote ${fixtures.size} deterministic Native PowerPoint fixtures.`);
}

async function checkFixtures() {
  for (const [name, expected] of fixtures) {
    const actual = await readFile(path.join(fixtureDirectory, name));
    assert.deepEqual(
      actual,
      Buffer.from(expected),
      `${name} does not match the deterministic generator. Run npm run test:update-fixtures.`,
    );
  }
  console.log(`Verified ${fixtures.size} deterministic Native PowerPoint fixtures.`);
}

if (mode === "--write") {
  await writeFixtures();
} else if (mode === "--check") {
  await checkFixtures();
} else {
  throw new Error(`Unknown fixture generator mode: ${mode}`);
}
