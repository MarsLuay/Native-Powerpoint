// Browser/mobile shim for Node's "stream" builtin.
//
// Obsidian Mobile (iOS/Android) runs inside a WebView with no Node runtime, so
// require("stream") throws there. The only consumers in our dependency graph are:
//   - sax, which reads `.Stream` and falls back to a no-op when it is missing.
//   - jszip, which probes `.Readable` to feature-detect Node stream support and
//     leaves it disabled when absent (the Node stream adapters are never used).
// Exporting an empty module preserves that "no Node streams" behavior while
// keeping the bundle from referencing an unavailable builtin on mobile.
module.exports = {};
