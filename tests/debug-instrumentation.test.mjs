import assert from "node:assert/strict";
import { test } from "node:test";
import { createRequire } from "node:module";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
let modulePromise;
async function loadModule() {
  modulePromise ??= bundleSource("src/debugInstrumentation.ts", "debug-instrumentation.cjs").then((outfile) => require(outfile));
  return modulePromise;
}

test("observer amplification trace aggregates cheap callbacks per user action", async () => {
  const { createObserverAmplificationTrace } = await loadModule();
  const trace = createObserverAmplificationTrace();
  trace.begin("keystroke");
  for (let index = 0; index < 25; index += 1) trace.record(index % 2 ? "host-editor" : "plugin", 1, 0.1);
  const record = trace.end();
  assert.equal(record.action, "keystroke");
  assert.equal(record.callbackCount, 25);
  assert.equal(record.mutationCount, 25);
  assert.equal(record.sources.plugin.callbacks, 13);
  assert.equal(record.sources["host-editor"].callbacks, 12);
  assert.equal(trace.end(), null);
});

test("observer amplification trace bounds source cardinality and preserves cumulative work", async () => {
  const { createObserverAmplificationTrace } = await loadModule();
  const trace = createObserverAmplificationTrace();
  trace.begin("drag-burst");
  for (let index = 0; index < 20; index += 1) trace.record(`source-${index}`, 3, 2);
  const record = trace.end();
  assert.equal(Object.keys(record.sources).length, 8);
  assert.equal(record.mutationCount, 60);
  assert.equal(record.callbackWorkMs, 40);
  assert.ok(record.durationMs >= 0);
});
