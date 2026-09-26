import { readdir } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const testsDirectory = path.join(projectRoot, "tests");
const matrixTests = (await readdir(testsDirectory))
  .filter((file) => /^matrix-pptx-.*\.test\.mjs$/.test(file))
  .sort();

if (matrixTests.length === 0) {
  console.error("No PPTX action matrix tests found in tests/.");
  process.exitCode = 1;
} else {
  console.log(`Running ${matrixTests.length} PPTX action matrix test files.`);
  const child = spawn(
    process.execPath,
    ["--test", ...matrixTests.map((file) => path.join("tests", file))],
    { cwd: projectRoot, stdio: "inherit" },
  );

  const exitCode = await new Promise((resolve, reject) => {
    child.once("error", reject);
    child.once("exit", (code, signal) => {
      if (signal) {
        reject(new Error(`PPTX action matrix tests terminated by ${signal}.`));
        return;
      }
      resolve(code ?? 1);
    });
  }).catch((error) => {
    console.error(error.message);
    return 1;
  });

  if (exitCode !== 0) process.exitCode = exitCode;
}
