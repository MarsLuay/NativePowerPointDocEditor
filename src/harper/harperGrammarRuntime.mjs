/**
 * Harper grammar runtime boundary.
 *
 * Production interactive linting uses WorkerLinter with the slim WebAssembly
 * binary. The materialized file is larger than NPDE's 5 MB per-runtime-file
 * limit, and embedding its gzip payload would also push main.js over its
 * budget, so this module is not imported by the plugin bundle. Callers pass
 * createHarperWorkerLinter into HarperGrammarService only after a packaging
 * check accepts the artifact.
 */
export async function createHarperWorkerLinter() {
	const harper = await import('harper.js');
	const { slimBinary } = await import('harper.js/slimBinary');
	const linter = new harper.WorkerLinter({
		binary: slimBinary,
		dialect: harper.Dialect.American,
	});
	return linter;
}
