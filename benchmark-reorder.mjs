import { performance } from 'perf_hooks';

// Simulate the logic in PresentationEngine.ts

function createData(size) {
  const shapes = [];
  for (let i = 0; i < size; i++) {
    shapes.push({ id: i });
  }
  return shapes;
}

const N = 10000;
const shapes = createData(N);
const shapeIndexes = [];
for (let i = 0; i < N; i += 2) {
  shapeIndexes.push(i);
}

const selected = new Set(
  shapeIndexes.map(index => shapes[index])
);

// Original approach
const startOriginal = performance.now();
for (let iter = 0; iter < 100; iter++) {
  const ordered = shapes.filter((element) => selected.has(element));
  const sourceIndexes = ordered.map((element) => shapes.indexOf(element));
}
const endOriginal = performance.now();
console.log(`Original: ${(endOriginal - startOriginal).toFixed(2)} ms`);

// Optimized approach
const startOptimized = performance.now();
for (let iter = 0; iter < 100; iter++) {
  const ordered = [];
  const sourceIndexes = [];
  for (let i = 0; i < shapes.length; i++) {
    const element = shapes[i];
    if (selected.has(element)) {
      ordered.push(element);
      sourceIndexes.push(i);
    }
  }
}
const endOptimized = performance.now();
console.log(`Optimized: ${(endOptimized - startOptimized).toFixed(2)} ms`);
