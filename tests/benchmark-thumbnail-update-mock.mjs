import { performance } from 'node:perf_hooks';

class MockElement {
  constructor(className) {
    this.className = className;
    this.classList = new Set();
    this.children = [];
  }

  toggleClass(cls, cond) {
    if (cond) {
      this.classList.add(cls);
    } else {
      this.classList.delete(cls);
    }
  }

  querySelectorAll(selector) {
    const cls = selector.replace('.', '');
    const results = [];
    const walk = (node) => {
      if (node.className === cls) {
        results.push(node);
      }
      for (const child of node.children) {
        walk(child);
      }
    };
    walk(this);
    // querySelectorAll adds a forEach
    results.forEach = Array.prototype.forEach;
    return results;
  }

  appendChild(child) {
    this.children.push(child);
  }
}

const container = new MockElement('container');

// Create 500 items (typical large slide deck)
for (let i = 0; i < 500; i++) {
  const div = new MockElement('native-powerpoint-thumbnail');
  container.appendChild(div);
}

const host = {
  thumbnailContainer: container,
  currentSlide: 10
};
const selectedSlideIndices = new Set([10, 11, 12]);

function updateThumbnailActiveStateQSA() {
  const items = host.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail');
  items.forEach((item, index) => {
    item.toggleClass('active', index === host.currentSlide);
    item.toggleClass('is-selected', selectedSlideIndices.has(index));
  });
}

function updateThumbnailActiveStateChildren() {
  const items = host.thumbnailContainer.children;
  for (let index = 0; index < items.length; index++) {
    const item = items[index];
    item.toggleClass('active', index === host.currentSlide);
    item.toggleClass('is-selected', selectedSlideIndices.has(index));
  }
}

// Warmup
for (let i = 0; i < 1000; i++) {
  updateThumbnailActiveStateQSA();
  updateThumbnailActiveStateChildren();
}

let start, end;

start = performance.now();
for (let i = 0; i < 10000; i++) {
  updateThumbnailActiveStateQSA();
}
end = performance.now();
console.log('querySelectorAll:', (end - start).toFixed(2), 'ms');

start = performance.now();
for (let i = 0; i < 10000; i++) {
  updateThumbnailActiveStateChildren();
}
end = performance.now();
console.log('children:', (end - start).toFixed(2), 'ms');
