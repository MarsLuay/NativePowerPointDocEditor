import assert from "node:assert/strict";
import { test } from "node:test";
import { loadTextHaloModule } from "./helpers/load-plugin-modules.mjs";

function parseSvg(svgString) {
  const parser = new globalThis.DOMParser();
  return parser.parseFromString(svgString, "image/svg+xml").documentElement;
}

test("table text halos inherit background color via SVG structure", async () => {
  const { applyBackgroundAwareTextHalos, DEFAULT_TEXT_HALO_COLOR } = await loadTextHaloModule();

  // A basic table-like SVG where a text element follows a rectangle.
  const svg = parseSvg(`
    <svg>
      <g>
        <rect fill="#ff0000" width="10" height="10" />
        <text>Hello</text>
      </g>
      <g>
        <rect fill="#0000ff" width="10" height="10" />
        <g>
          <text>World</text>
        </g>
      </g>
      <g>
        <text>Default</text>
      </g>
    </svg>
  `);

  applyBackgroundAwareTextHalos(svg, "table");

  const texts = svg.getElementsByTagName("text");

  // text "Hello" after "red" rect
  assert.equal(texts[0].getAttribute("data-native-powerpoint-halo-color"), "rgb(255, 0, 0)");
  assert.ok(texts[0].getAttribute("style").includes("--native-powerpoint-text-halo: rgb(255, 0, 0)"));

  // text "World" in a group after a "blue" rect
  assert.equal(texts[1].getAttribute("data-native-powerpoint-halo-color"), "rgb(0, 0, 255)");

  // text "Default" without a preceding rect (inherits default)
  assert.equal(texts[2].getAttribute("data-native-powerpoint-halo-color"), DEFAULT_TEXT_HALO_COLOR);
});

test("chart text halos inherit background color of largest rect", async () => {
  const { applyBackgroundAwareTextHalos, DEFAULT_TEXT_HALO_COLOR } = await loadTextHaloModule();

  const svg = parseSvg(`
    <svg>
      <rect fill="#00ff00" width="100" height="100" />
      <rect fill="#ff0000" width="50" height="50" />
      <text>Chart Title</text>
      <g>
        <rect fill="#0000ff" width="10" height="10" />
        <text>Axis Label</text>
      </g>
    </svg>
  `);

  // #00ff00 (green) is 100x100 = 10000 area. Largest!
  applyBackgroundAwareTextHalos(svg, "chart");

  const texts = svg.getElementsByTagName("text");
  assert.equal(texts[0].getAttribute("data-native-powerpoint-halo-color"), "rgb(0, 255, 0)");
  assert.equal(texts[1].getAttribute("data-native-powerpoint-halo-color"), "rgb(0, 255, 0)");
});

test("color parsing handles opacity and fill-opacity", async () => {
  const { applyBackgroundAwareTextHalos } = await loadTextHaloModule();

  // Test compositeAgainstWhite
  // Red rect, 50% opacity -> composed with white -> rgb(255, 128, 128)
  const svg = parseSvg(`
    <svg>
      <rect fill="#ff0000" opacity="0.5" width="10" height="10" />
      <text>Test</text>

      <rect fill="#000000" fill-opacity="0.25" width="10" height="10" />
      <text>Test2</text>
    </svg>
  `);

  applyBackgroundAwareTextHalos(svg, "table");

  const texts = svg.getElementsByTagName("text");

  // red 0.5 composed with white:
  // red = 255 * 0.5 + 255 * 0.5 = 255
  // green = 0 * 0.5 + 255 * 0.5 = 128
  // blue = 0 * 0.5 + 255 * 0.5 = 128
  // rgb(255, 128, 128)
  assert.equal(texts[0].getAttribute("data-native-powerpoint-halo-color"), "rgb(255, 128, 128)");

  // black 0.25 composed with white:
  // r, g, b = 0 * 0.25 + 255 * 0.75 = 191
  // rgb(191, 191, 191)
  assert.equal(texts[1].getAttribute("data-native-powerpoint-halo-color"), "rgb(191, 191, 191)");
});

test("color parsing handles transparent, none, and absent colors", async () => {
  const { applyBackgroundAwareTextHalos, DEFAULT_TEXT_HALO_COLOR } = await loadTextHaloModule();

  const svg = parseSvg(`
    <svg>
      <!-- Inherit from previous test/default -->
      <rect fill="none" width="10" height="10" />
      <text>A</text>

      <rect fill="transparent" width="10" height="10" />
      <text>B</text>

      <rect width="10" height="10" />
      <text>C</text>
    </svg>
  `);

  applyBackgroundAwareTextHalos(svg, "table");

  const texts = svg.getElementsByTagName("text");
  assert.equal(texts[0].getAttribute("data-native-powerpoint-halo-color"), DEFAULT_TEXT_HALO_COLOR);
  assert.equal(texts[1].getAttribute("data-native-powerpoint-halo-color"), DEFAULT_TEXT_HALO_COLOR);
  assert.equal(texts[2].getAttribute("data-native-powerpoint-halo-color"), DEFAULT_TEXT_HALO_COLOR);
});

test("color parsing handles named colors and functions", async () => {
  const { applyBackgroundAwareTextHalos } = await loadTextHaloModule();

  const svg = parseSvg(`
    <svg>
      <rect fill="white" width="10" height="10" />
      <text>White</text>

      <rect fill="black" width="10" height="10" />
      <text>Black</text>

      <rect fill="rgb(0, 100, 200)" width="10" height="10" />
      <text>RGB</text>

      <rect fill="rgba(0, 100, 200, 0.5)" width="10" height="10" />
      <text>RGBA</text>
    </svg>
  `);

  applyBackgroundAwareTextHalos(svg, "table");

  const texts = svg.getElementsByTagName("text");

  // white
  assert.equal(texts[0].getAttribute("data-native-powerpoint-halo-color"), "rgb(255, 255, 255)");

  // black
  assert.equal(texts[1].getAttribute("data-native-powerpoint-halo-color"), "rgb(0, 0, 0)");

  // rgb(0, 100, 200)
  assert.equal(texts[2].getAttribute("data-native-powerpoint-halo-color"), "rgb(0, 100, 200)");

  // rgba(0, 100, 200, 0.5) against white
  // red: 0*0.5 + 255*0.5 = 128
  // green: 100*0.5 + 255*0.5 = 178
  // blue: 200*0.5 + 255*0.5 = 228
  assert.equal(texts[3].getAttribute("data-native-powerpoint-halo-color"), "rgb(128, 178, 228)");
});
