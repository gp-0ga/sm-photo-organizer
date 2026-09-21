import test from "node:test";
import assert from "node:assert/strict";
import { detectOpeningRectangles, intersectionOverUnion } from "../src/opening-detector.js";

function fixture(width = 240, height = 170) {
  const data = new Uint8ClampedArray(width * height * 4).fill(255);
  function pixel(x, y) {
    const index = (y * width + x) * 4;
    data[index] = data[index + 1] = data[index + 2] = 0;
  }
  function rectangle(x, y, w, h) {
    for (let px = x; px <= x + w; px += 1) { pixel(px, y); pixel(px, y + h); }
    for (let py = y; py <= y + h; py += 1) { pixel(x, py); pixel(x + w, py); }
  }
  return { image: { width, height, data }, rectangle };
}

const wall = { id: "wall-a", points: [{ x: 20, y: 20 }, { x: 220, y: 20 }, { x: 220, y: 140 }, { x: 20, y: 140 }] };

test("finds two closed opening rectangles inside a wall", () => {
  const { image, rectangle } = fixture();
  rectangle(40, 45, 38, 50);
  rectangle(110, 48, 54, 46);
  const found = detectOpeningRectangles(image, [wall]);
  assert.equal(found.length, 2);
  assert.deepEqual(found.map(({ x, y, width, height }) => ({ x, y, width, height })), [
    { x: 40, y: 45, width: 38, height: 50 },
    { x: 110, y: 48, width: 54, height: 46 },
  ]);
  assert.ok(found.every((item) => item.confidence > 0.6 && item.wallId === "wall-a"));
});

test("does not take rectangles outside the confirmed wall", () => {
  const { image, rectangle } = fixture();
  rectangle(40, 45, 38, 50);
  rectangle(175, 145, 35, 20);
  const found = detectOpeningRectangles(image, [wall]);
  assert.equal(found.length, 1);
});

test("rejects incomplete rectangles and handles duplicates", () => {
  const { image, rectangle } = fixture();
  rectangle(40, 45, 38, 50);
  for (let y = 45; y <= 95; y += 1) {
    const index = (y * image.width + 78) * 4;
    image.data[index] = image.data[index + 1] = image.data[index + 2] = 255;
  }
  assert.equal(detectOpeningRectangles(image, [wall]).length, 0);
  assert.equal(intersectionOverUnion({ x: 0, y: 0, width: 20, height: 20 }, { x: 2, y: 2, width: 20, height: 20 }) > 0.55, true);
});
