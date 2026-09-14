import test from "node:test";
import assert from "node:assert/strict";
import {
  STANDARD_CATEGORIES,
  breakdownQuantityDigits,
  categoryById,
  measurementSuggestion,
  partsForCategory,
  roundBreakdownQuantity,
  roundMeasurementQuantity,
  standardClassificationLabel,
} from "../src/quantity-classification.js";

test("standard renovation categories keep the official display order", () => {
  assert.deepEqual(STANDARD_CATEGORIES.map((category) => category.label), [
    "防水改修",
    "外壁改修",
    "建具改修",
    "内装改修",
    "塗装改修",
  ]);
  assert.deepEqual(STANDARD_CATEGORIES.map((category) => category.order), [2, 3, 4, 5, 6]);
});

test("painting renovation only exposes renovation while the other categories include removal", () => {
  assert.deepEqual(Object.keys(categoryById("painting_renovation").divisions), ["renovation"]);
  assert.ok(categoryById("exterior_wall_renovation").divisions.removal);
});

test("classification label omits unselected levels", () => {
  assert.equal(
    standardClassificationLabel({ categoryLabel: "外壁改修", divisionLabel: "改修", item: "仕上塗材塗り" }),
    "外壁改修 / 改修 / 仕上塗材塗り",
  );
});

test("measurement quantities are rounded to two decimal places", () => {
  assert.equal(roundMeasurementQuantity(12.345), 12.35);
  assert.equal(roundMeasurementQuantity(0), 0);
});

test("breakdown quantities use one decimal below 100 and integers from 100", () => {
  assert.equal(roundBreakdownQuantity(12.345), 12.3);
  assert.equal(roundBreakdownQuantity(99.96), 100);
  assert.equal(roundBreakdownQuantity(100.49), 100);
  assert.equal(roundBreakdownQuantity(100.5), 101);
  assert.equal(breakdownQuantityDigits(99.96), 1);
  assert.equal(breakdownQuantityDigits(100), 0);
});

test("measurement suggestions follow common standard item units", () => {
  assert.deepEqual(measurementSuggestion("シーリング撤去"), { mode: "線数量", unit: "m" });
  assert.deepEqual(measurementSuggestion("AW アルミニウム製窓"), { mode: "個数", unit: "か所" });
  assert.deepEqual(measurementSuggestion("仕上塗材塗り"), { mode: "面積", unit: "㎡" });
});

test("part candidates follow the selected renovation category", () => {
  assert.deepEqual(partsForCategory("interior_renovation").slice(0, 3), ["床", "幅木・壁", "天井"]);
  assert.ok(partsForCategory("waterproofing_renovation").includes("屋上"));
  assert.deepEqual(partsForCategory("unknown"), []);
});
