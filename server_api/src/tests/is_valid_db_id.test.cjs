"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");

const { isValidDbId } = require("../utils/is_valid_db_id.cjs");

test("isValidDbId guards database identifiers from request input", async (t) => {
  await t.test("accepts positive integer ids as numbers and strings", () => {
    assert.equal(isValidDbId(1), true);
    assert.equal(isValidDbId(123), true);
    assert.equal(isValidDbId("123"), true);
  });

  await t.test("rejects null, undefined, and empty values", () => {
    assert.equal(isValidDbId(null), false);
    assert.equal(isValidDbId(undefined), false);
    assert.equal(isValidDbId(""), false);
  });

  await t.test("rejects zero as a number and a string", () => {
    assert.equal(isValidDbId(0), false);
    assert.equal(isValidDbId("0"), false);
  });

  await t.test("rejects id strings that contain a decimal point", () => {
    assert.equal(isValidDbId("12.5"), false);
    assert.equal(isValidDbId("3.0"), false);
  });

  await t.test("rejects non-numeric strings", () => {
    assert.equal(isValidDbId("abc"), false);
    assert.equal(isValidDbId("not-an-id"), false);
  });
});
