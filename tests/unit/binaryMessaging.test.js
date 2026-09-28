const test = require("node:test");
const assert = require("node:assert/strict");

const {
  encodeTextToBinary,
  decodeBinaryToText,
  isLikelyBinaryMessage,
} = require("../../app/browser/tools/binaryMessaging");

const HELLO_BINARY =
  "01001000 01100101 01101100 01101100 01101111";
const HELLO_WORLD_BINARY =
  "01001000 01100101 01101100 01101100 01101111 00100000 01010111 01101111 01110010 01101100 01100100";

test("encodeTextToBinary emits readable 8-bit UTF-8 groups", () => {
  assert.equal(encodeTextToBinary("Hello"), HELLO_BINARY);
  assert.equal(encodeTextToBinary("Hello World"), HELLO_WORLD_BINARY);
  assert.equal(encodeTextToBinary(""), "");
  assert.equal(encodeTextToBinary(" "), "00100000");
  assert.equal(encodeTextToBinary("é"), "11000011 10101001");
  assert.equal(
    encodeTextToBinary("😀"),
    "11110000 10011111 10011000 10000000"
  );
});

test("decodeBinaryToText accepts binary whitespace and strict UTF-8", () => {
  assert.equal(decodeBinaryToText(HELLO_BINARY), "Hello");
  assert.equal(decodeBinaryToText(HELLO_WORLD_BINARY), "Hello World");
  assert.equal(decodeBinaryToText(""), "");
  assert.equal(decodeBinaryToText("00100000"), " ");
  assert.equal(
    decodeBinaryToText(
      "01001000 01100101\n01101100\t01101100 01101111"
    ),
    "Hello"
  );
  assert.throws(
    () => decodeBinaryToText("11000011 00101000"),
    /encoded data was not valid/u
  );
});

test("UTF-8 text round-trips through binary", () => {
  for (const original of [
    "Hello",
    "Hello World",
    "",
    " ",
    "こんにちは",
    "नमस्ते",
    "é",
    "😀",
  ]) {
    assert.equal(decodeBinaryToText(encodeTextToBinary(original)), original);
  }
});

test("decodeBinaryToText rejects malformed groups", () => {
  for (const invalid of ["hello", "101", "12345678", "01010201", "10 10"]) {
    assert.throws(() => decodeBinaryToText(invalid), /8-bit groups/u);
  }
});

test("isLikelyBinaryMessage is conservative", () => {
  for (const ordinary of [
    "hello",
    "101",
    "12345678",
    "01010201",
    "10 10",
    "Meeting at 10:10",
    "Version 101",
    "",
    " ",
    "01000001",
    "11000011 00101000",
  ]) {
    assert.equal(isLikelyBinaryMessage(ordinary), false, ordinary);
  }

  assert.equal(isLikelyBinaryMessage(HELLO_BINARY), true);
  assert.equal(
    isLikelyBinaryMessage(
      "01001000 01100101\n01101100\t01101100 01101111"
    ),
    true
  );
  assert.equal(
    isLikelyBinaryMessage(
      "01001000\u00A001100101\u00A001101100\u00A001101100\u00A001101111"
    ),
    true
  );
  for (const original of ["こんにちは", "नमस्ते", "é", "😀"]) {
    assert.equal(isLikelyBinaryMessage(encodeTextToBinary(original)), true);
  }
});

const { normalizeWhitespace } = require("../../app/browser/tools/binaryMessaging");

test("normalizeWhitespace handles normal and non-breaking spaces and line breaks", () => {
  assert.equal(normalizeWhitespace(""), "");
  assert.equal(normalizeWhitespace("  hello   world  "), "hello world");
  assert.equal(
    normalizeWhitespace("hello\u00A0world\nnext\tline"),
    "hello world next line"
  );
  assert.equal(
    normalizeWhitespace("01001000\u00A001100101\n01101100"),
    "01001000 01100101 01101100"
  );
});

test("long messages round-trip through binary encoding accurately", () => {
  const longText = "This is a detailed and very long message testing binary messaging mode. "
    .repeat(30)
    .trim();
  const encoded = encodeTextToBinary(longText);
  assert.equal(isLikelyBinaryMessage(encoded), true);
  assert.equal(decodeBinaryToText(encoded), longText);
});

