import test from "node:test";
import assert from "node:assert/strict";
import { categorizeMessage, escapeHtml } from "../src/organize.js";

test("categorizes common action subjects without modifying messages", () => {
  assert.equal(categorizeMessage({ subject: "Approval needed today" }), "Action needed");
  assert.equal(categorizeMessage({ subject: "Team meeting agenda" }), "Meetings");
  assert.equal(categorizeMessage({ subject: "Payment receipt" }), "Finance");
  assert.equal(categorizeMessage({ subject: "Hello" }), "Other");
});

test("handles missing or non-string subjects", () => {
  assert.equal(categorizeMessage({}), "Other");
  assert.equal(categorizeMessage({ subject: null }), "Other");
  assert.equal(categorizeMessage({ subject: 42 }), "Other");
});

test("escapes untrusted message fields before rendering HTML", () => {
  assert.equal(escapeHtml(`<script a="b">&'`), "&lt;script a=&quot;b&quot;&gt;&amp;&#39;");
  assert.equal(escapeHtml(null), "");
});
