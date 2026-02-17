import { test } from "node:test";
import * as assert from "node:assert/strict";
import {
  appendPathSegmentsToUrl,
  buildRemotePathCandidates,
  getRelativePathByUrlPrefix,
  normalizeShareBaseUrl,
  toGraphShareId
} from "../resolver-utils";

test("buildRemotePathCandidates trims leading segments", () => {
  assert.deepEqual(buildRemotePathCandidates("/a/b/c.txt"), ["/a/b/c.txt", "/b/c.txt", "/c.txt"]);
});

test("normalizeShareBaseUrl fixes malformed protocol", () => {
  assert.equal(normalizeShareBaseUrl("https:/contoso.sharepoint.com/sites/Board/Shared Documents/"), "https://contoso.sharepoint.com/sites/Board/Shared Documents/");
});

test("appendPathSegmentsToUrl appends encoded path", () => {
  const value = appendPathSegmentsToUrl("https://contoso.sharepoint.com/sites/Board/Shared Documents/", ["General", "business-plan.md"]);
  assert.equal(value, "https://contoso.sharepoint.com/sites/Board/Shared%20Documents/General/business-plan.md");
});

test("getRelativePathByUrlPrefix returns remaining path", () => {
  const relative = getRelativePathByUrlPrefix(
    "https://contoso.sharepoint.com/sites/Board/Shared%20Documents/General/business-plan.md",
    "https://contoso.sharepoint.com/sites/Board/Shared Documents/"
  );
  assert.equal(relative, "General/business-plan.md");
});

test("toGraphShareId returns Graph share id format", () => {
  const shareId = toGraphShareId("https://contoso.sharepoint.com/sites/Board/Shared%20Documents/General/business-plan.md");
  assert.ok(shareId.startsWith("u!"));
  assert.ok(!shareId.includes("="));
});
