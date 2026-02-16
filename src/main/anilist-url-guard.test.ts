import test from "node:test";
import assert from "node:assert/strict";

import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from "./anilist-url-guard";

test("allows only AniList https URLs for external opens", () => {
  assert.equal(isAllowedAnilistExternalUrl("https://anilist.co"), true);
  assert.equal(
    isAllowedAnilistExternalUrl("https://www.anilist.co/settings/developer"),
    true,
  );
  assert.equal(isAllowedAnilistExternalUrl("http://anilist.co"), false);
  assert.equal(isAllowedAnilistExternalUrl("https://example.com"), false);
  assert.equal(isAllowedAnilistExternalUrl("file:///tmp/test"), false);
  assert.equal(isAllowedAnilistExternalUrl("not a url"), false);
});

test("allows only AniList https or data URLs for setup navigation", () => {
  assert.equal(
    isAllowedAnilistSetupNavigationUrl("https://anilist.co/api/v2/oauth/authorize"),
    true,
  );
  assert.equal(
    isAllowedAnilistSetupNavigationUrl(
      "data:text/html;charset=utf-8,%3Chtml%3E%3C%2Fhtml%3E",
    ),
    true,
  );
  assert.equal(
    isAllowedAnilistSetupNavigationUrl("https://example.com/redirect"),
    false,
  );
  assert.equal(isAllowedAnilistSetupNavigationUrl("javascript:alert(1)"), false);
});
