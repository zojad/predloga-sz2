/* global Office, Word */

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// ─────────────────────────────────────────────────
// Helpers for ribbon notifications
// ─────────────────────────────────────────────────
function clearNotification(id) {
  if (Office.NotificationMessages?.deleteAsync) {
    Office.NotificationMessages.deleteAsync(id);
  }
}
function showNotification(id, opts) {
  if (Office.NotificationMessages?.addAsync) {
    Office.NotificationMessages.addAsync(id, opts);
  }
}

/**
 * Decide correct preposition for S/Z and K/H.
 * @param {string} nextWord    — the text of the following word
 * @param {string} prepLower   — the candidate preposition, already lowercased ("s","z","k" or "h")
 * @returns {"s"|"z"|"k"|"h"|null}
 */
function determineCorrectPreposition(nextWord, prepLower) {
  if (!nextWord) return null;

  // grab first letter or digit of nextWord
  const m = nextWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const first = m[0].toLowerCase();

  // S/Z logic: unvoiced ⇒ "s", otherwise "z"
  if (prepLower === "s" || prepLower === "z") {
    const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
    const digitMap = {
      '1':'e','2':'d','3':'t','4':'š','5':'p',
      '6':'š','7':'s','8':'o','9':'d','0':'n'
    };
    const key = /\d/.test(first) ? digitMap[first] : first;
    return unvoiced.has(key) ? "s" : "z";
  }

  // K/H logic: before k or g ⇒ "h", otherwise "k"
  if (prepLower === "k" || prepLower === "h") {
    return (first === "k" || first === "g") ? "h" : "k";
  }

  return null;
}

// ─────────────────────────────────────────────────
// 1) Check S/Z/K/H: highlight all mismatches, select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // clear *all* highlights (so we can re-scan cleanly)
      context.document.body.font.highlightColor = null;
      await context.sync();

      // search for standalone "s","z","k","h"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      const kRes = context.document.body.search("k", opts);
      const hRes = context.document.body.search("h", opts);
      sRes.load("items"); zRes.load("items");
      kRes.load("items"); hRes.load("items");
      await context.sync();

      const mismatches = [];

      // flatten all four result sets
      for (const r of [
        ...sRes.items,
        ...zRes.items,
        ...kRes.items,
        ...hRes.items
      ]) {
        const raw = r.text.trim();
        const lower = raw.toLowerCase();
        if (!["s","z","k","h"].includes(lower)) continue;

        // get the next word
        const after = r
          .getRange("After")
          .getNextTextRange(
            [" ", "\n", ".", ",", ";", "?", "!"],
            true
          );
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        // decide what it should be
        const expected = determineCorrectPreposition(nxt, lower);
        if (!expected || expected === lower) continue;

        // highlight and queue
        context.trackedObjects.add(r);
        r.font.highlightColor = HIGHLIGHT_COLOR;
        mismatches.push(r);
      }

      await context.sync();

      if (!mismatches.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the first one
        const first = mismatches[0];
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Check failed; please try again."
    });
  }
}

// ─────────────────────────────────────────────────
// 2) Accept All: replace every mismatch in one batch
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      const kRes = context.document.body.search("k", opts);
      const hRes = context.document.body.search("h", opts);
      sRes.load("items"); zRes.load("items");
      kRes.load("items"); hRes.load("items");
      await context.sync();

      for (const r of [
        ...sRes.items,
        ...zRes.items,
        ...kRes.items,
        ...hRes.items
      ]) {
        const raw = r.text.trim();
        const lower = raw.toLowerCase();
        if (!["s","z","k","h"].includes(lower)) continue;

        // peek at next word
        const after = r
          .getRange("After")
          .getNextTextRange(
            [" ", "\n", ".", ",", ";", "?", "!"],
            true
          );
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        // get expected
        const expected = determineCorrectPreposition(nxt, lower);
        if (!expected || expected === lower) continue;

        // preserve uppercase if needed
        const replacement =
          raw === raw.toUpperCase() ? expected.toUpperCase() : expected;

        context.trackedObjects.add(r);
        r.insertText(replacement, Word.InsertLocation.replace);
        r.font.highlightColor = null;
      }

      await context.sync();
    });

    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Accepted all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Accept all failed."
    });
  }
}

// ─────────────────────────────────────────────────
// 3) Reject All: clear every pink mismatch
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      const kRes = context.document.body.search("k", opts);
      const hRes = context.document.body.search("h", opts);
      sRes.load("items"); zRes.load("items");
      kRes.load("items"); hRes.load("items");
      await context.sync();

      for (const r of [
        ...sRes.items,
        ...zRes.items,
        ...kRes.items,
        ...hRes.items
      ]) {
        const raw = r.text.trim();
        if (!/^[sSzZkKhH]$/.test(raw)) continue;
        context.trackedObjects.add(r);
        r.font.highlightColor = null;
      }

      await context.sync();
    });

    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Cleared all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Reject all failed."
    });
  }
}
