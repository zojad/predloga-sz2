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

  // normalize and trim
  const nw = nextWord.normalize("NFC").trim();
  if (!nw) return null;

  // grab the very first character
  let ch = nw[0];

  // if digit, map to letter; else lowercase letter
  const digitMap = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };
  const key = (ch >= '0' && ch <= '9')
    ? digitMap[ch]
    : ch.toLowerCase();

  // S/Z logic: unvoiced ⇒ "s", otherwise "z"
  if (prepLower === "s" || prepLower === "z") {
    const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
    return unvoiced.has(key) ? "s" : "z";
  }

  // K/H logic: before k or g ⇒ "h", otherwise "k"
  if (prepLower === "k" || prepLower === "h") {
    return (key === "k" || key === "g") ? "h" : "k";
  }

  return null;
}

// ─────────────────────────────────────────────────
// Utility: build list of body, headers, and footers
// ─────────────────────────────────────────────────
async function collectScanRanges(context) {
  const ranges = [];

  // include document body
  ranges.push(context.document.body);

  // include each section’s primary header & footer
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  for (const section of sections.items) {
    ranges.push(section.getHeader("primary"));
    ranges.push(section.getFooter("primary"));
  }

  return ranges;
}

// ─────────────────────────────────────────────────
// 1) Check S/Z/K/H: highlight all mismatches, select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // clear all existing highlights
      context.document.body.font.highlightColor = null;
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      for (const sec of sections.items) {
        sec.getHeader("primary").font.highlightColor = null;
        sec.getFooter("primary").font.highlightColor = null;
      }
      await context.sync();

      const opts = { matchWholeWord: true, matchCase: false };
      const scanRanges = await collectScanRanges(context);
      const mismatches = [];

      for (const rng of scanRanges) {
        const sRes = rng.search("s", opts);
        const zRes = rng.search("z", opts);
        const kRes = rng.search("k", opts);
        const hRes = rng.search("h", opts);
        sRes.load("items"); zRes.load("items");
        kRes.load("items"); hRes.load("items");
        await context.sync();

        for (const r of [...sRes.items, ...zRes.items, ...kRes.items, ...hRes.items]) {
          const raw   = r.text.trim();
          const lower = raw.toLowerCase();
          if (!["s","z","k","h"].includes(lower)) continue;

          const after = r
            .getRange("After")
            .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
          after.load("text");
          await context.sync();

          const nxt = after.text.trim();
          if (!nxt) continue;

          const expected = determineCorrectPreposition(nxt, lower);
          if (!expected || expected === lower) continue;

          context.trackedObjects.add(r);
          r.font.highlightColor = HIGHLIGHT_COLOR;
          mismatches.push(r);
        }
      }

      await context.sync();

      if (!mismatches.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
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
      const scanRanges = await collectScanRanges(context);

      for (const rng of scanRanges) {
        const sRes = rng.search("s", opts);
        const zRes = rng.search("z", opts);
        const kRes = rng.search("k", opts);
        const hRes = rng.search("h", opts);
        sRes.load("items"); zRes.load("items");
        kRes.load("items"); hRes.load("items");
        await context.sync();

        for (const r of [...sRes.items, ...zRes.items, ...kRes.items, ...hRes.items]) {
          const raw   = r.text.trim();
          const lower = raw.toLowerCase();
          if (!["s","z","k","h"].includes(lower)) continue;

          const after = r
            .getRange("After")
            .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
          after.load("text");
          await context.sync();

          const nxt = after.text.trim();
          if (!nxt) continue;

          const expected = determineCorrectPreposition(nxt, lower);
          if (!expected || expected === lower) continue;

          const replacement =
            raw === raw.toUpperCase()
              ? expected.toUpperCase()
              : expected;

          context.trackedObjects.add(r);
          r.insertText(replacement, Word.InsertLocation.replace);
          r.font.highlightColor = null;
        }
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
      const scanRanges = await collectScanRanges(context);

      for (const rng of scanRanges) {
        const sRes = rng.search("s", opts);
        const zRes = rng.search("z", opts);
        const kRes = rng.search("k", opts);
        const hRes = rng.search("h", opts);
        sRes.load("items"); zRes.load("items");
        kRes.load("items"); hRes.load("items");
        await context.sync();

        for (const r of [...sRes.items, ...zRes.items, ...kRes.items, ...hRes.items]) {
          const raw = r.text.trim();
          if (!/^[sSzZkKhH]$/.test(raw)) continue;
          context.trackedObjects.add(r);
          r.font.highlightColor = null;
        }
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
