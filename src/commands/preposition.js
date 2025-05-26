/* global Office, Word */

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

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
 */
function determineCorrectPreposition(nextWord, prepLower) {
  if (!nextWord) return null;
  const nw = nextWord.normalize("NFC").trim();
  if (!nw) return null;
  let ch = nw[0];
  const digitMap = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };
  const key = (ch >= '0' && ch <= '9') ? digitMap[ch] : ch.toLowerCase();

  if (prepLower === "s" || prepLower === "z") {
    const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
    return unvoiced.has(key) ? "s" : "z";
  }
  if (prepLower === "k" || prepLower === "h") {
    return (key === "k" || key === "g") ? "h" : "k";
  }
  return null;
}

/**
 * Collect every Range we want to scan:
 *  - main body
 *  - each section’s primary header.body & footer.body
 * We also load("text") on each so search() actually sees content.
 */
async function collectScanRanges(context) {
  const ranges = [];

  // 1) main document body
  const body = context.document.body;
  ranges.push(body);
  body.load("text");

  // 2) headers & footers
  if (context.document.sections) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (const section of sections.items) {
      // primary header
      try {
        const hdrBody = section.getHeader("primary").body;
        ranges.push(hdrBody);
        hdrBody.load("text");
      } catch { /* no header here */ }

      // primary footer
      try {
        const ftrBody = section.getFooter("primary").body;
        ranges.push(ftrBody);
        ftrBody.load("text");
      } catch { /* no footer here */ }
    }
  }

  // sync so every Range.text is populated
  await context.sync();
  return ranges;
}

export async function checkDocumentText() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      const opts = { matchWholeWord: true, matchCase: false };
      const scanRanges = await collectScanRanges(context);

      // clear any prior highlights everywhere
      for (const rng of scanRanges) {
        rng.font.highlightColor = null;
      }
      await context.sync();

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
          type: "informationalMessage", message: "✨ No mismatches!", icon: "Icon.80x80"
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
      type: "errorMessage", message: "Check failed; please try again."
    });
  }
}

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

          const replacement = raw === raw.toUpperCase()
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
      type: "informationalMessage", message: "Accepted all!", icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage", message: "Accept all failed."
    });
  }
}

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
      type: "informationalMessage", message: "Cleared all!", icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage", message: "Reject all failed."
    });
  }
}
