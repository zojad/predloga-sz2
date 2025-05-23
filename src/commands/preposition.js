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

// ─────────────────────────────────────────────────
// Decide “s” vs “z” based on the next word’s first letter
// ─────────────────────────────────────────────────
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: highlight all mismatches, select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // clear *only* your pink highlights (in case you run twice)
      // NOTE: if you previously used only batch accept/reject, you can skip this.
      context.document.body.font.highlightColor = null;
      await context.sync();

      // search for standalone "s" & "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const mismatches = [];

      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;

        // look at next word
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

        // highlight it pink
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
        // select the first mismatch
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
      sRes.load("items"); zRes.load("items");
      await context.sync();

      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;

        // peek at next word
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

        const suggestion = raw === raw.toUpperCase()
          ? expectedLower.toUpperCase()
          : expectedLower;

        context.trackedObjects.add(r);
        r.insertText(suggestion, Word.InsertLocation.replace);
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
      sRes.load("items"); zRes.load("items");
      await context.sync();

      for (const r of [...sRes.items, ...zRes.items]) {
        if (/^[sSzZ]$/.test(r.text.trim())) {
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
