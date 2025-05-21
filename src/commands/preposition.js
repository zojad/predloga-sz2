/* global Office, Word */

let state = {
  // each entry: { range: Word.Range, suggestion: "s"|"S"|"z"|"Z", originalColor: string|null }
  errors: [],
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID = "noErrors";

// — Helpers for ribbon notifications —
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
 * Given the next word, returns "s" or "z".
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: highlight mismatches, capture their original color, select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];

  try {
    await Word.run(async context => {
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // Load each range's current highlight color
      for (const r of [...sRes.items, ...zRes.items]) {
        r.font.load("highlightColor");
      }
      await context.sync();

      // Now decide which ones are wrong
      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;

        const originalColor = r.font.highlightColor; // may be null

        // peek at next word
        const after = r
          .getRange("After")
          .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

        // preserve case
        const suggestion = (raw === raw.toUpperCase())
          ? expectedLower.toUpperCase()
          : expectedLower;

        // highlight error only on that letter
        context.trackedObjects.add(r);
        r.font.highlightColor = HIGHLIGHT_COLOR;

        // queue it, with its original color
        state.errors.push({ range: r, suggestion, originalColor });
      }

      await context.sync();

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the very first mismatch
        const first = state.errors[0].range;
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
  } finally {
    state.isChecking = false;
  }
}

// ─────────────────────────────────────────────────
// 2) Accept One: replace the letter, re-apply its original color, select next
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;
  const { range, suggestion, originalColor } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    // replace text
    range.insertText(suggestion, Word.InsertLocation.replace);
    // restore whatever color was there before
    range.font.highlightColor = originalColor;
    await context.sync();
  });

  // if there’s another error queued, select it
  if (state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 3) Reject One: restore its original color, select next
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (!state.errors.length) return;
  const { range, originalColor } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = originalColor;
    await context.sync();
  });

  if (state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 4) Accept All: batch-replace every mismatch, restoring original colors
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    for (const { range, suggestion, originalColor } of state.errors) {
      context.trackedObjects.add(range);
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = originalColor;
    }
    await context.sync();
  });

  state.errors = [];
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: restore every flagged letter’s original color
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    for (const { range, originalColor } of state.errors) {
      context.trackedObjects.add(range);
      range.font.highlightColor = originalColor;
    }
    await context.sync();
  });

  state.errors = [];
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Restored original highlights!",
    icon: "Icon.80x80"
  });
}
