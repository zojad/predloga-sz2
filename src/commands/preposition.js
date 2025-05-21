/* global Office, Word */

let state = {
  // each entry now carries its originalColor
  errors: []          // Array<{ range: Word.Range, suggestion: string, originalColor: string }>
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Helpers for ribbon notifications
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
 * Decide “s” vs “z” from the first letter of rawWord.
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap = { '1':'e','2':'d','3':'t','4':'š','5':'p',
                     '6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

/**
 * 1) Check S/Z:
 *    - first restore any old pinks back to their original colour
 *    - clear our queue
 *    - find & highlight all mismatches, remembering original colour
 *    - select the first one
 */
export async function checkDocumentText() {
  // restore previous highlights
  if (state.errors.length > 0) {
    await Word.run(async context => {
      for (const { range, originalColor } of state.errors) {
        context.trackedObjects.add(range);
        range.font.highlightColor = originalColor;
      }
      await context.sync();
    });
  }
  // reset
  clearNotification(NOTIF_ID);
  state.errors = [];

  try {
    await Word.run(async context => {
      // search standalone "s" & "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // examine each candidate
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

        // compute expected
        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

        // preserve case
        const suggestion = raw === raw.toUpperCase()
          ? expectedLower.toUpperCase()
          : expectedLower;

        // **load & remember** the existing highlightColor
        r.font.load("highlightColor");
        await context.sync();
        const originalColor = r.font.highlightColor;

        // highlight pink & enqueue
        context.trackedObjects.add(r);
        r.font.highlightColor = HIGHLIGHT_COLOR;
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
        // select first mismatch
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
  }
}

/**
 * 2) Accept One:
 *    - take the first queued mismatch
 *    - replace it, clear its pink
 *    - re-run checkDocumentText() to restore other highlights & re-scan
 */
export async function acceptCurrentChange() {
  if (!state.errors.length) return;

  const { range, suggestion } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // re-scan: this will restore everyone else’s original colour
  await checkDocumentText();
}

/**
 * 3) Reject One: same as Accept One but just clear the pink
 */
export async function rejectCurrentChange() {
  if (!state.errors.length) return;

  const { range } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  await checkDocumentText();
}

/**
 * 4) Accept All: batch-replace every mismatch in one go,
 *              leave all other formatting alone
 */
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    for (const r of [...sRes.items, ...zRes.items]) {
      const raw = r.text.trim();
      if (!/^[sSzZ]$/.test(raw)) continue;

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

  // clear queue so next scan is fresh
  state.errors = [];

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

/**
 * 5) Reject All: batch-clear every pink mismatch
 */
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

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

  state.errors = [];

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Cleared all!",
    icon: "Icon.80x80"
  });
}
