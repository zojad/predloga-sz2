/* global Office, Word */

let state = {
  errors: [],        // array of { range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Helpers to push/pull ribbon notifications
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
 * Look at the first letter of rawWord and decide "s" vs "z".
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p',
                       '6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return voiceless.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Fresh scan & highlight all mismatches, select the first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];

  try {
    await Word.run(async context => {
      // ** flush _all_ highlights in the body **
      context.document.body.font.highlightColor = null;
      await context.sync();

      // search for standalone "s" & "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // for each letter candidate decide if it's wrong
      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;   // exactly one letter

        // peek at the next word
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        // decide what it should have been
        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

        // preserve upper/lower
        const suggestion = raw === raw.toUpperCase()
          ? expectedLower.toUpperCase()
          : expectedLower;

        // highlight just that single-letter range
        r.font.highlightColor = HIGHLIGHT_COLOR;
        state.errors.push({ range: r, suggestion });
      }

      await context.sync();

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the very first wrong letter
        const first = state.errors[0].range;
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
// 2) Accept One: fix the first mismatch & then re-scan
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;

  // pull off first item
  const { range, suggestion } = state.errors.shift();

  // replace it
  await Word.run(async context => {
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // re-run the full scan → next mismatch will be selected automatically
  await checkDocumentText();
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear highlight on first mismatch & then re-scan
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (!state.errors.length) return;

  // pull off first item
  const { range } = state.errors.shift();

  // just clear its highlight
  await Word.run(async context => {
    range.font.highlightColor = null;
    await context.sync();
  });

  // re-run a fresh scan
  await checkDocumentText();
}

// ─────────────────────────────────────────────────
// 4) Accept All: batch-replace every mismatch in one go
// ─────────────────────────────────────────────────
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

      r.insertText(suggestion, Word.InsertLocation.replace);
      r.font.highlightColor = null;
    }

    await context.sync();
  });

  // clear our queue so we can scan again
  state.errors = [];

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: batch-clear every highlight in one go
// ─────────────────────────────────────────────────
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
