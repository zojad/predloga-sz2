/* global Office, Word */

let state = {
  // queue of mismatches: { range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }[]
  errors: [],
  isChecking: false
};

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
 * Decide “s” vs “z” from the first letter of the next word.
 * Unvoiced initials ⇒ “s”, otherwise ⇒ “z”.
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p',
                       '6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: clear ALL highlights & tracked ranges,
//    re-scan, highlight mismatches, select first.
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];

  try {
    await Word.run(async context => {
      // *** Release any stale tracked ranges up front ***
      context.trackedObjects.releaseAll();
      await context.sync();

      // 1) clear old highlights
      const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // 2) find every standalone “s” or “z”
      const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // 3) evaluate each candidate
      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;

        const after = r
          .getRange("After")
          .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || raw.toLowerCase() === expectedLower) continue;

        const suggestion = raw === raw.toUpperCase()
          ? expectedLower.toUpperCase()
          : expectedLower;

        context.trackedObjects.add(r);
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
        // select the first mismatch
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
// 2) Accept One: replace the first in queue, clear its
//    highlight, then call checkDocumentText() again.
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;
  const { range, suggestion } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  await checkDocumentText();
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear the first in queue, then re-scan.
// ─────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────
// 4) Accept All: apply every queued suggestion in one shot.
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const { range, suggestion } of state.errors) {
      context.trackedObjects.add(range);
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = null;
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
// 5) Reject All: clear highlights for every queued range.
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const { range } of state.errors) {
      context.trackedObjects.add(range);
      range.font.highlightColor = null;
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
