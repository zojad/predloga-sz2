/* global Office, Word */

let state = {
  errors: [],        // Array<{ range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }>
  currentIndex: 0
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
 * Decide proper “s” vs “z” based on the first letter/digit of the next word
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
// 1) Check S/Z: single full scan, track & highlight all mismatches, select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  await Word.run(async context => {
    // A) Clear old highlights
    const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
    const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
    oldS.load("items"); oldZ.load("items");
    await context.sync();
    [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
    await context.sync();

    // B) Find all standalone “s” and “z”
    const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
    const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
    sRes.load("items"); zRes.load("items");
    await context.sync();

    // C) Evaluate & enqueue mismatches
    for (const r of [...sRes.items, ...zRes.items]) {
      const raw = r.text.trim();
      if (!/^[sSzZ]$/.test(raw)) continue;

      // look at the next word
      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();
      const nxt = after.text.trim();
      if (!nxt) continue;

      const expectedLower = determineCorrectPreposition(nxt);
      if (!expectedLower || expectedLower === raw.toLowerCase()) continue;

      // preserve original case
      const suggestion = raw === raw.toUpperCase()
        ? expectedLower.toUpperCase()
        : expectedLower;

      // track this range *now* and highlight it
      context.trackedObjects.add(r);
      r.font.highlightColor = HIGHLIGHT_COLOR;
      state.errors.push({ range: r, suggestion });
    }

    await context.sync();

    // D) select first mismatch (or notify “no mismatches”)
    if (!state.errors.length) {
      showNotification(NOTIF_ID, {
        type: "informationalMessage",
        message: "✨ No mismatches!",
        icon: "Icon.80x80"
      });
    } else {
      const firstRange = state.errors[0].range;
      context.trackedObjects.add(firstRange);
      firstRange.select();
      await context.sync();
    }
  });
}

// ─────────────────────────────────────────────────
// 2) Accept One: replace current, clear highlight, advance & select next
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  // get & remove the current mismatch
  const { range, suggestion } = state.errors[state.currentIndex];

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();

    // advance pointer
    state.currentIndex++;
    if (state.currentIndex < state.errors.length) {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    }
  });
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear highlight on current, advance & select next
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  const { range } = state.errors[state.currentIndex];

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();

    state.currentIndex++;
    if (state.currentIndex < state.errors.length) {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    }
  });
}

// ─────────────────────────────────────────────────
// 4) Accept All: replace & clear all mismatches in one pass
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    for (const { range, suggestion } of state.errors) {
      context.trackedObjects.add(range);
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = null;
    }
    await context.sync();
  });

  // reset state
  state.errors = [];
  state.currentIndex = 0;
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: clear all highlights at once
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    for (const { range } of state.errors) {
      context.trackedObjects.add(range);
      range.font.highlightColor = null;
    }
    await context.sync();
  });

  state.errors = [];
  state.currentIndex = 0;
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Cleared all!",
    icon: "Icon.80x80"
  });
}
