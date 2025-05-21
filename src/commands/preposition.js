/* global Office, Word */

let state = {
  errors: [],        // { range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }[]
  currentIndex: 0
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

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
 * Decide “s” vs “z” from the first letter of the next word.
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
// 1) Check S/Z: highlight *all* mismatches & select the first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  // reset our queue & UI
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // 1) clear any old highlights
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

        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const expected = determineCorrectPreposition(nxt);
        if (!expected || expected === raw.toLowerCase()) continue;

        // preserve case
        const suggestion = raw === raw.toUpperCase()
          ? expected.toUpperCase()
          : expected;

        context.trackedObjects.add(r);
        r.font.highlightColor = HIGHLIGHT_COLOR;
        state.errors.push({ range: r, suggestion });
      }

      await context.sync();

      // 4) no mismatches → notification, else select the first
      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select(Word.SelectionMode.select);
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
// 2) Accept One: replace & clear current, then select next
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  await Word.run(async context => {
    const { range, suggestion } = state.errors[state.currentIndex];
    context.trackedObjects.add(range);

    // replace text + clear its highlight
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;

    // advance index
    state.currentIndex++;

    // if more remain, select the next one
    if (state.currentIndex < state.errors.length) {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select(Word.SelectionMode.select);
    }

    await context.sync();
  });
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear highlight on current, then select next
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  await Word.run(async context => {
    const { range } = state.errors[state.currentIndex];
    context.trackedObjects.add(range);

    // clear highlight only
    range.font.highlightColor = null;

    // advance index
    state.currentIndex++;

    // select next if present
    if (state.currentIndex < state.errors.length) {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select(Word.SelectionMode.select);
    }

    await context.sync();
  });
}

// ─────────────────────────────────────────────────
// 4) Accept All: batch‐replace & clear every mismatch
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

  state.errors = [];
  state.currentIndex = 0;

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: batch‐clear every highlight
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
