/* global Office, Word */

const state = {
  errors: [],        // { range: Word.Range, suggestion: "s" | "z" }[]
  currentIndex: 0,
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

// Your existing next-letter logic, purely in JS:
function determineCorrectPreposition(word) {
  if (!word) return null;
  const m = word.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const first = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(first) ? digitMap[first] : first;
  return unvoiced.has(key) ? "s" : "z";
}

// ───────────────────────────────────────────────────
// 1) checkDocumentText: find & highlight all mismatches
// ───────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  // reset
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // 1. clear old highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);

      // 2. search for standalone "s" or "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // filter to exactly "s"/"z", then find the next word after each
      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => r.text.trim().toLowerCase() === "s" || r.text.trim().toLowerCase() === "z");

      for (const r of candidates) {
        // grab next word range
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nextWord = after.text.trim();
        if (!nextWord) continue;

        const actual   = r.text.trim().toLowerCase();
        const expected = determineCorrectPreposition(nextWord);
        if (expected && actual !== expected) {
          // track & highlight this mismatch
          context.trackedObjects.add(r);
          r.font.highlightColor = HIGHLIGHT_COLOR;
          state.errors.push({ range: r, suggestion: expected });
        }
      }

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the first
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error(e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Check failed; please try again."
    });
  } finally {
    state.isChecking = false;
  }
}


// ───────────────────────────────────────────────────
// 2) Accept current: replace, clear highlight, auto-advance
// ───────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  const { range, suggestion } = state.errors[state.currentIndex];

  try {
    await Word.run(async context => {
      context.trackedObjects.add(range);

      // replace the single letter
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = null;

      // move to next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const next = state.errors[state.currentIndex].range;
        context.trackedObjects.add(next);
        next.select();
      }

      await context.sync();
    });

    // remove the handled error from our list
    state.errors.splice(state.currentIndex - 1, 1);
    if (state.currentIndex > state.errors.length) {
      state.currentIndex = 0;
    }
  } catch (e) {
    console.error(e);
  }
}


// ───────────────────────────────────────────────────
// 3) Reject current: clear highlight, auto-advance
// ───────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  const { range } = state.errors[state.currentIndex];

  try {
    await Word.run(async context => {
      context.trackedObjects.add(range);

      // just clear the highlight
      range.font.highlightColor = null;

      // move to next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const next = state.errors[state.currentIndex].range;
        context.trackedObjects.add(next);
        next.select();
      }

      await context.sync();
    });

    // remove the skipped error
    state.errors.splice(state.currentIndex - 1, 1);
    if (state.currentIndex > state.errors.length) {
      state.currentIndex = 0;
    }
  } catch (e) {
    console.error(e);
  }
}


// ───────────────────────────────────────────────────
// 4) Accept all: replace every mismatch in one batch
// ───────────────────────────────────────────────────
export async function acceptAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      // loop and replace/clear
      for (const { range, suggestion } of state.errors) {
        context.trackedObjects.add(range);
        range.insertText(suggestion, Word.InsertLocation.replace);
        range.font.highlightColor = null;
      }
      await context.sync();
    });

    // reset local state
    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Accepted all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error(e);
  }
}


// ───────────────────────────────────────────────────
// 5) Reject all: clear all highlights in one batch
// ───────────────────────────────────────────────────
export async function rejectAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (const { range } of state.errors) {
        context.trackedObjects.add(range);
        range.font.highlightColor = null;
      }
      await context.sync();
    });

    // reset local state
    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Cleared all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error(e);
  }
}

