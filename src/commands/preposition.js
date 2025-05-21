/* global Office, Word */

// In‐memory state for mismatches
let state = {
  errors: [],        // Array of { range: Word.Range, suggestion: "s"|"z" }
  currentIndex: 0,
  isChecking: false
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

// Decide “s” vs “z” from the first letter of the next word
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return voiceless.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: highlight all mismatches & select the first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  // reset state
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // 1. Clear previous highlights (if any)
      for (const e of state.errors) {
        context.trackedObjects.add(e.range);
        e.range.font.highlightColor = null;
      }
      await context.sync();

      // 2. Find standalone "s" and "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => ['s','z'].includes(r.text.trim().toLowerCase()));

      // 3. For each candidate, get the next word and compare
      for (const r of candidates) {
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nxt = after.text.trim();
        if (!nxt) continue;

        const actual   = r.text.trim().toLowerCase();
        const expected = determineCorrectPreposition(nxt);
        if (expected && actual !== expected) {
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
        // Select the very first mismatch
        state.currentIndex = 0;
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
// 2) Accept One: replace current & auto-advance
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  // Remove current from queue
  const { range, suggestion } = state.errors.splice(state.currentIndex, 1)[0];

  // Step 1: replace & clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: select next mismatch if any
  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear current & auto-advance
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  // Remove current from queue
  const { range } = state.errors.splice(state.currentIndex, 1)[0];

  // Step 1: clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: select next mismatch if any
  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 4) Accept All: repopulate if needed, then replace all
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  // If queue is empty, repopulate via a fresh check
  if (!state.errors.length) {
    await checkDocumentText();
  }
  console.log("▶ acceptAllChanges; errors:", state.errors.length);
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
  state.currentIndex = 0;
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: repopulate if needed, then clear all
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  // If queue is empty, repopulate via a fresh check
  if (!state.errors.length) {
    await checkDocumentText();
  }
  console.log("▶ rejectAllChanges; errors:", state.errors.length);
  if (!state.errors.length) return;

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

