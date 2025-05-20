/* global Office, Word */

// In‐memory state for mismatches
const state = {
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

// Decide whether the next word wants "s" or "z"
function determineCorrectPreposition(word) {
  if (!word) return null;
  const m = word.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return voiceless.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Highlight all mismatches & select the first one
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // Clear any old highlights
      state.errors.forEach(e => {
        context.trackedObjects.add(e.range);
        e.range.font.highlightColor = null;
      });
      await context.sync();

      // Find all standalone "s" or "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => ['s','z'].includes(r.text.trim().toLowerCase()));

      // Evaluate each candidate
      for (const r of candidates) {
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nextWord = after.text.trim();
        if (!nextWord) continue;

        const actual   = r.text.trim().toLowerCase();
        const expected = determineCorrectPreposition(nextWord);
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
        // Select the first mismatch
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
// 2) Accept one: replace current & auto‐advance
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  console.log("▶ acceptCurrentChange() fired;", 
    "errors:", state.errors.map(e => e.range.text),
    "currentIndex:", state.currentIndex
  );
  if (!state.errors.length) return;

  // Pull off the first mismatch
  const { range, suggestion } = state.errors.shift();

  // Step 1: replace the text and clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: if another remains, select it
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
// 3) Reject one: clear current & auto‐advance
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  console.log("▶ rejectCurrentChange() fired;", 
    "errors:", state.errors.map(e => e.range.text),
    "currentIndex:", state.currentIndex
  );
  if (!state.errors.length) return;

  // Drop the first mismatch
  const { range } = state.errors.shift();

  // Step 1: clear its highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: if another remains, select it
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
// 4) Accept all: replace every mismatch at once
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  console.log("▶ acceptAllChanges() fired;", 
    "errors:", state.errors.map(e => e.range.text)
  );
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
// 5) Reject all: clear all highlights at once
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  console.log("▶ rejectAllChanges() fired;", 
    "errors:", state.errors.map(e => e.range.text)
  );
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
