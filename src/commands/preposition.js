/* global Office, Word */

let state = {
  errors: [],        // Array<{ range: Word.Range, suggestion: "s"|"z" }>
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Ribbon notification helpers
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
// 1) Check S/Z: highlight all mismatches & select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  // reset our queue
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // Clear any old highlights
      const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // Find standalone "s" and "z"
      const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // Evaluate each candidate
      for (const r of [...sRes.items, ...zRes.items]) {
        const t = r.text.trim().toLowerCase();
        if (t !== "s" && t !== "z") continue;

        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nxt = after.text.trim();
        if (!nxt) continue;

        const expected = determineCorrectPreposition(nxt);
        if (expected && expected !== t) {
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
// 2) Accept One: replace current & auto-advance
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  // Pull out the current mismatch
  const { range, suggestion } = state.errors.splice(state.currentIndex, 1)[0];

  // Step 1: replace & clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: select next mismatch (if any)
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

  // Remove current mismatch from queue
  const { range } = state.errors.splice(state.currentIndex, 1)[0];

  // Step 1: clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  // Step 2: select next mismatch (if any)
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
// 4) Accept All: fresh search + replace every mismatch
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
      const t = r.text.trim().toLowerCase();
      if (t !== "s" && t !== "z") continue;

      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();

      const nxt = after.text.trim();
      const expected = determineCorrectPreposition(nxt);
      if (expected && expected !== t) {
        r.insertText(expected, Word.InsertLocation.replace);
        r.font.highlightColor = null;
      }
    }

    await context.sync();
  });

  // Reset queue
  state.errors = [];
  state.currentIndex = 0;

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: fresh search + clear every highlight
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
      const t = r.text.trim().toLowerCase();
      if (t === "s" || t === "z") {
        r.font.highlightColor = null;
      }
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
