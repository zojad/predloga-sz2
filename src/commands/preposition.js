/* global Office, Word */

const state = {
  errors: [],        // Array of { range: Word.Range, suggestion: "s"|"z" }
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

// 1) Highlight everything & select first
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // clear old highlights
      state.errors.forEach(e => {
        context.trackedObjects.add(e.range);
        e.range.font.highlightColor = null;
      });
      await context.sync();

      // find s/z
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => ['s','z'].includes(r.text.trim().toLowerCase()));

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
          // highlight & queue
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
        // select the first mismatch
        state.currentIndex = 0;
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error(e);
    showNotification(NOTIF_ID, { type: "errorMessage", message: "Check failed" });
  } finally {
    state.isChecking = false;
  }
}

// 2) Accept one: operate on current selection
export async function acceptCurrentChange() {
  console.log("▶ acceptCurrentChange ▶ errors:", state.errors.length, "index:", state.currentIndex);
  if (state.currentIndex >= state.errors.length) return;

  const suggestion = state.errors[state.currentIndex].suggestion;

  await Word.run(async context => {
    // replace the selected preposition
    const sel = context.document.getSelection();
    sel.insertText(suggestion, Word.InsertLocation.replace);
    sel.font.highlightColor = null;

    // move to next index and select the next mismatch
    state.currentIndex++;
    if (state.currentIndex < state.errors.length) {
      const nextRange = state.errors[state.currentIndex].range;
      context.trackedObjects.add(nextRange);
      nextRange.select();
    }

    await context.sync();
  });
}

// 3) Reject one: clear highlight & move on
export async function rejectCurrentChange() {
  console.log("▶ rejectCurrentChange ▶ errors:", state.errors.length, "index:", state.currentIndex);
  if (state.currentIndex >= state.errors.length) return;

  await Word.run(async context => {
    // clear highlight on the currently selected mismatch
    const sel = context.document.getSelection();
    sel.font.highlightColor = null;

    // advance and select next
    state.currentIndex++;
    if (state.currentIndex < state.errors.length) {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
    }

    await context.sync();
  });
}

// 4) Accept all at once
export async function acceptAllChanges() {
  console.log("▶ acceptAllChanges ▶ errors:", state.errors.length);
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
  showNotification(NOTIF_ID, { type: "informationalMessage", message: "Accepted all!" });
}

// 5) Reject all at once
export async function rejectAllChanges() {
  console.log("▶ rejectAllChanges ▶ errors:", state.errors.length);
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const { range } of state.errors) {
      context.trackedObjects.add(range);
      range.font.highlightColor = null;
    }
    await context.sync();
  });

  state.errors = [];
  showNotification(NOTIF_ID, { type: "informationalMessage", message: "Cleared all!" });
}
