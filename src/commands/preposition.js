/* global Office, Word */

let state = {
  errors: [],
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

function showNotification(id, options) {
  if (Office.NotificationMessages?.addAsync) {
    Office.NotificationMessages.addAsync(id, options);
  }
}

function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const first = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  if (/\d/.test(first)) return unvoiced.has(numMap[first]) ? "s" : "z";
  return unvoiced.has(first) ? "s" : "z";
}

// --- CORE: find and highlight errors ---
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // reset previous highlights & state
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items");
      zRes.load("items");
      await context.sync();

      const all = [...sRes.items, ...zRes.items]
        .filter(r => {
          const t = r.text.trim().toLowerCase();
          return t === "s" || t === "z";
        });

      for (const prep of all) {
        const after = prep
          .getRange("After")
          .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nxt = after.text.trim();
        if (!nxt) continue;
        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          context.trackedObjects.add(prep);
          state.errors.push({ range: prep, suggestion: expect });
        }
      }

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // highlight all and select the first one
        state.errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
        await context.sync();
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

// --- Accept one: replace current and auto-advance ---
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);

      // replace text and clear highlight
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;

      // advance to next and select it
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const next = state.errors[state.currentIndex].range;
        context.trackedObjects.add(next);
        next.select();
      }

      await context.sync();
    });
  } catch (e) {
    console.error("acceptCurrentChange error", e);
  }
}

// --- Reject one: clear highlight and auto-advance ---
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);

      // clear highlight
      err.range.font.highlightColor = null;

      // advance to next and select it
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const next = state.errors[state.currentIndex].range;
        context.trackedObjects.add(next);
        next.select();
      }

      await context.sync();
    });
  } catch (e) {
    console.error("rejectCurrentChange error", e);
  }
}

// --- Accept all: replace every mismatch at once ---
export async function acceptAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
    });

    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Accepted all mismatches!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
  }
}

// --- Reject all: clear all highlights at once ---
export async function rejectAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.font.highlightColor = null;
      }
      await context.sync();
    });

    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Cleared all highlights!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
  }
}
