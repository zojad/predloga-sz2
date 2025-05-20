/* global Office, Word */

let state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID = "noErrors";

function clearNotification(id) {
  if (Office.NotificationMessages && typeof Office.NotificationMessages.deleteAsync === "function") {
    Office.NotificationMessages.deleteAsync(id);
  }
}

function showNotification(id, options) {
  if (Office.NotificationMessages && typeof Office.NotificationMessages.addAsync === "function") {
    Office.NotificationMessages.addAsync(id, options);
  }
}

function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const match = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!match) return null;
  const first = match[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap = {'1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n'};
  return (/[0-9]/.test(first)
    ? (unvoiced.has(numMap[first]) ? 's' : 'z')
    : (unvoiced.has(first) ? 's' : 'z')
  );
}

export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // Clear previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];

      // Search for standalone 's' and 'z'
      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const all = [...sRes.items, ...zRes.items];
      for (const prep of all) {
        const after = prep.getRange("After");
        const nextRange = after.getNextTextRange([" ","\n",".",",",";","?","!"], true);
        nextRange.load("text");
        await context.sync();

        const nxt = nextRange.text.trim();
        if (!nxt) continue;
        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          context.trackedObjects.add(prep);
          state.errors.push({ range: prep, suggestion: expect });
        }
      }

      // Highlight and select first
      if (state.errors.length) {
        state.errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
        await context.sync();
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      } else {
        showNotification(NOTIF_ID, { type: 'informationalMessage', message: "✨ No mismatches!", icon: 'Icon.80x80' });
      }
    });
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification("checkError", { type: 'errorMessage', message: 'Check failed', persistent: false });
  } finally {
    state.isChecking = false;
  }
}

export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);
      // Replace and clear highlight
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();

      // Advance index
      state.currentIndex++;

      // Select next, if any
      if (state.currentIndex < state.errors.length) {
        const nextErr = state.errors[state.currentIndex];
        context.trackedObjects.add(nextErr.range);
        nextErr.range.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("acceptCurrentChange error", e);
  }
}

export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);
      // Clear highlight
      err.range.font.highlightColor = null;
      await context.sync();

      // Advance index
      state.currentIndex++;

      // Select next, if any
      if (state.currentIndex < state.errors.length) {
        const nextErr = state.errors[state.currentIndex];
        context.trackedObjects.add(nextErr.range);
        nextErr.range.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("rejectCurrentChange error", e);
  }
}

export async function acceptAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
        await context.sync();
      }
      state.currentIndex = state.errors.length;
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
  }
}

export async function rejectAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.font.highlightColor = null;
        await context.sync();
      }
      state.currentIndex = state.errors.length;
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
  }
}
