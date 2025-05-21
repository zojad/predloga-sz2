/* global Office, Word */

let state = {
  errors: [],         // { range: Word.Range, suggestion: "s"|"z" }[]
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

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
  const numMap   = {'1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n'};
  return (/\d/.test(first) ? unvoiced.has(numMap[first]) : unvoiced.has(first))
    ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: ALWAYS reset & rescan on each click
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  // **Reset** on every click:
  clearNotification(NOTIF_ID);
  state.errors.forEach(e => e.range.font.highlightColor = null);
  state.errors = [];
  state.currentIndex = 0;

  if (state.isChecking) return;
  state.isChecking = true;

  try {
    await Word.run(async context => {
      // Clear previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => ["s","z"].includes(r.text.trim().toLowerCase()));

      const errors = [];
      for (const prep of candidates) {
        const after = prep.getRange("After");
        const nxtR = after.getNextTextRange(
          [" ", "\n", ".", ",", ";", "?", "!"], 
          true
        );
        nxtR.load("text");
        await context.sync();

        const nxt = nxtR.text.trim();
        if (!nxt) continue;

        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          errors.push({ range: prep, suggestion: expect });
        }
      }

      state.errors = errors;

      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No 's'/'z' mismatches!",
          icon: "Icon.80x80",
          persistent: false
        });
      } else {
        errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
        await context.sync();
        errors[0].range.select();
      }
    });
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification("checkError", {
      type: "errorMessage",
      message: "Check failed; please try again.",
      persistent: false
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

  const err = state.errors[state.currentIndex];
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();

      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("acceptCurrentChange error", e);
  }
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear current & auto-advance
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  const err = state.errors[state.currentIndex];
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.font.highlightColor = null;
      await context.sync();

      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("rejectCurrentChange error", e);
  }
}

// ─────────────────────────────────────────────────
// 4) Accept All: replace every mismatch at once
// ─────────────────────────────────────────────────
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
      state.errors = [];
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
  }
}

// ─────────────────────────────────────────────────
// 5) Reject All: clear all highlights at once
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
  }
}

