/* global Office, Word */

// — Log immediately when the task logic bundle loads —
console.log("⭐ preposition.js loaded");

let state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

//–– Helpers ––//
function clearNotification(id) {
  if (Office.NotificationMessages &&
      typeof Office.NotificationMessages.deleteAsync === "function") {
    Office.NotificationMessages.deleteAsync(id);
  }
}

function showNotification(id, options) {
  if (Office.NotificationMessages &&
      typeof Office.NotificationMessages.addAsync === "function") {
    Office.NotificationMessages.addAsync(id, options);
  }
}

//–– Core logic helper ––//
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const word = rawWord.normalize("NFC");
  const match = word.match(/[\p{L}0-9]/u);
  if (!match) return null;
  const first = match[0].toLowerCase();

  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap   = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };

  if (/\d/.test(first)) {
    return unvoiced.has(numMap[first]) ? "s" : "z";
  }
  return unvoiced.has(first) ? "s" : "z";
}

//–– Exposed commands ––//

export async function checkDocumentText() {
  console.log("checkDocumentText()", { errors: state.errors, isChecking: state.isChecking });
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      console.log("→ Word.run(checkDocumentText) start");

      // Clear previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      // **WILDCARD search** for any standalone “s” or “z”
      const searchOptions = {
        matchCase: false,
        matchWholeWord: true,
        includeWildcards: true
      };
      const foundRanges = context.document.body
        .search("[sz]", searchOptions);
      foundRanges.load("items");
      await context.sync();

      // Post-filter to exact “s” or “z”
      const candidates = foundRanges.items.filter(r =>
        ["s","z"].includes(r.text.trim().toLowerCase())
      );

      console.log("→ raw candidates:", candidates.length);

      // Now look at the next word to see if it really is a mismatch
      const errors = [];
      for (let prep of candidates) {
        const after = prep.getRange("After");
        after.expandTo(Word.TextRangeUnit.word);
        after.load("text");
        await context.sync();

        const nextWord = after.text.trim();
        if (!nextWord) continue;

        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nextWord);
        if (expect && actual !== expect) {
          errors.push({ range: prep, suggestion: expect });
        }
      }

      state.errors = errors;
      console.log("→ Found mismatches:", errors);

      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "🎉 No ‘s’/‘z’ mismatches!",
          icon: "Icon.80x80",
          persistent: false
        });
      } else {
        // Highlight + select first
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

export async function acceptCurrentChange() {
  console.log("acceptCurrentChange()", { currentIndex: state.currentIndex, total: state.errors.length });
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();

      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (e) {
    console.error("acceptCurrentChange error", e);
    showNotification("acceptError", {
      type: "errorMessage",
      message: "Failed to apply change. Please re-run the check.",
      persistent: false
    });
  }
}

export async function rejectCurrentChange() {
  console.log("rejectCurrentChange()", { currentIndex: state.currentIndex });
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const err = state.errors[state.currentIndex];
      err.range.font.highlightColor = null;
      await context.sync();

      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (e) {
    console.error("rejectCurrentChange error", e);
    showNotification("rejectError", {
      type: "errorMessage",
      message: "Failed to reject change. Please re-run the check.",
      persistent: false
    });
  }
}

export async function acceptAllChanges() {
  console.log("acceptAllChanges()", { total: state.errors.length });
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
    showNotification("acceptAllError", {
      type: "errorMessage",
      message: "Failed to apply all changes. Please try again.",
      persistent: false
    });
  }
}

export async function rejectAllChanges() {
  console.log("rejectAllChanges()", { total: state.errors.length });
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      state.errors.forEach(e => e.range.font.highlightColor = null);
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
    showNotification("rejectAllError", {
      type: "errorMessage",
      message: "Failed to clear changes. Please try again.",
      persistent: false
    });
  }
}
