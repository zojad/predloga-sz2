/* global Office, Word */

// –– State ––
let state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// –– Helpers ––
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

// –– Core logic helper ––
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

// –– Exposed commands ––//

export async function checkDocumentText() {
  console.log("▶ checkDocumentText()", state);
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      console.log("→ Word.run(start)");

      // clear any prior highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      // do two whole-word searches: “s” and “z”
      const opts = { matchCase: false, matchWholeWord: true };
      const sSearch = context.document.body.search("s", opts);
      const zSearch = context.document.body.search("z", opts);
      sSearch.load("items");
      zSearch.load("items");
      await context.sync();

      const allRanges = [...sSearch.items, ...zSearch.items];
      console.log("→ raw hits:", allRanges.length);

      // filter down to exact lowercase “s” or “z”
      const candidates = allRanges.filter(r =>
        ["s","z"].includes(r.text.trim().toLowerCase())
      );
      console.log("→ filtered candidates:", candidates.length);

      // now inspect each candidate
      const errors = [];
      for (let prep of candidates) {
        const after = prep.getRange(Word.RangeLocation.After);
        after.expandTo(Word.TextRangeUnit.Word);
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
      console.log("→ mismatches found:", errors.length);

      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "🎉 No ‘s’/‘z’ mismatches!",
          icon: "Icon.80x80",
          persistent: false
        });
      } else {
        // highlight and select the first one
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
  console.log("▶ acceptCurrentChange()", state.currentIndex, "of", state.errors.length);
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
      console.log("→ moved to index", state.currentIndex);
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
  console.log("▶ rejectCurrentChange()", state.currentIndex);
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
      console.log("→ moved to index", state.currentIndex);
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
  console.log("▶ acceptAllChanges()", state.errors.length);
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
      console.log("→ accepted all");
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
  console.log("▶ rejectAllChanges()", state.errors.length);
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      state.errors.forEach(e => e.range.font.highlightColor = null);
      await context.sync();
      state.errors = [];
      console.log("→ rejected all");
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
