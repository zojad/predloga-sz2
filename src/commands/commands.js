/* global Office, Word */

// State for errors and control flow
const state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

// Highlight color for detected errors: light pink
const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID = "noErrors";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    try {
      Office.actions.associate("checkDocumentText", checkDocumentText);
      Office.actions.associate("acceptAllChanges", acceptAllChanges);
      Office.actions.associate("rejectAllChanges", rejectAllChanges);
      Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
      Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    } catch (error) {
      console.error("Function registration failed:", error);
      Office.NotificationMessages.addAsync("regError", {
        type: "errorMessage",
        message: "Add-in initialization failed. Please reload.",
        persistent: false
      });
    }
  }
});

// Map a raw word (letter or digit start) to the correct 's' or 'z'
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  // Normalize to NFC so É vs É are consistent
  const word = rawWord.normalize("NFC");
  // Find first letter or digit via Unicode property escape
  const match = word.match(/[\p{L}0-9]/u);
  if (!match) return null;
  const firstChar = match[0].toLowerCase();

  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numPron = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };

  if (/\d/.test(firstChar)) {
    const pron = numPron[firstChar];
    return unvoiced.has(pron) ? "s" : "z";
  }
  return unvoiced.has(firstChar) ? "s" : "z";
}

// Main scan: highlight mismatches, or notify if none
async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  Office.NotificationMessages.deleteAsync(NOTIF_ID);

  try {
    await Word.run(async (context) => {
      // Clear previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const searchOptions = { matchCase: false, matchWholeWord: true };
      let allRanges = [];

      // helper: find standalone 's' or 'z'
      async function addSearchResults(scope) {
        const res = scope.search("\\b[sz]\\b", searchOptions);
        res.load("items");
        await context.sync();
        allRanges.push(...res.items);
      }

      // Body, headers, footers, content controls, tables
      await addSearchResults(context.document.body);
      const secs = context.document.sections;
      secs.load("items"); await context.sync();
      for (const s of secs.items) {
        await addSearchResults(s.getHeader("Primary"));
        await addSearchResults(s.getFooter("Primary"));
      }
      const ccs = context.document.contentControls;
      ccs.load("items"); await context.sync();
      for (const cc of ccs.items) await addSearchResults(cc);
      const tables = context.document.body.tables;
      tables.load("items"); await context.sync();
      for (const t of tables.items) await addSearchResults(t.getRange());

      // Filter exactly “s”/“z”
      const candidates = allRanges.filter(r =>
        ["s","z"].includes(r.text.trim().toLowerCase())
      );

      const errors = [];
      for (const prep of candidates) {
        const after = prep.getRange("After");
        after.expandTo(Word.TextRangeUnit.word);
        after.load("text");
        await context.sync();

        const nextWord = after.text.trim();
        if (!nextWord) continue;

        const curr = prep.text.trim().toLowerCase();
        const corr = determineCorrectPreposition(nextWord);
        if (corr && curr !== corr) errors.push({range: prep, suggestion: corr});
      }

      state.errors = errors;
      if (errors.length === 0) {
        Office.NotificationMessages.addAsync(NOTIF_ID, {
          type: "informationalMessage",
          message: "🎉 No mismatched ‘s’/‘z’ prepositions found.",
          icon: "Icon.80x80",
          persistent: false
        });
        return;
      }

      // Highlight our errors and select the first
      for (const e of errors) e.range.font.highlightColor = HIGHLIGHT_COLOR;
      await context.sync();
      errors[0].range.select();
    });
  } catch (e) {
    console.error("checkDocumentText failed:", e);
    Office.NotificationMessages.addAsync("checkError", {
      type: "errorMessage",
      message: "Preposition check failed. Please try again.",
      persistent: false
    });
  } finally {
    state.isChecking = false;
  }
}

// Accept/reject functions
async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  try {
    await Word.run(async (context) => {
      const err = state.errors[state.currentIndex];
      try {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      } catch {
        await checkDocumentText(); // resync
        return;
      }
      await context.sync();
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (e) {
    console.error("acceptCurrentChange failed:", e);
    Office.NotificationMessages.addAsync("acceptError", {
      type: "errorMessage",
      message: "Failed to apply change. Please re-run the check.",
      persistent: false
    });
  }
}

async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  try {
    await Word.run(async (context) => {
      const err = state.errors[state.currentIndex];
      err.range.font.highlightColor = null;
      await context.sync();
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (e) {
    console.error("rejectCurrentChange failed:", e);
    Office.NotificationMessages.addAsync("rejectError", {
      type: "errorMessage",
      message: "Failed to reject change. Please re-run the check.",
      persistent: false
    });
  }
}

async function acceptAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("acceptAllChanges failed:", e);
    Office.NotificationMessages.addAsync("acceptAllError", {
      type: "errorMessage",
      message: "Failed to apply all changes. Please try again.",
      persistent: false
    });
  }
}

async function rejectAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("rejectAllChanges failed:", e);
    Office.NotificationMessages.addAsync("rejectAllError", {
      type: "errorMessage",
      message: "Failed to clear changes. Please try again.",
      persistent: false
    });
  }
}

// Expose to ribbon/UI
window.checkDocumentText   = checkDocumentText;
window.acceptCurrentChange = acceptCurrentChange;
window.rejectCurrentChange = rejectCurrentChange;
window.acceptAllChanges    = acceptAllChanges;
window.rejectAllChanges    = rejectAllChanges;
