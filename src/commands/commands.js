/* global Office, Word */

const state = {
  errors: [],
  currentIndex: 0,
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("checkDocumentText", checkDocumentText);
    Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
    Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    Office.actions.associate("acceptAllChanges", acceptAllChanges);
    Office.actions.associate("rejectAllChanges", rejectAllChanges);

    console.log("Ribbon commands registered.");
  }
});

const UNVOICED = new Set(["p", "t", "k", "f", "s", "š", "c", "č", "h"]);
const VOICED_OR_VOWELS = new Set(["b", "d", "g", "v", "z", "ž", "j", "m", "n", "l", "r", "a", "e", "i", "o", "u"]);

function determineCorrectPreposition(nextWord) {
  if (!nextWord) return null;
  const firstChar = [...nextWord.toLowerCase()].find((c) => /[a-zčšž]/i.test(c));
  if (!firstChar) return null;
  return UNVOICED.has(firstChar) ? "s" : "z";
}

async function checkDocumentText() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const ranges = body.getTextRanges([" "], true);
    ranges.load("items/text");
    await context.sync();

    for (const err of state.errors) {
      err.range.font.highlightColor = null;
    }

    state.errors = [];

    for (let i = 0; i < ranges.items.length - 1; i++) {
      const currentText = ranges.items[i].text.trim().toLowerCase();
      const nextWord = ranges.items[i + 1].text.trim();
      if (currentText !== "s" && currentText !== "z") continue;

      const correct = determineCorrectPreposition(nextWord);
      if (correct && correct !== currentText) {
        ranges.items[i].font.highlightColor = "yellow";
        state.errors.push({
          index: i,
          range: ranges.items[i],
          suggestion: correct,
        });
      }
    }

    await context.sync();

    if (state.errors.length > 0) {
      state.currentIndex = 0;
      state.errors[0].range.select();
    } else {
      body.insertComment("No incorrect 's/z' prepositions found.", "start");
    }
  });
}

async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  await Word.run(async (context) => {
    const error = state.errors[state.currentIndex];
    const ranges = context.document.body.getTextRanges([" "], true);
    ranges.load("items");
    await context.sync();

    const range = ranges.items[error.index];
    const originalText = range.text;
    const replacement = /^[A-ZČŠŽ]/.test(originalText)
      ? error.suggestion.toUpperCase()
      : error.suggestion;

    range.insertText(replacement, Word.InsertLocation.replace);
    range.font.highlightColor = null;

    await context.sync();

    state.errors.splice(state.currentIndex, 1);
    if (state.currentIndex >= state.errors.length) state.currentIndex = 0;
    if (state.errors.length > 0) state.errors[state.currentIndex].range.select();
  });
}

async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  await Word.run(async (context) => {
    const error = state.errors[state.currentIndex];
    error.range.font.highlightColor = null;
    await context.sync();

    state.errors.splice(state.currentIndex, 1);
    if (state.currentIndex >= state.errors.length) state.currentIndex = 0;
    if (state.errors.length > 0) state.errors[state.currentIndex].range.select();
  });
}

async function acceptAllChanges() {
  if (state.errors.length === 0) return;

  await Word.run(async (context) => {
    const ranges = context.document.body.getTextRanges([" "], true);
    ranges.load("items");
    await context.sync();

    for (const err of state.errors) {
      const range = ranges.items[err.index];
      const originalText = range.text;
      const replacement = /^[A-ZČŠŽ]/.test(originalText)
        ? err.suggestion.toUpperCase()
        : err.suggestion;

      range.insertText(replacement, Word.InsertLocation.replace);
      range.font.highlightColor = null;
    }

    await context.sync();
    state.errors = [];
  });
}

async function rejectAllChanges() {
  if (state.errors.length === 0) return;

  await Word.run(async (context) => {
    for (const err of state.errors) {
      err.range.font.highlightColor = null;
    }

    await context.sync();
    state.errors = [];
  });
}

// Make functions available to taskpane
window.checkDocumentText = checkDocumentText;
window.acceptCurrentChange = acceptCurrentChange;
window.rejectCurrentChange = rejectCurrentChange;
window.acceptAllChanges = acceptAllChanges;
window.rejectAllChanges = rejectAllChanges;
