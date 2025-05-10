/* global Office, Word */

const state = {
  errors: [],
  currentIndex: 0,
};

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
    }
  }
});

function determineCorrectPreposition(word) {
  if (!word) return null;

  const unvoicedConsonants = new Set(['c', 'č', 'f', 'h', 'k', 'p', 's', 'š', 't']);
  const numberPronunciations = {
    '1': 'e', '2': 'd', '3': 't', '4': 'š', '5': 'p',
    '6': 'š', '7': 's', '8': 'o', '9': 'd', '0': 'n'
  };

  let firstChar = "";
  for (const char of word) {
    if (char.match(/[a-zA-ZčČšŠžŽ0-9]/)) {
      firstChar = char.toLowerCase();
      break;
    }
  }

  if (!firstChar) return null;

  if (firstChar >= '0' && firstChar <= '9') {
    const pronunciation = numberPronunciations[firstChar];
    return unvoicedConsonants.has(pronunciation) ? "s" : "z";
  }

  return unvoicedConsonants.has(firstChar) ? "s" : "z";
}

async function checkDocumentText() {
  try {
    await Word.run(async (context) => {
      state.errors.forEach(err => {
        err.range.font.highlightColor = null;
      });
      state.errors = [];
      state.currentIndex = 0;

      const searchOptions = { matchCase: false, matchWholeWord: true };
      const sResults = context.document.body.search("s", searchOptions);
      const zResults = context.document.body.search("z", searchOptions);
      sResults.load("items");
      zResults.load("items");
      await context.sync();

      const errors = [...sResults.items, ...zResults.items]
        .filter(prep => ['s', 'z'].includes(prep.text.trim().toLowerCase()))
        .map(prep => {
          if (!prep || typeof prep.getNextTextRange !== "function") return null;

          try {
            const nextRange = prep.getNextTextRange(Word.TextRangeUnit.word);
            return {
              prepositionRange: prep,
              nextWordRange: nextRange,
            };
          } catch (e) {
            console.warn("Failed to get next text range for:", prep.text);
            return null;
          }
        })
        .filter(Boolean);

      errors.forEach(e => e.nextWordRange.load("text"));
      await context.sync();

      state.errors = errors
        .map(({ prepositionRange, nextWordRange }) => {
          const currentPrep = prepositionRange.text.trim().toLowerCase();
          const correctPrep = determineCorrectPreposition(nextWordRange.text.trim());
          return correctPrep && currentPrep !== correctPrep ? {
            range: prepositionRange,
            suggestion: correctPrep
          } : null;
        })
        .filter(Boolean);

      state.errors.forEach(err => {
        err.range.font.highlightColor = "Yellow";
      });

      await context.sync();

      if (state.errors.length > 0) {
        state.errors[0].range.select();
      } else {
        context.document.body.insertComment("No preposition errors found.", "start");
      }
    });
  } catch (error) {
    console.error("Document check failed:", error);
  }
}

async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async (context) => {
      const err = state.errors[state.currentIndex];
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();
      state.currentIndex++;

      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (error) {
    console.error("Failed to accept change:", error);
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
  } catch (error) {
    console.error("Failed to reject change:", error);
  }
}

async function acceptAllChanges() {
  if (state.errors.length === 0) return;

  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (error) {
    console.error("Failed to accept all changes:", error);
  }
}

async function rejectAllChanges() {
  if (state.errors.length === 0) return;

  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (error) {
    console.error("Failed to reject all changes:", error);
  }
}

window.checkDocumentText = checkDocumentText;
window.acceptAllChanges = acceptAllChanges;
window.rejectAllChanges = rejectAllChanges;
window.acceptCurrentChange = acceptCurrentChange;
window.rejectCurrentChange = rejectCurrentChange;
