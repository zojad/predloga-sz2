/* global Office, Word */

// State
const state = {
  errors: [],
  currentIndex: 0
};

// Office initialization
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("checkDocumentText", checkDocumentText);
    Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
    Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    Office.actions.associate("acceptAllChanges", acceptAllChanges);
    Office.actions.associate("rejectAllChanges", rejectAllChanges);
  }
});

// Determines correct preposition (supports letters + numbers)
function determineCorrectPreposition(word) {
  if (!word) return null;

  const unvoiced = new Set(['c', 'č', 'f', 'h', 'k', 'p', 's', 'š', 't']);
  const digitPronunciation = {
    '1': 'e', '2': 'd', '3': 't', '4': 'š', '5': 'p',
    '6': 'š', '7': 's', '8': 'o', '9': 'd', '0': 'n'
  };

  let firstChar = '';
  for (const char of word) {
    if (char.match(/[a-zA-ZčČšŠžŽ0-9]/)) {
      firstChar = char.toLowerCase();
      break;
    }
  }

  if (!firstChar) return null;

  if (/\d/.test(firstChar)) {
    const sound = digitPronunciation[firstChar];
    return unvoiced.has(sound) ? 's' : 'z';
  }

  return unvoiced.has(firstChar) ? 's' : 'z';
}

// Check document for errors
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

      const errorCandidates = [...sResults.items, ...zResults.items]
        .filter(prep => ['s', 'z'].includes(prep.text.trim().toLowerCase()))
        .map(prep => {
          const next = prep.getNextTextRange("Word");
          if (next) next.load("text");
          return { prepositionRange: prep, nextWordRange: next };
        });

      await context.sync();

      state.errors = errorCandidates
        .map(({ prepositionRange, nextWordRange }) => {
          const currentPrep = prepositionRange.text.trim().toLowerCase();
          const nextWord = nextWordRange?.text?.trim?.();
          const correctPrep = determineCorrectPreposition(nextWord);
          return correctPrep && currentPrep !== correctPrep
            ? { range: prepositionRange, suggestion: correctPrep }
            : null;
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

// Accept one
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

// Reject one
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

// Accept all
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

// Reject all
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

// For taskpane testing if needed
window.checkDocumentText = checkDocumentText;
window.acceptCurrentChange = acceptCurrentChange;
window.rejectCurrentChange = rejectCurrentChange;
window.acceptAllChanges = acceptAllChanges;
window.rejectAllChanges = rejectAllChanges;

