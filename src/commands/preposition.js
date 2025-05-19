/* global Office, Word */

let state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID       = "noErrors";

//–– Helpers ––//
function clearNotification(id) {
  if (
    Office.NotificationMessages &&
    typeof Office.NotificationMessages.deleteAsync === "function"
  ) {
    Office.NotificationMessages.deleteAsync(id);
  }
}

function showNotification(id, options) {
  if (
    Office.NotificationMessages &&
    typeof Office.NotificationMessages.addAsync === "function"
  ) {
    Office.NotificationMessages.addAsync(id, options);
  }
}

//–– Logic ––//
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

/**
 * Checks the doc for standalone “s”/“z” and highlights mismatches.
 */
export async function checkDocumentText(event) {
  // prevent re-entry
  if (state.isChecking) {
    event.completed();
    return;
  }
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // clear old highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const opts = { matchCase: false, matchWholeWord: true };
      let allRanges = [];

      async function find(szScope) {
        const r = szScope.search("\\b[sz]\\b", opts);
        r.load("items");
        await context.sync();
        allRanges.push(...r.items);
      }

      await find(context.document.body);

      // filter exactly "s" or "z"
      const candidates = allRanges.filter(r =>
        ["s","z"].includes(r.text.trim().toLowerCase())
      );

      let errors = [];
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
      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "🎉 No ‘s’/‘z’ mismatches!",
          icon: "Icon.80x80",
          persistent: false
        });
        return;
      }

      // highlight and select first
      errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
      await context.sync();
      errors[0].range.select();
    });
  } catch (e) {
    console.error(e);
    showNotification("checkError", {
      type: "errorMessage",
      message: "Check failed; please try again.",
      persistent: false
    });
  } finally {
    state.isChecking = false;
    event.completed();
  }
}

/**
 * Replaces the current error with its suggestion and advances to the next.
 */
export async function acceptCurrentChange(event) {
  if (state.currentIndex >= state.errors.length) {
    event.completed();
    return;
  }

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
    console.error(e);
    showNotification("acceptError", {
      type: "errorMessage",
      message: "Failed to apply change. Please re-run the check.",
      persistent: false
    });
  } finally {
    event.completed();
  }
}

/**
 * Clears highlight on current error (i.e. “rejects” it) and advances.
 */
export async function rejectCurrentChange(event) {
  if (state.currentIndex >= state.errors.length) {
    event.completed();
    return;
  }

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
    console.error(e);
    showNotification("rejectError", {
      type: "errorMessage",
      message: "Failed to reject change. Please re-run the check.",
      persistent: false
    });
  } finally {
    event.completed();
  }
}

/**
 * Applies all suggestions in one go.
 */
export async function acceptAllChanges(event) {
  if (!state.errors.length) {
    event.completed();
    return;
  }

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
    console.error(e);
    showNotification("acceptAllError", {
      type: "errorMessage",
      message: "Failed to apply all changes. Please try again.",
      persistent: false
    });
  } finally {
    event.completed();
  }
}

/**
 * Clears all highlights (i.e. “rejects” everything).
 */
export async function rejectAllChanges(event) {
  if (!state.errors.length) {
    event.completed();
    return;
  }

  try {
    await Word.run(async context => {
      state.errors.forEach(e => e.range.font.highlightColor = null);
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error(e);
    showNotification("rejectAllError", {
      type: "errorMessage",
      message: "Failed to clear changes. Please try again.",
      persistent: false
    });
  } finally {
    event.completed();
  }
}
