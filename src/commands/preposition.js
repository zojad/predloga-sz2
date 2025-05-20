/* global Office, Word */

let state = {
  errors: [],
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
  const word = rawWord.normalize("NFC");
  const match = word.match(/[\p{L}0-9]/u);
  if (!match) return null;
  const first = match[0].toLowerCase();

  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap   = {
    '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n'
  };

  if (/\d/.test(first)) {
    return unvoiced.has(numMap[first]) ? "s" : "z";
  }
  return unvoiced.has(first) ? "s" : "z";
}

export async function checkDocumentText() {
  console.log("▶ checkDocumentText()", state);
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      console.log("→ Word.run(start)");

      // clear any previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      // find every standalone "s" or "z"
      const opts    = { matchCase: false, matchWholeWord: true };
      const sSearch = context.document.body.search("s", opts);
      const zSearch = context.document.body.search("z", opts);
      sSearch.load("items");
      zSearch.load("items");
      await context.sync();

      const allRanges   = [...sSearch.items, ...zSearch.items];
      const candidates  = allRanges.filter(r => ["s","z"].includes(r.text.trim().toLowerCase()));
      console.log("→ found", candidates.length, "s/z candidates");

      const errors = [];
      for (const prep of candidates) {
        // get a zero-length range right after the "s" or "z"
        const after = prep.getRange("After");

        // expand it to capture the next word up to whitespace/punctuation
        const nextWordRange = after.getNextTextRange(
          [" ", "\n", ".", ",", ";", "?", "!"], /* trimSpacing= */ true
        );
        nextWordRange.load("text");
        await context.sync();

        const nextWord = nextWordRange.text.trim();
        if (!nextWord) continue;

        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nextWord);
        if (expect && actual !== expect) {
          // *track* this Range so later batches can modify it
          context.trackedObjects.add(prep);
          errors.push({ range: prep, suggestion: expect });
        }
      }

      state.errors = errors;
      console.log("→ mismatches found:", errors.length);

      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No 's'/'z' mismatches!",
          icon: "Icon.80x80",
          persistent: false
        });
      } else {
        // highlight them and select the first one
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
  }
}

export async function rejectCurrentChange() {
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
  }
}

export async function acceptAllChanges() {
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
  }
}

export async function rejectAllChanges() {
  if (!state.errors.length) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (e) {
    console.error("rejectAllChanges error", e);
  }
}
