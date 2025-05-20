/* global Office, Word */

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID = "noErrors";

let state = {
  errors: [],
  currentIndex: 0,
  isChecking: false,
};

// –– Helpers ––
function clearNotification(id) {
  if (Office?.NotificationMessages?.deleteAsync) {
    Office.NotificationMessages.deleteAsync(id);
  }
}

function showNotification(id, options) {
  if (Office?.NotificationMessages?.addAsync) {
    Office.NotificationMessages.addAsync(id, options);
  }
}

// –– Logic to choose correct “s” or “z” ––
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;

  const word = rawWord.normalize("NFC");
  const match = word.match(/[\p{L}0-9]/u);
  if (!match) return null;

  const first = match[0].toLowerCase();
  const unvoiced = new Set(["c", "č", "f", "h", "k", "p", "s", "š", "t"]);
  const numMap = {
    "1": "e", "2": "d", "3": "t", "4": "š", "5": "p",
    "6": "š", "7": "s", "8": "o", "9": "d", "0": "n"
  };

  if (/\d/.test(first)) {
    return unvoiced.has(numMap[first]) ? "s" : "z";
  }

  return unvoiced.has(first) ? "s" : "z";
}

// –– Main Command ––
export async function checkDocumentText() {
  console.log("▶ checkDocumentText()", state);
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const opts = { matchCase: false, matchWholeWord: true };
      const allRanges = [];

      async function find(scope) {
        const s = scope.search("s", opts);
        const z = scope.search("z", opts);
        s.load("items"); z.load("items");
        await context.sync();
        allRanges.push(...s.items, ...z.items);
      }

      // Scan body
      await find(context.document.body);

      // Scan all table cells
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      for (const table of tables.items) {
        for (let r = 0; r < table.rowCount; r++) {
          for (let c = 0; c < table.columnCount; c++) {
            const cell = table.getCell(r, c);
            await find(cell.body);
          }
        }
      }

      const candidates = allRanges.filter(r =>
        ["s", "z"].includes(r.text.trim().toLowerCase())
      );

      console.log(`→ found ${candidates.length} s/z candidates`);

      const errors = [];

      for (let prep of candidates) {
        const after = prep.getRange(Word.RangeLocation.After);
        after.expandTo(Word.TextRangeUnit.Word);
        after.load("text");
        await context.sync();

        const nextWord = after.text.replace(/^[\s.,;:!?]+/, "").trim();
        const actual = prep.text.trim().toLowerCase();
        const expected = determineCorrectPreposition(nextWord);

        if (expected && actual !== expected) {
          errors.push({ range: prep, suggestion: expected });
        }
      }

      state.errors = errors;
      console.log(`→ mismatches found: ${errors.length}`);

      if (!errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "🎉 No ‘s’/‘z’ mismatches!",
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

// –– Accept/Reject Commands ––
export async function acceptCurrentChange() {
  console.log("▶ acceptCurrentChange()", state.currentIndex, state.errors.length);
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
    });
  } catch (e) {
    console.error("rejectCurrentChange error", e);
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
    });
  } catch (e) {
    console.error("acceptAllChanges error", e);
  }
}

export async function rejectAllChanges() {
  console.log("▶ rejectAllChanges()", state.errors.length);
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
