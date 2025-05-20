/* global Office, Word */

// state holds the list of mismatched ranges and which one we're on
const state = {
  errors: [],        // Word.Range[] of each bad “s”/“z”
  currentIndex: 0,   // which error is “current”
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Utilities for showing/clearing the little Office ribbon notifications
function clearNotification(id) {
  if (Office.NotificationMessages?.deleteAsync) {
    Office.NotificationMessages.deleteAsync(id);
  }
}
function showNotification(id, options) {
  if (Office.NotificationMessages?.addAsync) {
    Office.NotificationMessages.addAsync(id, options);
  }
}

// Your logic for choosing “s” vs “z” based on the next word’s first character
function determineCorrectPreposition(nextWord) {
  if (!nextWord) return null;
  const m = nextWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const first = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  if (/\d/.test(first)) {
    // Map digits to their Slovene‐sound letter, then treat that as voiceless/voiced
    const digitMap = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
    return voiceless.has(digitMap[first]) ? "s" : "z";
  }
  return voiceless.has(first) ? "s" : "z";
}

// ----------------------
// 1) Check & highlight
// ----------------------
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // 1. Undo any old highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);

      // 2. Find every standalone “s” or “z” (both cases)
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      const candidates = [...sRes.items, ...zRes.items]
        .filter(r => /^[sSzZ]$/.test(r.text));

      // 3. For each one, pull the very next word and compare
      for (let r of candidates) {
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const actual = r.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          // track & highlight it
          context.trackedObjects.add(r);
          r.font.highlightColor = HIGHLIGHT_COLOR;
          state.errors.push(r);
        }
      }

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the first mismatch so the user sees where to click “Accept”
        const first = state.errors[0];
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error(e);
    showNotification(NOTIF_ID, { type: "errorMessage", message: "Check failed" });
  } finally {
    state.isChecking = false;
  }
}

// ---------------------------------
// 2) Accept one (replace & advance)
// ---------------------------------
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const r = state.errors[state.currentIndex];
      context.trackedObjects.add(r);

      // load the wrong letter
      context.load(r, "text");
      await context.sync();
      const wrong = r.text;
      // decide the correct letter, respecting uppercase
      const corr =
        wrong === "s" ? "z" :
        wrong === "S" ? "Z" :
        wrong === "z" ? "s" :
        wrong === "Z" ? "S" :
        wrong;

      r.insertText(corr, Word.InsertLocation.replace);
      r.font.highlightColor = null;

      // advance index and select next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nxt = state.errors[state.currentIndex];
        context.trackedObjects.add(nxt);
        nxt.select();
      }

      await context.sync();
    });

    // remove the one we fixed
    state.errors.splice(state.currentIndex - 1, 1);
    if (state.currentIndex > state.errors.length) {
      state.currentIndex = 0;
    }
  } catch (e) {
    console.error(e);
  }
}

// ---------------------------------
// 3) Reject one (clear & advance)
// ---------------------------------
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      const r = state.errors[state.currentIndex];
      context.trackedObjects.add(r);

      r.font.highlightColor = null;

      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nxt = state.errors[state.currentIndex];
        context.trackedObjects.add(nxt);
        nxt.select();
      }

      await context.sync();
    });

    // remove the one we skipped
    state.errors.splice(state.currentIndex - 1, 1);
    if (state.currentIndex > state.errors.length) {
      state.currentIndex = 0;
    }
  } catch (e) {
    console.error(e);
  }
}

// --------------------------------
// 4) Accept All (bulk replace all)
// --------------------------------
export async function acceptAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      // load all texts
      for (let r of state.errors) context.load(r, "text");
      await context.sync();

      // replace & clear all
      for (let r of state.errors) {
        const w = r.text;
        const c =
          w === "s" ? "z" :
          w === "S" ? "Z" :
          w === "z" ? "s" :
          w === "Z" ? "S" :
          w;
        r.insertText(c, Word.InsertLocation.replace);
        r.font.highlightColor = null;
      }
      await context.sync();
    });

    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Accepted all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error(e);
  }
}

// ---------------------------------
// 5) Reject All (bulk clear highlights)
// ---------------------------------
export async function rejectAllChanges() {
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      for (let r of state.errors) {
        r.font.highlightColor = null;
      }
      await context.sync();
    });

    state.errors = [];
    state.currentIndex = 0;
    showNotification(NOTIF_ID, {
      type: "informationalMessage",
      message: "Cleared all!",
      icon: "Icon.80x80"
    });
  } catch (e) {
    console.error(e);
  }
}

