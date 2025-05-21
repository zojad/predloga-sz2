/* global Office, Word */

let state = {
  errors: [],        // Array of { range: Word.Range, suggestion: "s"|"z" }
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

function clearNotification(id) {
  if (Office.NotificationMessages?.deleteAsync) {
    Office.NotificationMessages.deleteAsync(id);
  }
}
function showNotification(id, opts) {
  if (Office.NotificationMessages?.addAsync) {
    Office.NotificationMessages.addAsync(id, opts);
  }
}

function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return voiceless.has(key) ? "s" : "z";
}

export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // clear old highlights
      const oldS = context.document.body.search("s", { matchWholeWord:true, matchCase:false });
      const oldZ = context.document.body.search("z", { matchWholeWord:true, matchCase:false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // find new mismatches
      const sRes = context.document.body.search("s", { matchWholeWord:true, matchCase:false });
      const zRes = context.document.body.search("z", { matchWholeWord:true, matchCase:false });
      sRes.load("items"); zRes.load("items");
      await context.sync();

      for (const r of [...sRes.items, ...zRes.items]) {
        const t = r.text.trim().toLowerCase();
        if (t!=="s" && t!=="z") continue;
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;
        const expected = determineCorrectPreposition(nxt);
        if (expected && expected!==t) {
          r.font.highlightColor = HIGHLIGHT_COLOR;
          state.errors.push({ range: r, suggestion: expected });
        }
      }

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select first
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("check error", e);
  } finally {
    state.isChecking = false;
  }
}


export async function acceptCurrentChange() {
  console.log("▶ acceptCurrentChange fired; errors in queue:", state.errors.length, "currentIndex:", state.currentIndex);
  if (state.currentIndex >= state.errors.length) return;

  // pull out the current mismatch
  const { range, suggestion } = state.errors.splice(state.currentIndex, 1)[0];
  console.log("   replacing:", range.text, "→", suggestion);

  // 1) replace & clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // 2) select the next one, if present
  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}


export async function rejectCurrentChange() {
  console.log("▶ rejectCurrentChange fired; errors in queue:", state.errors.length, "currentIndex:", state.currentIndex);
  if (state.currentIndex >= state.errors.length) return;

  const { range } = state.errors.splice(state.currentIndex, 1)[0];
  console.log("   clearing highlight for:", range.text);

  // 1) clear highlight
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  // 2) select next
  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}


export async function acceptAllChanges() {
  console.log("▶ acceptAllChanges fired");
  // fresh scan so nothing is left behind
  await checkDocumentText();
  console.log("   after rescan, errors:", state.errors.length);
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const { range, suggestion } of state.errors) {
      context.trackedObjects.add(range);
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = null;
    }
    await context.sync();
  });
  state.errors = [];
  state.currentIndex = 0;
  showNotification(NOTIF_ID, { type: "informationalMessage", message: "Accepted all!" });
}


export async function rejectAllChanges() {
  console.log("▶ rejectAllChanges fired");
  // fresh scan so nothing is left behind
  await checkDocumentText();
  console.log("   after rescan, errors:", state.errors.length);
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const { range } of state.errors) {
      context.trackedObjects.add(range);
      range.font.highlightColor = null;
    }
    await context.sync();
  });
  state.errors = [];
  state.currentIndex = 0;
  showNotification(NOTIF_ID, { type: "informationalMessage", message: "Cleared all!" });
}

