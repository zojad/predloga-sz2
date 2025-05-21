/* global Office, Word */

let state = {
  errors: [],        // { range: Word.Range, suggestion: "s"|"z" }[]
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

/**
 * Decide between "s" vs "z":
 *  - next letter unvoiced ⇒ "s"
 *  - otherwise ⇒ "z"
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p',
                       '6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: populate state.errors, highlight all, select first
//    — now resets & rescans on every click
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  // **NEW**: clear previous notification & queue on every invocation
  clearNotification(NOTIF_ID);
  state.errors = [];

  if (state.isChecking) return;
  state.isChecking = true;

  try {
    await Word.run(async context => {
      // A) clear any old highlights
      const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // B) find every standalone s or z
      const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // C) evaluate each candidate
      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim().toLowerCase();
        if (raw !== "s" && raw !== "z") continue;

        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const actual = raw;
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          // track, highlight, queue
          context.trackedObjects.add(r);
          r.font.highlightColor = HIGHLIGHT_COLOR;
          state.errors.push({ range: r, suggestion: expect });
        }
      }

      await context.sync();

      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No 's'/'z' mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        // select the first mismatch
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification("checkError", {
      type: "errorMessage",
      message: "Check failed; please try again."
    });
  } finally {
    state.isChecking = false;
  }
}

// ─────────────────────────────────────────────────
// 2) Accept One: fix the first queued item, then re-check
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;

  const { range, suggestion } = state.errors.shift();
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // re-scan to rediscover & select the next one
  await checkDocumentText();
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear the first queued item, then re-check
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (!state.errors.length) return;

  const { range } = state.errors.shift();
  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  // re-scan to rediscover & select the next one
  await checkDocumentText();
}

// ─────────────────────────────────────────────────
// 4) Accept All: one-shot fresh scan → replace every mismatch
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    for (const r of [...sRes.items, ...zRes.items]) {
      const raw = r.text.trim().toLowerCase();
      if (raw !== "s" && raw !== "z") continue;

      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();
      const nxt = after.text.trim();
      if (!nxt) continue;

      const expect = determineCorrectPreposition(nxt);
      if (expect && raw !== expect) {
        context.trackedObjects.add(r);
        r.insertText(expect, Word.InsertLocation.replace);
        r.font.highlightColor = null;
      }
    }
    await context.sync();
    state.errors = [];
  });

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: one-shot fresh scan → clear every highlight
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    for (const r of [...sRes.items, ...zRes.items]) {
      const raw = r.text.trim().toLowerCase();
      if (raw === "s" || raw === "z") {
        context.trackedObjects.add(r);
        r.font.highlightColor = null;
      }
    }
    await context.sync();
    state.errors = [];
  });

  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Cleared all!",
    icon: "Icon.80x80"
  });
}

