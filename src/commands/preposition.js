/* global Office, Word */

const state = {
  errors: [],        // [{ range: Word.Range, suggestion: "s"|"z" }, …]
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

// Pure-JS: decide whether the next word wants "s" or "z"
function determineCorrectPreposition(word) {
  if (!word) return null;
  const m = word.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const voiceless = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap   = { '1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n' };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return voiceless.has(key) ? "s" : "z";
}

// ───────────────────────────────────────
// 1) Highlight all mismatches & select first
// ───────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);
  state.errors = [];

  await Word.run(async context => {
    // 1. clear old highlights
    context.document.body.search("s", { matchWholeWord: true, matchCase: false })
      .load("items");
    context.document.body.search("z", { matchWholeWord: true, matchCase: false })
      .load("items");
    await context.sync();
    // Actually clear them:
    const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false }).items;
    const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false }).items;
    [...oldS, ...oldZ].forEach(r => r.font.highlightColor = null);
    await context.sync();

    // 2. find fresh candidates
    const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
    const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
    sRes.load("items"); zRes.load("items");
    await context.sync();

    const candidates = [...sRes.items, ...zRes.items]
      .filter(r => {
        const t = r.text.trim().toLowerCase();
        return t === "s" || t === "z";
      });

    // 3. for each candidate, grab the next word and compare
    for (const r of candidates) {
      const after = r.getRange("After")
                     .getNextTextRange(
                       [" ", "\n", ".", ",", ";", "?", "!"],
                       /*trimSpacing=*/ true
                     );
      after.load("text");
      await context.sync();
      const nxt = after.text.trim();
      if (!nxt) continue;

      const actual   = r.text.trim().toLowerCase();
      const expected = determineCorrectPreposition(nxt);
      if (expected && actual !== expected) {
        context.trackedObjects.add(r);
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
      // select the first mismatch
      const first = state.errors[0].range;
      context.trackedObjects.add(first);
      first.select();
      await context.sync();
    }
  });
}

// ───────────────────────────────────────
// 2) Accept One → replace & jump to next
// ───────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;

  // Pull off the head of the queue:
  const { suggestion } = state.errors.shift();

  await Word.run(async context => {
    // We're already sitting on that letter (from checkDocumentText)
    const sel = context.document.getSelection();
    sel.insertText(suggestion, Word.InsertLocation.replace);
    sel.font.highlightColor = null;

    // Now if anything remains, select the next mismatch
    if (state.errors.length) {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
    }

    await context.sync();
  });
}

// ───────────────────────────────────────
// 3) Reject One → clear highlight & jump
// ───────────────────────────────────────
export async function rejectCurrentChange() {
  if (!state.errors.length) return;

  // Just remove this one from the queue
  state.errors.shift();

  await Word.run(async context => {
    const sel = context.document.getSelection();
    sel.font.highlightColor = null;

    if (state.errors.length) {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
    }

    await context.sync();
  });
}

// ───────────────────────────────────────
// 4) Accept All → fix ‘em all in one go
// ───────────────────────────────────────
export async function acceptAllChanges() {
  if (!state.errors.length) return;

  await Word.run(async context => {
    // Load each range's text so we know what to replace
    for (const e of state.errors) {
      context.load(e.range, "text");
    }
    await context.sync();

    // Replace & clear
    for (const { range, suggestion } of state.errors) {
      range.insertText(suggestion, Word.InsertLocation.replace);
      range.font.highlightColor = null;
    }
    await context.sync();
  });

  state.errors = [];
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ───────────────────────────────────────
// 5) Reject All → clear all highlights
// ───────────────────────────────────────
export async function rejectAllChanges() {
  if (!state.errors.length) return;

  await Word.run(async context => {
    for (const e of state.errors) {
      e.range.font.highlightColor = null;
    }
    await context.sync();
  });

  state.errors = [];
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Cleared all!",
    icon: "Icon.80x80"
  });
}

