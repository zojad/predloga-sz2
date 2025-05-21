/* global Office, Word */

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Show or clear ribbon notifications
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

// Decide “s” vs “z” from the first letter of the next word
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

// ─────────────────────────────────────────────────
// 1) Check S/Z: highlight all mismatches & select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    // Clear any previous highlights
    const prevS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
    const prevZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
    prevS.load("items"); prevZ.load("items");
    await context.sync();
    [...prevS.items, ...prevZ.items].forEach(r => r.font.highlightColor = null);
    await context.sync();

    // Find all standalone s/z
    const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
    const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
    sRes.load("items"); zRes.load("items");
    await context.sync();

    const mismatches = [];
    for (const r of [...sRes.items, ...zRes.items]) {
      const txt = r.text.trim().toLowerCase();
      if (txt !== "s" && txt !== "z") continue;

      // Grab the very next word
      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();
      const nextWord = after.text.trim();
      if (!nextWord) continue;

      const expected = determineCorrectPreposition(nextWord);
      if (expected && txt !== expected) {
        r.font.highlightColor = HIGHLIGHT_COLOR;
        mismatches.push(r);
      }
    }

    if (!mismatches.length) {
      showNotification(NOTIF_ID, {
        type: "informationalMessage",
        message: "✨ No mismatches!",
        icon: "Icon.80x80"
      });
      return;
    }

    // Select the first mismatch
    const first = mismatches[0];
    first.select();
    await context.sync();
  });
}

// ─────────────────────────────────────────────────
// 2) Accept One: replace current mismatch & go to next
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  await Word.run(async context => {
    const sel = context.document.getSelection();
    sel.load("text");
    await context.sync();

    const actual = sel.text.trim().toLowerCase();
    if (actual !== "s" && actual !== "z") {
      // nothing to accept here
      return;
    }

    // Figure out the replacement
    const after = sel.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
    after.load("text");
    await context.sync();
    const nextWord = after.text.trim();
    const expected = determineCorrectPreposition(nextWord);
    if (expected && expected !== actual) {
      sel.insertText(expected, Word.InsertLocation.replace);
      sel.font.highlightColor = null;
      await context.sync();
    }

    // Now find the next mismatch after this spot
    const afterSel = sel.getRange("After");
    const ns = afterSel.search("s", { matchWholeWord: true, matchCase: false });
    const nz = afterSel.search("z", { matchWholeWord: true, matchCase: false });
    ns.load("items"); nz.load("items");
    await context.sync();

    for (const r of [...ns.items, ...nz.items]) {
      const t = r.text.trim().toLowerCase();
      if (t !== "s" && t !== "z") continue;
      const a2 = r.getRange("After").getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      a2.load("text");
      await context.sync();
      const nw = a2.text.trim();
      const exp2 = determineCorrectPreposition(nw);
      if (exp2 && exp2 !== t) {
        r.select();
        await context.sync();
        return;
      }
    }
  });
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear current highlight & go to next
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  await Word.run(async context => {
    const sel = context.document.getSelection();
    sel.font.highlightColor = null;
    await context.sync();

    // Find next mismatch just like in Accept One
    const afterSel = sel.getRange("After");
    const ns = afterSel.search("s", { matchWholeWord: true, matchCase: false });
    const nz = afterSel.search("z", { matchWholeWord: true, matchCase: false });
    ns.load("items"); nz.load("items");
    await context.sync();

    for (const r of [...ns.items, ...nz.items]) {
      const t = r.text.trim().toLowerCase();
      if (t !== "s" && t !== "z") continue;
      const a2 = r.getRange("After").getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      a2.load("text");
      await context.sync();
      const nw = a2.text.trim();
      const exp2 = determineCorrectPreposition(nw);
      if (exp2 && exp2 !== t) {
        r.select();
        await context.sync();
        return;
      }
    }
  });
}

// ─────────────────────────────────────────────────
// 4) Accept All: bulk replace everywhere
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    const all = [...sRes.items, ...zRes.items]
      .filter(r => ['s','z'].includes(r.text.trim().toLowerCase()));

    for (const r of all) {
      const actual = r.text.trim().toLowerCase();
      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();
      const nextWord = after.text.trim();
      const expected = determineCorrectPreposition(nextWord);
      if (expected && expected !== actual) {
        r.insertText(expected, Word.InsertLocation.replace);
        r.font.highlightColor = null;
        await context.sync();
      }
    }
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: clear all highlights
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    const all = [...sRes.items, ...zRes.items];
    all.forEach(r => r.font.highlightColor = null);
    await context.sync();
  });
}
