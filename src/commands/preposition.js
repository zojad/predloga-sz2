/* global Office, Word */

let state = {
  errors: []   // Array<{ range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }>
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Ribbon‐notification helpers
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
 * Decide “s” vs “z”:
 *  - unvoiced consonants (c,č,f,h,k,p,s,š,t) ⇒ "s"
 *  - otherwise ⇒ "z"
 */
function determineCorrectPreposition(rawWord) {
  if (!rawWord) return null;
  const m = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!m) return null;
  const c = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const digitMap = {
    '1':'e','2':'d','3':'t','4':'š','5':'p',
    '6':'š','7':'s','8':'o','9':'d','0':'n'
  };
  const key = /\d/.test(c) ? digitMap[c] : c;
  return unvoiced.has(key) ? "s" : "z";
}

// ─────────────────────────────────────────────────
// 1) Check S/Z: always resets & rescans on each click
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  // **always** clear any old notification + previously queued errors
  clearNotification(NOTIF_ID);
  state.errors = [];

  try {
    await Word.run(async context => {
      // A) Clear old highlights
      const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // B) Find standalone "s" and "z"
      const opts = { matchWholeWord: true, matchCase: false };
      const sRes = context.document.body.search("s", opts);
      const zRes = context.document.body.search("z", opts);
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // C) Filter to pure single‐letter, sort in document order
      let all = [...sRes.items, ...zRes.items]
        .filter(r => /^[sSzZ]$/.test(r.text.trim()))
        .sort((a, b) => a.compareLocationWith(b, Word.CompareLocation.start));

      // D) Evaluate each candidate
      for (const r of all) {
        const raw = r.text.trim();
        const actualLower = raw.toLowerCase();

        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();
        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === actualLower) continue;

        // preserve uppercase
        const suggestion = raw === raw.toUpperCase()
          ? expectedLower.toUpperCase()
          : expectedLower;

        context.trackedObjects.add(r);
        r.font.highlightColor = HIGHLIGHT_COLOR;
        state.errors.push({ range: r, suggestion });
      }

      await context.sync();

      // E) Notify or select the first mismatch
      if (!state.errors.length) {
        showNotification(NOTIF_ID, {
          type: "informationalMessage",
          message: "✨ No mismatches!",
          icon: "Icon.80x80"
        });
      } else {
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Check failed; please try again."
    });
  }
}

// ─────────────────────────────────────────────────
// 2) Accept One: replace first queued mismatch & select next
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (!state.errors.length) return;
  const { range, suggestion } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;

    if (state.errors.length) {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
    }
    await context.sync();
  });
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear first queued mismatch & select next
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (!state.errors.length) return;
  const { range } = state.errors.shift();

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;

    if (state.errors.length) {
      const next = state.errors[0].range;
      context.trackedObjects.add(next);
      next.select();
    }
    await context.sync();
  });
}

// ─────────────────────────────────────────────────
// 4) Accept All: fresh search & replace every mismatch
// ─────────────────────────────────────────────────
export async function acceptAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    let all = [...sRes.items, ...zRes.items]
      .filter(r => /^[sSzZ]$/.test(r.text.trim()))
      .sort((a, b) => a.compareLocationWith(b, Word.CompareLocation.start));

    for (const r of all) {
      const raw = r.text.trim();
      const actualLower = raw.toLowerCase();

      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();
      const nxt = after.text.trim();
      if (!nxt) continue;

      const expectedLower = determineCorrectPreposition(nxt);
      if (!expectedLower || expectedLower === actualLower) continue;

      const suggestion = raw === raw.toUpperCase()
        ? expectedLower.toUpperCase()
        : expectedLower;

      r.insertText(suggestion, Word.InsertLocation.replace);
      r.font.highlightColor = null;
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

// ─────────────────────────────────────────────────
// 5) Reject All: fresh search & clear every highlight
// ─────────────────────────────────────────────────
export async function rejectAllChanges() {
  clearNotification(NOTIF_ID);

  await Word.run(async context => {
    const opts = { matchWholeWord: true, matchCase: false };
    const sRes = context.document.body.search("s", opts);
    const zRes = context.document.body.search("z", opts);
    sRes.load("items"); zRes.load("items");
    await context.sync();

    let all = [...sRes.items, ...zRes.items]
      .filter(r => /^[sSzZ]$/.test(r.text.trim()))
      .sort((a, b) => a.compareLocationWith(b, Word.CompareLocation.start));

    for (const r of all) {
      r.font.highlightColor = null;
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

