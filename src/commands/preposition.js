/* global Office, Word */

let state = {
  errors: [],        // Array<{ range: Word.Range, suggestion: "s"|"S"|"z"|"Z" }>
  currentIndex: 0,
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID        = "noErrors";

// Ribbon notification helpers
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
 * For the next word, return 's' if unvoiced (c,č,f,h,k,p,s,š,t),
 * otherwise 'z'.  (Digits map via digitMap.)
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
// 1) Check S/Z: highlight all mismatches & select first
// ─────────────────────────────────────────────────
export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  // Reset queue
  state.errors = [];
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // 1) Clear old highlights
      const oldS = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const oldZ = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      oldS.load("items"); oldZ.load("items");
      await context.sync();
      [...oldS.items, ...oldZ.items].forEach(r => r.font.highlightColor = null);
      await context.sync();

      // 2) Search for standalone s / z
      const sRes = context.document.body.search("s", { matchWholeWord: true, matchCase: false });
      const zRes = context.document.body.search("z", { matchWholeWord: true, matchCase: false });
      sRes.load("items"); zRes.load("items");
      await context.sync();

      // 3) Filter to pure single-letter and compute suggestion
      for (const r of [...sRes.items, ...zRes.items]) {
        const raw = r.text.trim();
        if (!/^[sSzZ]$/.test(raw)) continue;       // only exact "s" or "z"

        const actualLower = raw.toLowerCase();
        const after = r.getRange("After")
                       .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
        after.load("text");
        await context.sync();

        const nxt = after.text.trim();
        if (!nxt) continue;

        const expectedLower = determineCorrectPreposition(nxt);
        if (!expectedLower || expectedLower === actualLower) continue;

        // preserve case
        const suggestion = (raw === raw.toUpperCase())
          ? expectedLower.toUpperCase()
          : expectedLower;

        // *** track this range so it survives into the Accept/Reject handlers! ***
        context.trackedObjects.add(r);

        // highlight & queue
        r.font.highlightColor = HIGHLIGHT_COLOR;
        state.errors.push({ range: r, suggestion });
      }

      await context.sync();

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
  } catch (e) {
    console.error("checkDocumentText error", e);
    showNotification(NOTIF_ID, {
      type: "errorMessage",
      message: "Check failed; please try again."
    });
  } finally {
    state.isChecking = false;
  }
}

// ─────────────────────────────────────────────────
// 2) Accept One: replace current & auto-advance
// ─────────────────────────────────────────────────
export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  // splice out the queued mismatch
  const { range, suggestion } = state.errors.splice(state.currentIndex, 1)[0];

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.insertText(suggestion, Word.InsertLocation.replace);
    range.font.highlightColor = null;
    await context.sync();
  });

  // select the next mismatch, if any
  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 3) Reject One: clear current & auto-advance
// ─────────────────────────────────────────────────
export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;

  const { range } = state.errors.splice(state.currentIndex, 1)[0];

  await Word.run(async context => {
    context.trackedObjects.add(range);
    range.font.highlightColor = null;
    await context.sync();
  });

  if (state.currentIndex < state.errors.length) {
    await Word.run(async context => {
      const next = state.errors[state.currentIndex].range;
      context.trackedObjects.add(next);
      next.select();
      await context.sync();
    });
  }
}

// ─────────────────────────────────────────────────
// 4) Accept All: fresh search + replace every mismatch
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
      const raw = r.text.trim();
      if (!/^[sSzZ]$/.test(raw)) continue;
      const actualLower = raw.toLowerCase();

      const after = r.getRange("After")
                     .getNextTextRange([" ", "\n", ".", ",", ";", "?", "!"], true);
      after.load("text");
      await context.sync();

      const nxt = after.text.trim();
      if (!nxt) continue;

      const expectedLower = determineCorrectPreposition(nxt);
      if (!expectedLower || expectedLower === actualLower) continue;

      const suggestion = (raw === raw.toUpperCase())
        ? expectedLower.toUpperCase()
        : expectedLower;

      r.insertText(suggestion, Word.InsertLocation.replace);
      r.font.highlightColor = null;
    }

    await context.sync();
  });

  // reset queue
  state.errors = [];
  state.currentIndex = 0;
  showNotification(NOTIF_ID, {
    type: "informationalMessage",
    message: "Accepted all!",
    icon: "Icon.80x80"
  });
}

// ─────────────────────────────────────────────────
// 5) Reject All: fresh search + clear every highlight
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
      const raw = r.text.trim();
      if (!/^[sSzZ]$/.test(raw)) continue;
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
}

