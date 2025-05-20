/* global Office, Word */

let state = {
  errors: [],
  isChecking: false
};

const HIGHLIGHT_COLOR = "#FFC0CB";
const NOTIF_ID = "noErrors";

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
  const match = rawWord.normalize("NFC").match(/[\p{L}0-9]/u);
  if (!match) return null;
  const first = match[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap = {'1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n'};
  return (/[0-9]/.test(first)
    ? (unvoiced.has(numMap[first]) ? 's' : 'z')
    : (unvoiced.has(first) ? 's' : 'z')
  );
}

// --- CORE: find and highlight errors ---
export async function checkDocumentText() {
  console.log('▶ checkDocumentText');
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // clear old highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];

      // search standalone s/z
      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search('s', opts);
      const zRes = context.document.body.search('z', opts);
      sRes.load('items'); zRes.load('items');
      await context.sync();

      const all = [...sRes.items, ...zRes.items];
      for (const prep of all) {
        const after = prep.getRange('After');
        const nxtR = after.getNextTextRange([' ', '\n', '.', ',', ';', '?', '!'], true);
        nxtR.load('text'); await context.sync();
        const nxt = nxtR.text.trim();
        if (!nxt) continue;
        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          context.trackedObjects.add(prep);
          state.errors.push({ range: prep, suggestion: expect });
        }
      }

      console.log(`→ mismatches: ${state.errors.length}`);
      if (!state.errors.length) {
        showNotification(NOTIF_ID, { type: 'informationalMessage', message: "✨ No mismatches!", icon: 'Icon.80x80' });
      } else {
        state.errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
        await context.sync();
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select(); await context.sync();
      }
    });
  } catch (err) {
    console.error('checkDocumentText error', err);
    showNotification('checkError', { type: 'errorMessage', message: 'Check failed', persistent: false });
  } finally {
    state.isChecking = false;
  }
}

// --- Accept one --- process first error and select next ---
export async function acceptCurrentChange() {
  console.log('▶ acceptCurrentChange');
  if (!state.errors.length) return;

  const err = state.errors.shift();
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();
    });
  } catch (e) {
    console.error('acceptCurrentChange error', e);
  }

  // select next if available
  if (state.errors.length) {
    try {
      await Word.run(async context => {
        const next = state.errors[0].range;
        context.trackedObjects.add(next);
        next.select(); await context.sync();
      });
    } catch (e) {
      console.error('acceptCurrentChange select next error', e);
    }
  } else {
    showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'All changes accepted!', icon: 'Icon.80x80' });
  }
}

// --- Reject one --- clear first highlight and select next ---
export async function rejectCurrentChange() {
  console.log('▶ rejectCurrentChange');
  if (!state.errors.length) return;

  const err = state.errors.shift();
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.font.highlightColor = null;
      await context.sync();
    });
  } catch (e) {
    console.error('rejectCurrentChange error', e);
  }

  if (state.errors.length) {
    try {
      await Word.run(async context => {
        const next = state.errors[0].range;
        context.trackedObjects.add(next);
        next.select(); await context.sync();
      });
    } catch (e) {
      console.error('rejectCurrentChange select next error', e);
    }
  } else {
    showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'All highlights cleared!', icon: 'Icon.80x80' });
  }
}

// --- Accept all --- apply all suggestions ---
export async function acceptAllChanges() {
  console.log('▶ acceptAllChanges');
  if (!state.errors.length) return;

  const errorsToAccept = state.errors.slice();
  state.errors = [];

  try {
    await Word.run(async context => {
      for (const err of errorsToAccept) {
        context.trackedObjects.add(err.range);
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
        await context.sync();
      }
    });
    showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'Accepted all!', icon: 'Icon.80x80' });
  } catch (e) {
    console.error('acceptAllChanges error', e);
  }
}

// --- Reject all --- clear all highlights ---
export async function rejectAllChanges() {
  console.log('▶ rejectAllChanges');
  if (!state.errors.length) return;

  const errorsToClear = state.errors.slice();
  state.errors = [];

  try {
    await Word.run(async context => {
      for (const err of errorsToClear) {
        context.trackedObjects.add(err.range);
        err.range.font.highlightColor = null;
        await context.sync();
      }
    });
    showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'Cleared all!', icon: 'Icon.80x80' });
  } catch (e) {
    console.error('rejectAllChanges error', e);
  }
}
