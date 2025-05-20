/* global Office, Word */

let state = {
  errors: [],
  currentIndex: 0,
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
  const w = rawWord.normalize("NFC");
  const m = w.match(/[\p{L}0-9]/u);
  if (!m) return null;
  const first = m[0].toLowerCase();
  const unvoiced = new Set(['c','č','f','h','k','p','s','š','t']);
  const numMap = {'1':'e','2':'d','3':'t','4':'š','5':'p','6':'š','7':'s','8':'o','9':'d','0':'n'};
  if (/\d/.test(first)) {
    return unvoiced.has(numMap[first]) ? 's' : 'z';
  }
  return unvoiced.has(first) ? 's' : 'z';
}

export async function checkDocumentText() {
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);
  state.currentIndex = 0;

  try {
    await Word.run(async context => {
      // clear previous highlights
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];

      // search standalone 's' and 'z'
      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search('s', opts);
      const zRes = context.document.body.search('z', opts);
      sRes.load('items'); zRes.load('items');
      await context.sync();

      // filter exact 's' or 'z'
      const candidates = [...sRes.items, ...zRes.items].filter(r => {
        const t = r.text.trim().toLowerCase();
        return t === 's' || t === 'z';
      });

      // check each candidate
      for (const prep of candidates) {
        const after = prep.getRange('After');
        const nxtRange = after.getNextTextRange([' ', '\n', '.', ',', ';', '?', '!'], true);
        nxtRange.load('text');
        await context.sync();
        const nxt = nxtRange.text.trim();
        if (!nxt) continue;
        const actual = prep.text.trim().toLowerCase();
        const expect = determineCorrectPreposition(nxt);
        if (expect && actual !== expect) {
          context.trackedObjects.add(prep);
          state.errors.push({ range: prep, suggestion: expect });
        }
      }

      if (state.errors.length === 0) {
        showNotification(NOTIF_ID, { type: 'informationalMessage', message: "✨ No mismatches!", icon: 'Icon.80x80' });
      } else {
        // highlight all and select first
        state.errors.forEach(e => e.range.font.highlightColor = HIGHLIGHT_COLOR);
        await context.sync();
        const first = state.errors[0].range;
        context.trackedObjects.add(first);
        first.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error('checkDocumentText error', e);
    showNotification('checkError', { type: 'errorMessage', message: 'Check failed', persistent: false });
  } finally {
    state.isChecking = false;
  }
}

export async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  const err = state.errors[state.currentIndex];
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();
      // advance and select next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nextErr = state.errors[state.currentIndex].range;
        context.trackedObjects.add(nextErr);
        nextErr.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error('acceptCurrentChange error', e);
  }
}

export async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  const err = state.errors[state.currentIndex];
  try {
    await Word.run(async context => {
      context.trackedObjects.add(err.range);
      err.range.font.highlightColor = null;
      await context.sync();
      // advance and select next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nextErr = state.errors[state.currentIndex].range;
        context.trackedObjects.add(nextErr);
        nextErr.select();
        await context.sync();
      }
    });
  } catch (e) {
    console.error('rejectCurrentChange error', e);
  }
}

export async function acceptAllChanges() {
  if (state.errors.length === 0) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
        await context.sync();
      }
      state.currentIndex = state.errors.length;
    });
  } catch (e) {
    console.error('acceptAllChanges error', e);
  }
}

export async function rejectAllChanges() {
  if (state.errors.length === 0) return;
  try {
    await Word.run(async context => {
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.font.highlightColor = null;
        await context.sync();
      }
      state.currentIndex = state.errors.length;
    });
  } catch (e) {
    console.error('rejectAllChanges error', e);
  }
}
