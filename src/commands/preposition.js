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
  console.log('▶ checkDocumentText()', state);
  if (state.isChecking) return;
  state.isChecking = true;
  clearNotification(NOTIF_ID);

  try {
    await Word.run(async context => {
      // reset
      state.errors.forEach(e => e.range.font.highlightColor = null);
      state.errors = [];
      state.currentIndex = 0;

      const opts = { matchCase: false, matchWholeWord: true };
      const sRes = context.document.body.search('s', opts);
      const zRes = context.document.body.search('z', opts);
      sRes.load('items'); zRes.load('items');
      await context.sync();

      const all = [...sRes.items, ...zRes.items];
      const cand = all.filter(r => ['s','z'].includes(r.text.trim().toLowerCase()));
      console.log(`→ found ${cand.length} candidates`);

      for (const prep of cand) {
        const after = prep.getRange('After');
        const nxtR = after.getNextTextRange([' ', '\n', '.', ',', ';', '?', '!'], true);
        nxtR.load('text');
        await context.sync();
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
        first.select();
        await context.sync();
      }
    });
  } catch (err) {
    console.error('checkDocumentText error', err);
    showNotification('checkError', { type: 'errorMessage', message: 'Check failed', persistent: false });
  } finally { state.isChecking = false; }
}

// --- Accept one -- replaces current and moves on ---
export async function acceptCurrentChange() {
  console.log('▶ acceptCurrentChange()', state.currentIndex, state.errors.length);
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      console.log('→ acceptCurrentChange context start');
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);
      err.range.select(); await context.sync();

      // use selection proxy to replace
      const sel = context.document.getSelection();
      sel.load('text'); await context.sync();
      console.log(`   replacing '${sel.text}' → '${err.suggestion}'`);
      sel.insertText(err.suggestion, Word.InsertLocation.replace);
      sel.font.highlightColor = null;
      await context.sync();

      // next
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nxt = state.errors[state.currentIndex].range;
        context.trackedObjects.add(nxt);
        nxt.select(); await context.sync();
      }
    });
  } catch (err) { console.error('acceptCurrentChange error', err); }
}

// --- Reject one --- clears highlight and moves on ---
export async function rejectCurrentChange() {
  console.log('▶ rejectCurrentChange()', state.currentIndex, state.errors.length);
  if (state.currentIndex >= state.errors.length) return;

  try {
    await Word.run(async context => {
      console.log('→ rejectCurrentChange context start');
      const err = state.errors[state.currentIndex];
      context.trackedObjects.add(err.range);
      err.range.select(); await context.sync();
      const sel = context.document.getSelection(); sel.load('text'); await context.sync();
      console.log(`   clearing highlight for '${sel.text}'`);
      sel.font.highlightColor = null;
      await context.sync();
      state.currentIndex++;
      if (state.currentIndex < state.errors.length) {
        const nxt = state.errors[state.currentIndex].range;
        context.trackedObjects.add(nxt);
        nxt.select(); await context.sync();
      }
    });
  } catch (err) { console.error('rejectCurrentChange error', err); }
}

// --- Accept all --- iterates through all ---
export async function acceptAllChanges() {
  console.log('▶ acceptAllChanges()', state.errors.length);
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      console.log(`→ accepting all ${state.errors.length}`);
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.select(); await context.sync();
        const sel = context.document.getSelection(); sel.load('text'); await context.sync();
        console.log(`   replacing '${sel.text}' → '${err.suggestion}'`);
        sel.insertText(err.suggestion, Word.InsertLocation.replace);
        sel.font.highlightColor = null;
        await context.sync();
      }
      state.errors = [];
      state.currentIndex = 0;
      showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'Accepted all!', icon: 'Icon.80x80' });
    });
  } catch (err) { console.error('acceptAllChanges error', err); }
}

// --- Reject all --- clears all highlights ---
export async function rejectAllChanges() {
  console.log('▶ rejectAllChanges()', state.errors.length);
  if (!state.errors.length) return;

  try {
    await Word.run(async context => {
      console.log(`→ rejecting all ${state.errors.length}`);
      for (const err of state.errors) {
        context.trackedObjects.add(err.range);
        err.range.select(); await context.sync();
        const sel = context.document.getSelection(); sel.load('text'); await context.sync();
        console.log(`   clearing highlight for '${sel.text}'`);
        sel.font.highlightColor = null;
        await context.sync();
      }
      state.errors = [];
      state.currentIndex = 0;
      showNotification(NOTIF_ID, { type: 'informationalMessage', message: 'Cleared all!', icon: 'Icon.80x80' });
    });
  } catch (err) { console.error('rejectAllChanges error', err); }
}
