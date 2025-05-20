// Global state for tracking mismatches
const state = {
  errors: [],
  currentIndex: 0,
  isChecking: false
};

function checkDocumentText() {
  console.log('checkDocumentText()', state);
  if (state.isChecking) return;
  state.isChecking = true;
  state.errors = [];
  state.currentIndex = 0;
  Word.run(async context => {
    console.log('Word.run(start)');
    const body = context.document.body;
    // Find all instances of "s" or "z" (including uppercase) as whole words
    const sResults = body.search("s", { matchWholeWord: true });
    const zResults = body.search("z", { matchWholeWord: true });
    const SResults = body.search("S", { matchWholeWord: true });
    const ZResults = body.search("Z", { matchWholeWord: true });
    context.load(sResults, 'items');
    context.load(zResults, 'items');
    context.load(SResults, 'items');
    context.load(ZResults, 'items');
    await context.sync();
    const candidates = sResults.items.concat(zResults.items, SResults.items, ZResults.items);
    console.log(`found ${candidates.length} s/z candidates`);
    // Set of voiceless consonants (if next word starts with these, the correct preposition is "s")
    const voiceless = new Set(['c','č','f','h','k','p','s','š','t','x',
                                'C','Č','F','H','K','P','S','Š','T','X']);
    // Prepare to get the text of the next word after each candidate
    const nextRanges = [];
    for (let cand of candidates) {
      // Get the next text range after the preposition, up to the next punctuation/space
      let nextRange = cand.getNextTextRange([",", ".", ":", ";", "!", "?", " ", "\t", "\r", "\n"], true);
      context.load(nextRange, 'text');
      nextRanges.push(nextRange);
    }
    await context.sync();
    // Evaluate each candidate and mark mismatches
    for (let i = 0; i < candidates.length; i++) {
      const prepositionRange = candidates[i];
      const nextText = nextRanges[i].text;
      if (!nextText || nextText.length === 0) {
        // No following word (end of paragraph or document)
        continue;
      }
      const firstChar = nextText[0];
      // Determine what the correct preposition should be ("z" if next word starts with voiced or vowel, "s" if voiceless)
      let shouldUseZ = true;
      if (voiceless.has(firstChar)) {
        shouldUseZ = false;
      }
      const currentPreposition = prepositionRange.text;
      if (!currentPreposition) continue;
      const currentLower = currentPreposition.toLowerCase();
      const correctPreposition = shouldUseZ ? 'z' : 's';
      if (currentLower !== correctPreposition) {
        // Mismatch found – highlight it and add to errors list
        prepositionRange.font.highlightColor = 'Yellow';
        prepositionRange.track();  // keep the range for later use
        state.errors.push(prepositionRange);
      }
    }
    console.log(`mismatches found: ${state.errors.length}`);
  })
  .catch(err => {
    console.error(err);
  })
  .finally(() => {
    state.isChecking = false;
  });
}

function acceptCurrentChange() {
  console.log('acceptCurrentChange()', { currentIndex: state.currentIndex, errors: state.errors.length });
  if (!state.errors || state.errors.length === 0) {
    return;
  }
  Word.run(async context => {
    console.log('inside acceptCurrentChange Word.run');
    let errorRange = state.errors[state.currentIndex];
    // Load the current text of the range (should be the wrong preposition)
    context.load(errorRange, 'text');
    await context.sync();
    let wrongChar = errorRange.text || '';
    // Determine the correct character, preserving case
    let correctChar;
    if (wrongChar === 's') correctChar = 'z';
    else if (wrongChar === 'S') correctChar = 'Z';
    else if (wrongChar === 'z') correctChar = 's';
    else if (wrongChar === 'Z') correctChar = 'S';
    else {
      console.warn('Unexpected preposition character:', wrongChar);
      return;
    }
    console.log(`Replacing '${wrongChar}' -> '${correctChar}'`);
    // Replace the wrong preposition with the correct one and remove highlight
    errorRange.insertText(correctChar, Word.InsertLocation.replace);
    errorRange.font.highlightColor = null;
    // Select the next mismatch in the document (if any)
    if (state.currentIndex + 1 < state.errors.length) {
      const nextRange = state.errors[state.currentIndex + 1];
      nextRange.select();
    }
    // Stop tracking the fixed range
    context.trackedObjects.remove(errorRange);
    await context.sync();
  })
  .then(() => {
    // Remove the resolved error from the list and update index
    state.errors.splice(state.currentIndex, 1);
    if (state.currentIndex >= state.errors.length) {
      // Reset index if we reached the end of the list
      state.currentIndex = 0;
    }
  })
  .catch(error => {
    console.error(error);
  });
}

function rejectCurrentChange() {
  console.log('rejectCurrentChange()', { currentIndex: state.currentIndex, errors: state.errors.length });
  if (!state.errors || state.errors.length === 0) {
    return;
  }
  Word.run(async context => {
    console.log('inside rejectCurrentChange Word.run');
    let errorRange = state.errors[state.currentIndex];
    // Keep the text as is, just remove the highlight
    errorRange.font.highlightColor = null;
    // Select the next mismatch (if any)
    if (state.currentIndex + 1 < state.errors.length) {
      const nextRange = state.errors[state.currentIndex + 1];
      nextRange.select();
    }
    // Stop tracking this (ignored) range
    context.trackedObjects.remove(errorRange);
    await context.sync();
  })
  .then(() => {
    // Remove the skipped error from the list and update index
    state.errors.splice(state.currentIndex, 1);
    if (state.currentIndex >= state.errors.length) {
      state.currentIndex = 0;
    }
  })
  .catch(error => {
    console.error(error);
  });
}

function acceptAllChanges() {
  console.log('acceptAllChanges()', { total: state.errors.length });
  if (!state.errors || state.errors.length === 0) {
    return;
  }
  Word.run(async context => {
    console.log('inside acceptAllChanges Word.run');
    // Load text for all error ranges to determine their current letters
    for (let errorRange of state.errors) {
      context.load(errorRange, 'text');
    }
    await context.sync();
    // Replace all mismatched prepositions at once
    for (let errorRange of state.errors) {
      const wrongChar = errorRange.text || '';
      let correctChar;
      if (wrongChar === 's') correctChar = 'z';
      else if (wrongChar === 'S') correctChar = 'Z';
      else if (wrongChar === 'z') correctChar = 's';
      else if (wrongChar === 'Z') correctChar = 'S';
      else continue;  // skip if unexpected value
      console.log(`Replacing '${wrongChar}' -> '${correctChar}'`);
      errorRange.insertText(correctChar, Word.InsertLocation.replace);
      errorRange.font.highlightColor = null;
      // Stop tracking this range after replacement
      context.trackedObjects.remove(errorRange);
    }
    await context.sync();
  })
  .then(() => {
    // All errors fixed; clear the list and reset index
    state.errors = [];
    state.currentIndex = 0;
  })
  .catch(error => {
    console.error(error);
  });
}

function rejectAllChanges() {
  console.log('rejectAllChanges()', { total: state.errors.length });
  if (!state.errors || state.errors.length === 0) {
    return;
  }
  Word.run(async context => {
    console.log('inside rejectAllChanges Word.run');
    // Remove highlights from all marked prepositions without changing text
    for (let errorRange of state.errors) {
      errorRange.font.highlightColor = null;
      context.trackedObjects.remove(errorRange);
    }
    await context.sync();
  })
  .then(() => {
    // Clear all errors and reset index since we've ignored all
    state.errors = [];
    state.currentIndex = 0;
  })
  .catch(error => {
    console.error(error);
  });
}
