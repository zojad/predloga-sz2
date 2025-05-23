/* global document, Office */
import {
  checkDocumentText,
  acceptAllChanges,
  rejectAllChanges
} from "./preposition.js";

console.log("⭐ commands.js loaded");

Office.onReady(info => {
  console.log("▶️ Office.onReady", info);

  if (info.host === Office.HostType.Word) {
    console.log("🔗 Associating actions…");

    const makeHandler = fn => async event => {
      console.log(`▶️ ${fn.name} invoked`);
      try {
        await fn();
      } catch (e) {
        console.error(`${fn.name} threw:`, e);
      } finally {
        event.completed();    // tell Word we’re done
      }
    };

    Office.actions.associate(
      "checkDocumentText",
      makeHandler(checkDocumentText)
    );
    Office.actions.associate(
      "acceptAllChanges",
      makeHandler(acceptAllChanges)
    );
    Office.actions.associate(
      "rejectAllChanges",
      makeHandler(rejectAllChanges)
    );

    console.log("✅ Actions associated");
  }
});
