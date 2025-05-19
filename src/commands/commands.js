/* global document, Office */
import {
  checkDocumentText,
  acceptCurrentChange,
  rejectCurrentChange,
  acceptAllChanges,
  rejectAllChanges
} from "./preposition.js";

// — Log immediately when the commands bundle loads —
console.log("⭐ commands.js loaded");

Office.onReady(info => {
  console.log("▶️ Office.onReady", info);

  if (info.host === Office.HostType.Word) {
    console.log("🔗 Associating actions…");
    Office.actions.associate("checkDocumentText",   checkDocumentText);
    Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
    Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    Office.actions.associate("acceptAllChanges",    acceptAllChanges);
    Office.actions.associate("rejectAllChanges",    rejectAllChanges);
    console.log("✅ Actions associated");
  }
});
