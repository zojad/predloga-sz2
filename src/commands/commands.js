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

    Office.actions.associate("checkDocumentText",   (...args) => {
      console.log("▶️ OfficeAction invoked: checkDocumentText", args);
      return checkDocumentText(...args);
    });

    Office.actions.associate("acceptCurrentChange", (...args) => {
      console.log("▶️ OfficeAction invoked: acceptCurrentChange", args);
      return acceptCurrentChange(...args);
    });

    Office.actions.associate("rejectCurrentChange", (...args) => {
      console.log("▶️ OfficeAction invoked: rejectCurrentChange", args);
      return rejectCurrentChange(...args);
    });

    Office.actions.associate("acceptAllChanges",    (...args) => {
      console.log("▶️ OfficeAction invoked: acceptAllChanges", args);
      return acceptAllChanges(...args);
    });

    Office.actions.associate("rejectAllChanges",    (...args) => {
      console.log("▶️ OfficeAction invoked: rejectAllChanges", args);
      return rejectAllChanges(...args);
    });

    console.log("✅ Actions associated");
  }
});
