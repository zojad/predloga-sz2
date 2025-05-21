/* global document, Office */
import {
  checkDocumentText,
  acceptCurrentChange,
  rejectCurrentChange,
  acceptAllChanges,
  rejectAllChanges
} from "./preposition.js";

console.log("⭐ commands.js loaded");
Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // — Check S/Z —
    Office.actions.associate("checkDocumentText", async (event) => {
      console.log("▶️ checkDocumentText");
      await checkDocumentText();
      event.completed();               // ← tell Word we’re done
    });

    // — Accept One —
    Office.actions.associate("acceptCurrentChange", async (event) => {
      console.log("▶️ acceptCurrentChange");
      await acceptCurrentChange();
      event.completed();
    });

    // — Reject One —
    Office.actions.associate("rejectCurrentChange", async (event) => {
      console.log("▶️ rejectCurrentChange");
      await rejectCurrentChange();
      event.completed();
    });

    // — Accept All —
    Office.actions.associate("acceptAllChanges", async (event) => {
      console.log("▶️ acceptAllChanges");
      await acceptAllChanges();
      event.completed();
    });

    // — Reject All —
    Office.actions.associate("rejectAllChanges", async (event) => {
      console.log("▶️ rejectAllChanges");
      await rejectAllChanges();
      event.completed();
    });
  }
});
