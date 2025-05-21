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

    Office.actions.associate("checkDocumentText", async (event) => {
      console.log("▶️ OfficeAction invoked: checkDocumentText");
      try {
        await checkDocumentText();
      } catch (e) {
        console.error("checkDocumentText threw", e);
      } finally {
        event.completed();
      }
    });

    Office.actions.associate("acceptCurrentChange", async (event) => {
      console.log("▶️ OfficeAction invoked: acceptCurrentChange");
      try {
        await acceptCurrentChange();
      } catch (e) {
        console.error("acceptCurrentChange threw", e);
      } finally {
        event.completed();
      }
    });

    Office.actions.associate("rejectCurrentChange", async (event) => {
      console.log("▶️ OfficeAction invoked: rejectCurrentChange");
      try {
        await rejectCurrentChange();
      } catch (e) {
        console.error("rejectCurrentChange threw", e);
      } finally {
        event.completed();
      }
    });

    Office.actions.associate("acceptAllChanges", async (event) => {
      console.log("▶️ OfficeAction invoked: acceptAllChanges");
      try {
        await acceptAllChanges();
      } catch (e) {
        console.error("acceptAllChanges threw", e);
      } finally {
        event.completed();
      }
    });

    Office.actions.associate("rejectAllChanges", async (event) => {
      console.log("▶️ OfficeAction invoked: rejectAllChanges");
      try {
        await rejectAllChanges();
      } catch (e) {
        console.error("rejectAllChanges threw", e);
      } finally {
        event.completed();
      }
    });

    console.log("✅ Actions associated");
  }
});
