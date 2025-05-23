/* global document, Office */

import {
  checkDocumentText,
  acceptAllChanges,
  rejectAllChanges
} from "../commands/preposition.js";

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // “Check S/Z”
    document.getElementById("checkTextButton").onclick = async () => {
      await checkDocumentText();
    };
    // “Accept All”
    document.getElementById("acceptAllButton").onclick = async () => {
      await acceptAllChanges();
    };
    // “Reject All”
    document.getElementById("rejectAllButton").onclick = async () => {
      await rejectAllChanges();
    };
  }
});
