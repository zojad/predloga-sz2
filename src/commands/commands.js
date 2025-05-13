/* global document, Office */
import {
  checkDocumentText,
  acceptCurrentChange,
  rejectCurrentChange,
  acceptAllChanges,
  rejectAllChanges
} from "./preposition.js";

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("checkDocumentText",   checkDocumentText);
    Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
    Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    Office.actions.associate("acceptAllChanges",    acceptAllChanges);
    Office.actions.associate("rejectAllChanges",    rejectAllChanges);
  }
});
