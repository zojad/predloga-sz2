/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    window.addEventListener("DOMContentLoaded", () => {
      const wire = (id, fnName) => {
        const el = document.getElementById(id);
        if (el && typeof window[fnName] === "function") {
          el.addEventListener("click", window[fnName]);
        }
      };

      wire("checkTextButton", "checkDocumentText");
      wire("acceptChangeButton", "acceptCurrentChange");
      wire("rejectChangeButton", "rejectCurrentChange");
      wire("acceptAllButton", "acceptAllChanges");
      wire("rejectAllButton", "rejectAllChanges");

      console.log("Taskpane UI is ready.");
    });
  }
});


