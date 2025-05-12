/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Wait for the taskpane DOM to finish loading
    window.addEventListener("DOMContentLoaded", () => {
      const wire = (id, actionName) => {
        const el = document.getElementById(id);
        if (el) {
          el.addEventListener("click", () => {
            // Invoke the registered command
            Office.actions.executeFunction(actionName)
              .catch(err => console.error(`Error running ${actionName}:`, err));
          });
        }
      };

      wire("checkTextButton",       "checkDocumentText");
      wire("acceptChangeButton",    "acceptCurrentChange");
      wire("rejectChangeButton",    "rejectCurrentChange");
      wire("acceptAllButton",       "acceptAllChanges");
      wire("rejectAllButton",       "rejectAllChanges");

      console.log("✅ Taskpane UI is ready and wired to Office.actions.");
    });
  }
});
