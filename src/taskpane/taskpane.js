/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    window.addEventListener("DOMContentLoaded", () => {
      const wire = (id, fn) => {
        const el = document.getElementById(id);
        if (el) {
          el.addEventListener("click", fn);
        }
      };

      wire("checkTextButton", () => console.log("Clicked: Check Text"));
      wire("acceptChangeButton", () => console.log("Clicked: Accept Current Change"));
      wire("rejectChangeButton", () => console.log("Clicked: Reject Current Change"));
      wire("acceptAllButton", () => console.log("Clicked: Accept All Changes"));
      wire("rejectAllButton", () => console.log("Clicked: Reject All Changes"));

      console.log("✅ Taskpane UI is ready.");
    });
  }
});
