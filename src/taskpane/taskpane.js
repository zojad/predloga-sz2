/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import {
  checkDocumentText,
  acceptCurrentChange,
  rejectCurrentChange,
  acceptAllChanges,
  rejectAllChanges
} from "../commands/preposition.js";

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("checkTextButton").onclick   = checkDocumentText;
    document.getElementById("acceptChangeButton").onclick = acceptCurrentChange;
    document.getElementById("rejectChangeButton").onclick = rejectCurrentChange;
    document.getElementById("acceptAllButton").onclick    = acceptAllChanges;
    document.getElementById("rejectAllButton").onclick    = rejectAllChanges;
  }
});
