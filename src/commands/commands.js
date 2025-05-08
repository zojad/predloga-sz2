/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log(" commands.js loaded — Office is ready!");

    // Register dummy functions to test wiring
    Office.actions.associate("checkDocumentText", () => {
      console.log(" checkDocumentText triggered from ribbon!");
      return true;
    });

    Office.actions.associate("acceptCurrentChange", () => {
      console.log(" acceptCurrentChange triggered from ribbon!");
      return true;
    });

    Office.actions.associate("rejectCurrentChange", () => {
      console.log(" rejectCurrentChange triggered from ribbon!");
      return true;
    });

    Office.actions.associate("acceptAllChanges", () => {
      console.log(" acceptAllChanges triggered from ribbon!");
      return true;
    });

    Office.actions.associate("rejectAllChanges", () => {
      console.log(" rejectAllChanges triggered from ribbon!");
      return true;
    });
  }
});
