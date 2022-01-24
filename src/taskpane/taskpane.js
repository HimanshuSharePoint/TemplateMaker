/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
let TextboxforTemplate;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of office supports all the Office.js APIs that are used in the tutorial
    if(!Office.context.requirements.isSetSupported('wordApi','1.3')){
      console.log("Sorry,The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }


    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
       

  }



});


function insertParagraph() {
  Word.run(function (context) {
    var docBody = context.document.body;
    var Valueoftext = document.getElementById("txtInput").value;
    var ValueofDatatType = document.getElementById("TemplateDataTypeSelect").value;
    docBody.insertText(" <<" + Valueoftext +" /" + ValueofDatatType + ">> ",Word.InsertLocation.end);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}



