var Winnable_PlannerProposal_BAL = {
   initializePage: function () {
      alert("This is from External Planner Proposal");
      renderPageControls();
   }
};

function renderPageControls() {

   $(document).on("click", "#btncreateNewOpportunity", function () {
      try {
         $("#kendoCreateProposal").kendoWindow({
            actions: ["Close"],
            modal: true,
            visible: false,
            title: "Create Task",
            close: this.onCloseKendoFade
         }).data("kendoWindow").center().open();

      } catch (ex) {
         console.log("Error while btnProposalView" + ex.message);
      }
   });

   $("#kendoCreateProposal").kendoWindow({
      actions: ["Close"],
      modal: true,
      visible: false,
      title: "Create Task",
      close: this.onCloseKendoFade
   }).data("kendoWindow").center();
}

function onCloseKendoFade() {
   try {
      $("#kendoCreateProposal").fadeIn()
   } catch (ex) {
      console.log("Error while onCloseKendoFade" + ex.message);
   }
}
