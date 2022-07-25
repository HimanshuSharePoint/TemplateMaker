// JavaScript source code
var Winnable_ResponseLibrary_BAL = {
    initializePage: function () {
        alert("I am from initializePage Method");
	    renderPageControls();
    }
};

function renderPageControls() {
	alert("I am from Render Page Controls");
	kendoGrid_SnippetLib_Bind();
	hideSpinner("#spinner-ResponseLib-GridView");
	showControl("#divSnippetLib_GridView_Container");

	var ddlSnippetCategory = $("#ddlSnippetCategory").kendoDropDownList({
		optionLabel: "Select Category",
		dataSource: ["Prices", "Resumes", "Case Studies", "Human Resource Information", "Capacity", "Security Process", "Expertise", "Service Approach"]
	}).data("kendoDropDownList");

	$("#txtSnippetTitle").kendoTextBox({
		placeholder: "Define the content in a form of question"
	});

	$("#txtSnippetDescription").kendoEditor();

	var kendoWindow_CreateSnippet = $("#KendoWindow_CreateSnippet");
	kendoWindow_CreateSnippet.kendoWindow({
		width: "850px",
		height: "500px",
		actions: ["Close"],
		title: "Add New Snippet",
		close: function () {
			// enable button to stop further clicks btnCreateFreshTemplate.fadeIn();
		}
	});

	$("#btnAddSnippet").click(function () {
		kendoWindow_CreateSnippet.data("kendoWindow").center().open();
		// disable button to stop further clicks btnCreateFreshTemplate.fadeOut();
	});
	
}



function hideSpinner(controlID) {
	$(controlID).addClass("d-none");
	$(controlID).removeClass("d-flex");
}
function showSpinner(controlID) {
	$(controlID).removeClass("d-none");
	$(controlID).addClass("d-flex");
}

function showControl(controlID) {
	$(controlID).removeClass("d-none");
}
