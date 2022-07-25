// JavaScript source code
var Winnable_ResponseLibrary_BAL = {
    initializePage: function () {
        alert("I am from initializePage Method");
	    renderPageControls();
    }
};

function renderPageControls() {
	alert("I am from Render Page Controls");
	
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
function bindResponseLibrary() {

    var grid = $("#divSnippetLib_GridView_Container").data("kendoGrid");
    if (grid != null) grid.destroy();

    $("#divSnippetLib_GridView_Container").kendoGrid({
        dataSource: {
            data: dataSourceResponseLibrary.data(),
            pageSize: 10
        },
        toolbar: ["search"],
        scrollable: true,
        resizable: true,
        sortable: true,
        reorderable: true,
        noRecords: true,
        columnMenu: false,
        groupable: true,
        pageable: {
            refresh: true,
            pageSizes: true,
            buttonCount: 5,
        },
        filterable: true,
        columns: [
            {
                field: "SnippetCategory",
                title: "Snippet Category",

            }, {
                field: "SnippetTitle",
                title: "Snippet Title",
            }, {
                field: "SnippetDescription",
                title: "Snippet Description",
            }
        ],
        dataBound: function (e) {
            /* Fired when the widget is bound to data from its data source. Use this to show control & hide spinner */

        }
    });
}
