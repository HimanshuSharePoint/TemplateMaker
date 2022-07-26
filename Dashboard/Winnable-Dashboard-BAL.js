var Winnable_Dashboard_BAL = {
    initializePage: function () {
        renderPageControls();
    }
};

function renderPageControls() {
    try {
        console.log("Deugging");
        hideControl("#proposalContainer");

        $("#divDetailView").kendoWindow({
            actions: ["Close"],
            modal: true,
            visible: false,
            title: "View Details",
            close: this.onCloseKendoFade
        }).data("kendoWindow").maximize();

        $("#kendoCreateAppWindow").kendoWindow({
            actions: ["Close"],
            width: "1000px",
            height: "500px",
            modal: true,
            visible: false,
            title: "Create new Opportunity",
            close: this.onCloseKendo
        }).data("kendoWindow").center();

        $("#txtContractOwner").kendoTextBox({
            placeholder: "Please Enter Contract Owner"
        });
        $("#txtOpportunityType").kendoTextBox({
            placeholder: "Please Enter Opportunity Type"
        });
        $("#txtClientName").kendoTextBox({
            placeholder: "Enter Client Name"
        });
        $("#txtphone").kendoTextBox({
            placeholder: "Enter Phone"
        });
        $("#txtOpportunityName").kendoTextBox({
            placeholder: "Enter Opportunity Name"
        });
        $("#txtScopeofWork").kendoTextBox({
            placeholder: "Enter Scope"
        });
        $("#txtCountry").kendoTextBox({
            placeholder: "Enter Country"
        });
        $("#dateOpportunity").kendoDatePicker({

        });
        $("#txtEmail").kendoTextBox({
            placeholder: "Enter Email"
        });
        $("#txtStage").kendoTextBox({
            placeholder: "Enter Stage"
        })

        $(document).on("click", "#btnProposalView", function () {
            try {
                hideControl("#proposalDivShowhide");
                hideControl("#divContent_Wrapper");
                showControl("#divAllProposalsContent_Wrapper");
                showControl("#proposalContainer");
                kendoGrid_AllProposal_Bind();
                AllProposal_FullCalendarView_Bind();

                // Bind TabStript


                $("#proposalview").kendoTabStrip({
                    animation: {
                        open: {
                            effects: "fadeIn"
                        }
                    }
                });
            } catch (ex) {
                console.log("Error while btnProposalView" + ex.message);
            }
        });


        $(document).on("click", "#checkinsight", function () {
            try {
                $("#insights").toggle(this.checked);
            } catch (ex) {
                console.log("Error while Checking Insights" + ex.message);
            }
        });
        $(document).on("click", "#btnCardoptionone", function () {
            try {
                $("#dropdownCardone").toggleClass("show");
                $("#btnCardoptionone").toggleClass("show menu-dropdown");
            } catch (ex) {
                console.log("Error while Checking Option card" + ex.message);
            }
        });

        $(document).on("click", "#btnOption2", function () {
            try {
                $("#drpoption2").toggleClass("show");
                $("#btnOption2").toggleClass("show menu-dropdown");
            } catch (ex) {
                console.log("Error while Checking Option card" + ex.message);
            }
        });

        $(document).on("click", "#btnoption3", function () {
            try {
                $("#drpoption3").toggleClass("show");
                $("#btnoption3").toggleClass("show menu-dropdown");
            } catch (ex) {
                console.log("Error while Checking Option card" + ex.message);
            }
        });

        $(document).on("click", "#btnoption4", function () {
            try {
                $("#dropoption4").toggleClass("show");
                $("#btnoption4").toggleClass("show menu-dropdown");
            } catch (ex) {
                console.log("Error while Checking Option card" + ex.message);
            }
        });


        $(document).on("click", ".detalink", function () {
            try {
                $("#divDetailView").kendoWindow({
                    actions: ["Close"],
                    modal: true,
                    visible: false,
                    title: "View Details",
                    close: this.onCloseKendoFade
                }).data("kendoWindow").maximize().open();
                showControl("#divDetailView");
                // hideControl("#proposalDivShowhide");
                // hideControl("#divAllProposalsContent_Wrapper");
                // hideControl("#proposalview");

                $("#divStatusGrid_Container").kendoTabStrip({
                    animation: {
                        open: {
                            effects: "fadeIn"
                        }
                    }
                });
                bindGridView_StatusBySection();
                bindGridView_StatusByReview();
                createChart_CompletionStatus();
                createChart_CountdownTimer();

            } catch (ex) {
                console.log("Error while Clicking Details Button" + ex.message);
            }
        });

        $(document).on("click", "#btncreateNewOpportunity", function () {
            try {
                $("#kendoCreateAppWindow").kendoWindow({
                    actions: ["Close"],
                    width: "1000px",
                    height: "750px",
                    modal: true,
                    visible: false,
                    title: "Create new Opportunity",
                    close: this.onclosekendo
                }).data("kendoWindow").center().open();

            } catch (ex) {
                console.log("Error while click on btncreateNewOpportunity" + ex.message);
            }
        });
        // End Kendo conversion
        this.createChart_ProposalByStage();
        createChart_ProposalByOwner();
        createChart_ProposalByReview();
        createChart_ProposalByIndustry();
        createChart_ProposalByValue();
        hideSpinner("#spinner-TemplateLib-GridView");
    } catch (ex) {
        console.log("Error while renderPageControls" + ex.message);
    }
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
function hideControl(controlID) {
    $(controlID).addClass("d-none");
}
function selectControl(controlID) {
    $(controlID).addClass("active");
}
function deselectControl(controlID) {
    $(controlID).removeClass("active");
}
function triggerControl_Click(controlID) {
    $(controlID).trigger('click');
}






function kendoGrid_AllProposal_Bind() {

    var grid = $("#divProposal_GridView").data("kendoGrid");
    if (grid != null) grid.destroy();

    /* Load Data From Database & Bind DataSource Here*/

    $("#divProposal_GridView").kendoGrid({
        dataSource: {
            type: "odata",
            transport: {
                read: "https://demos.telerik.com/kendo-ui/service/Northwind.svc/Customers"
            },
            pageSize: 20
        },
        height: 550,
        groupable: true,
        sortable: true,
        pageable: {
            refresh: true,
            pageSizes: true,
            buttonCount: 5
        },
        columns: [{
            template:
                "<div class='customer-name'>#: ContactName #</div>",
            field: "ContactName",
            title: "Contact Name",
            width: 240
        }, {
            field: "ContactTitle",
            title: "Contact Title"
        }, {
            field: "CompanyName",
            title: "Company Name"
        }, {
            field: "Country",
            width: 150
        }],
        dataBound: function (e) {
            /* Fired when the widget is bound to data from its data source. Use this to show control & hide spinner */
            hideSpinner("#spinner-Proposal-GridView");
            showControl("#divProposal_GridView");
        }
    });

}
function AllProposal_FullCalendarView_Bind() {
    $("#divProposal_FullCalendarView").kendoScheduler({
        date: new Date("2013/6/13"),
        startTime: new Date("2013/6/13 07:00 AM"),
        //height: 72vh,
        views: [
            "day",
            { type: "workWeek", selected: true },
            "week",
            "month",
            "year",
            "agenda",
            { type: "timeline", eventHeight: 50 }
        ],
        timezone: "Etc/UTC",
        dataSource: {
            batch: true,
            transport: {
                read: {
                    url: "https://demos.telerik.com/kendo-ui/service/tasks",
                    dataType: "jsonp"
                },
                update: {
                    url: "https://demos.telerik.com/kendo-ui/service/tasks/update",
                    dataType: "jsonp"
                },
                create: {
                    url: "https://demos.telerik.com/kendo-ui/service/tasks/create",
                    dataType: "jsonp"
                },
                destroy: {
                    url: "https://demos.telerik.com/kendo-ui/service/tasks/destroy",
                    dataType: "jsonp"
                },
                parameterMap: function (options, operation) {
                    if (operation !== "read" && options.models) {
                        return { models: kendo.stringify(options.models) };
                    }
                }
            },
            schema: {
                model: {
                    id: "taskId",
                    fields: {
                        taskId: { from: "TaskID", type: "number" },
                        title: { from: "Title", defaultValue: "No title", validation: { required: true } },
                        start: { type: "date", from: "Start" },
                        end: { type: "date", from: "End" },
                        startTimezone: { from: "StartTimezone" },
                        endTimezone: { from: "EndTimezone" },
                        description: { from: "Description" },
                        recurrenceId: { from: "RecurrenceID" },
                        recurrenceRule: { from: "RecurrenceRule" },
                        recurrenceException: { from: "RecurrenceException" },
                        ownerId: { from: "OwnerID", defaultValue: 1 },
                        isAllDay: { type: "boolean", from: "IsAllDay" }
                    }
                }
            },
            filter: {
                logic: "or",
                filters: [
                    { field: "ownerId", operator: "eq", value: 1 },
                    { field: "ownerId", operator: "eq", value: 2 }
                ]
            }
        },
        resources: [
            {
                field: "ownerId",
                title: "Owner",
                dataSource: [
                    { text: "Alex", value: 1, color: "#f8a398" },
                    { text: "Bob", value: 2, color: "#2572c0" },
                    { text: "Charlie", value: 3, color: "#118640" }
                ]
            }
        ]
    });
}

function bindGridView_StatusBySection() {
    var grid = $("#grid_StatusByBySection").data("kendoGrid");
    if (grid != null) grid.destroy();

    $("#grid_StatusByBySection").kendoGrid({
        dataSource: {
            type: "odata",
            transport: {
                read: "https://demos.telerik.com/kendo-ui/service/Northwind.svc/Orders"
            },
            schema: {
                model: {
                    fields: {
                        OrderID: { type: "number" },
                        Freight: { type: "number" },
                        ShipName: { type: "string" },
                        OrderDate: { type: "date" },
                        ShipCity: { type: "string" }
                    }
                }
            },
            pageSize: 20
        },
        height: 550,
        groupable: true,
        sortable: true,
        pageable: {
            refresh: true,
            pageSizes: true,
            buttonCount: 5
        },
        columns: [{
            field: "OrderID",
            filterable: false
        },
            "Freight",
        {
            field: "OrderDate",
            title: "Order Date",
            format: "{0:MM/dd/yyyy}"
        }, {
            field: "ShipName",
            title: "Ship Name"
        }, {
            field: "ShipCity",
            title: "Ship City"
        }
        ],
        dataBound: function (e) {
            /* Fired when the widget is bound to data from its data source. Use this to show control & hide spinner */
            hideSpinner("#spinner-ResponseLib-GridView");
            showControl("#grid_StatusByBySection");
        }
    });
}
function bindGridView_StatusByReview() {
    var grid = $("#grid_StatusByReview").data("kendoGrid");
    if (grid != null) grid.destroy();

    $("#grid_StatusByReview").kendoGrid({
        dataSource: {
            type: "odata",
            transport: {
                read: "https://demos.telerik.com/kendo-ui/service/Northwind.svc/Orders"
            },
            schema: {
                model: {
                    fields: {
                        OrderID: { type: "number" },
                        Freight: { type: "number" },
                        ShipName: { type: "string" },
                        OrderDate: { type: "date" },
                        ShipCity: { type: "string" }
                    }
                }
            },
            pageSize: 20
        },
        height: 550,
        groupable: true,
        sortable: true,
        pageable: {
            refresh: true,
            pageSizes: true,
            buttonCount: 5
        },
        columns: [{
            field: "OrderID",
            filterable: false
        },
            "Freight",
        {
            field: "OrderDate",
            title: "Order Date",
            format: "{0:MM/dd/yyyy}"
        }, {
            field: "ShipName",
            title: "Ship Name"
        }, {
            field: "ShipCity",
            title: "Ship City"
        }
        ],
        dataBound: function (e) {
            /* Fired when the widget is bound to data from its data source. Use this to show control & hide spinner */
            hideSpinner("#spinner-ResponseLib-GridView");
            showControl("#grid_StatusByReview");
        }
    });
}


function createChart_ProposalByOwner() {
    try {
        var data = [
            {
                "source": "Maha",
                "percentage": 22,
                "explode": true
            },
            {
                "source": "Harsha",
                "percentage": 2
            },
            {
                "source": "Mohit",
                "percentage": 49
            },
            {
                "source": "Andrew",
                "percentage": 27
            }
        ];
        $("#chart_ProposalByOwner").kendoChart({
            title: {
                text: "Proposal by Owner"
            },
            legend: {
                position: "bottom"
            },
            dataSource: {
                data: data
            },
            series: [{
                type: "pie",
                field: "percentage",
                categoryField: "source",
                explodeField: "explode"
            }],
            seriesColors: ["#03a9f4", "#ff9800", "#fad84a", "#4caf50"],
            tooltip: {
                visible: true,
                template: "${category} - ${value}%"
            }
        });
    } catch (ex) {
        console.log("Error While createChart_ProposalByOwner" + ex.message);
    }
}
function createChart_ProposalByReview() {
    try {
        $("#chart_ProposalByReview").kendoChart({
            title: {
                text: "Proposal by Review"
            },
            legend: {
                position: "bottom"
            },
            seriesDefaults: {
                labels: {
                    template: "#= category # - #= kendo.format('{0:P}', percentage)#",
                    position: "outsideEnd",
                    visible: false,
                    background: "transparent"
                }
            },
            series: [{
                type: "donut",
                data: [{
                    category: "White",
                    value: 35
                }, {
                    category: "Pink",
                    value: 25
                }, {
                    category: "Golden",
                    value: 20
                }, {
                    category: "Green",
                    value: 10
                }, {
                    category: "No Review",
                    value: 10
                }]
            }],
            tooltip: {
                visible: true,
                template: "#= category # - #= kendo.format('{0:P}', percentage) #"
            }
        });
    } catch (ex) {
        console.log("Error while Bind Chart" + ex.message);
    }
}

function createChart_ProposalByIndustry() {
    $("#barChart_ProposalByIndustry").kendoChart({
        title: {
            text: "Proposal by Industry"
        },
        legend: {
            visible: false
        },
        seriesDefaults: {
            type: "bar"
        },
        series: [{
            name: "Total Proposal",
            data: [56000, 63000, 74000, 91000, 117000]
        }],
        valueAxis: {
            max: 140000,
            line: {
                visible: false
            },
            minorGridLines: {
                visible: true
            },
            labels: {
                rotation: "auto"
            }
        },
        categoryAxis: {
            categories: ["Dept. of Transportation", "Nestle", "Health Care", "P&G", "Tesla"],
            majorGridLines: {
                visible: false
            }
        },
        tooltip: {
            visible: true,
            template: "#= series.name #: #= value #"
        }
    });
}
function createChart_ProposalByValue() {
    try {
        $("#barChart_ProposalByValue").kendoChart({
            title: {
                text: "Proposal by Value"
            },
            legend: {
                visible: false
            },
            seriesDefaults: {
                type: "bar"
            },
            series: [{
                name: "Total Value",
                data: [10, 4, 7, 2, 12]
            }],
            valueAxis: {
                max: 12,
                line: {
                    visible: false
                },
                minorGridLines: {
                    visible: true
                },
                labels: {
                    rotation: "auto"
                }
            },
            categoryAxis: {
                categories: ["$2-5B", "$10-20M", "$5-10M", "$2-5M", "1-2M"],
                majorGridLines: {
                    visible: false
                }
            },
            tooltip: {
                visible: true,
                template: "#= series.name #: #= value #"
            }
        });
    } catch (ex) {
        console.log("Error while createChart_ProposalByValue" + ex.message);
    }
}

function createChart_CompletionStatus() {

    var pb = $("#divCompletionStatus_ProgressBar").kendoProgressBar({
        type: "percent",
        ariaRole: true,
        animation: {
            duration: 600
        }
    }).data("kendoProgressBar");
    pb.value(0);

    $("#divCompletionStatus_ProgressBar").data("kendoProgressBar").value(73);


}

function createChart_CountdownTimer() {
    // Set the date we're counting down to
    var countDownDate = new Date("November 29, 2022 15:37:25").getTime();

    // Update the count down every 1 second
    var x = setInterval(function () {

        // Get today's date and time
        var now = new Date().getTime();

        // Find the distance between now and the count down date
        var distance = countDownDate - now;

        // Time calculations for days, hours, minutes and seconds
        var days = Math.floor(distance / (1000 * 60 * 60 * 24));
        var hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
        var seconds = Math.floor((distance % (1000 * 60)) / 1000);

        // Output the result in an element with id="demo"
        document.getElementById("divCountdownTimer").innerHTML = "Time to submit: " + days + "d " + hours + "h "
            + minutes + "m " + seconds + "s ";

        // If the count down is over, write some text
        if (distance < 0) {
            clearInterval(x);
            document.getElementById("divCountdownTimer").innerHTML = "EXPIRED";
        }
    }, 1000);
}

function createChart_StatusBySection() {
    var progressbars = [];
    $(".poll-results div").each(function () {
        var pb = $(this).kendoProgressBar({
            type: "percent",
            ariaRole: true,
            animation: {
                duration: 600
            }
        }).data("kendoProgressBar");
        progressbars.push(pb);
    });

    $.each(progressbars, function (i, pb) {
        pb.value(0);
    });

    $("#fastAndFurious").data("kendoProgressBar").value(50);
    $("#theHelp").data("kendoProgressBar").value(30);
    $("#thePerks").data("kendoProgressBar").value(10);

    $.each(progressbars, function (i, pb) {
        if (pb.value() === 0) {
            pb.value(5);
        }
    });
}
function createChart_StatusByStakeholder() {
    $("#chart_StatusByStakeholder").kendoChart({
        title: {
            text: "World population by age group and sex"
        },
        legend: {
            visible: false
        },
        seriesDefaults: {
            type: "column",
            stack: {
                type: "100%"
            }
        },
        series: [{
            name: "70",
            stack: {
                group: "Female"
            },
            data: [854622, 925844, 984930, 1044982, 1100941, 1139797, 1172929, 1184435, 1184654]
        }, {
            name: "30",
            stack: {
                group: "Female"
            },
            data: [490550, 555695, 627763, 718568, 810169, 883051, 942151, 1001395, 1058439]
        }, {
            name: "45",
            stack: {
                group: "Female"
            },
            data: [379788, 411217, 447201, 484739, 395533, 435485, 499861, 569114, 655066]
        }, {
            name: "25",
            stack: {
                group: "Female"
            },
            data: [97894, 113287, 128808, 137459, 152171, 170262, 191015, 210767, 226956]
        }, {
            name: "55",
            stack: {
                group: "Female"
            },
            data: [16358, 18576, 24586, 30352, 36724, 42939, 46413, 54984, 66029]
        }],
        seriesColors: ["#cd1533", "#d43851", "#dc5c71", "#e47f8f", "#eba1ad",
            "#009bd7", "#26aadd", "#4db9e3", "#73c8e9", "#99d7ef"],
        valueAxis: {
            line: {
                visible: false
            }
        },
        categoryAxis: {
            categories: [1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010],
            majorGridLines: {
                visible: false
            }
        },
        tooltip: {
            visible: true,
            template: "#= series.stack.group #s, age #= series.name #"
        }
    });

    //PUT THIS CODE in TemplateLibrary_DAL.js
    //Put this function in DAL | Load data through API
    function loadData_DraftResponseLib() {

    }
    //Put this function in DAL | Load data through API
    function loadData_FinalResponseLib() {

    }


}

function createChart_ProposalByStage() {
    try {
        $("#chart_ProposalByStage").kendoChart({
            title: {
                text: "Proposal by stage"
            },
            legend: {
                position: "bottom"
            },
            seriesDefaults: {
                labels: {
                    template: "#= category # - #= kendo.format('{0:P}', percentage)#",
                    position: "outsideEnd",
                    visible: false,
                    background: "transparent"
                }
            },
            series: [{
                type: "donut",
                data: [{
                    category: "Won",
                    value: 35
                }, {
                    category: "Lost",
                    value: 25
                }, {
                    category: "Decline",
                    value: 20
                }, {
                    category: "Result awaited",
                    value: 10
                }, {
                    category: "Others",
                    value: 10
                }]
            }],
            tooltip: {
                visible: true,
                template: "#= category # - #= kendo.format('{0:P}', percentage) #"
            }
        });
    } catch (ex) {
        console.log("Error while createChart_ProposalByStage" + ex.message);
    }
}
function onCloseKendo() {
    try {
        $("#kendoCreateAppWindow").fadeIn();
    } catch (ex) {
        console.log("error while onCloseKendo" + ex.message);
    }
}

function onCloseKendoFade() {
    try {
        $("#divDetailView").fadeIn()
    } catch (ex) {
        console.log("Error while onCloseKendoFade" + ex.message);
    }
}
