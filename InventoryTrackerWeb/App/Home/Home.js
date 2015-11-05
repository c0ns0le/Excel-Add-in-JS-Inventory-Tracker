/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#create-table-and-highlight').click(createTableAndHighlight);
            $('#create-chart').click(createChart);

            //To Do 1: Check my host version
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
                return;
            }
        });
    };

    //To Do 2: Create a function to create a table and hightlight the ones with lowest inventory
    function createTableAndHighlight() {
        Excel.run(function (ctx) {

            //TODO 3: Queue requests to create table and set formula to calculate the minimum
            var inventoryTable = ctx.workbook.tables.add("Sheet1!D1:E5", true);
            inventoryTable.name = "InventoryTable";
            inventoryTable.showTotals = true;
            var totalRowRange = inventoryTable.getTotalRowRange();
            totalRowRange.getCell(0,0).values = [["Minimum"]];
            totalRowRange.getCell(0,1).formulas = [["=SUBTOTAL(105,[Inventory])"]];
            var dataRange = inventoryTable.getDataBodyRange();
            totalRowRange.load("values");
            dataRange.load("rowCount, values");
            return ctx.sync()

            //TODO 4: Add logic and scan through each row and highlight the minimum inventory
            .then(function () {
                for (var i = 0; i < dataRange.rowCount; i++) {
                    if (dataRange.values[i][1] == totalRowRange.values[0][1]) {
                        dataRange.getRow(i).format.fill.color = "yellow";
                    }
                }

            })
                .then(ctx.sync);
        })
        .catch(function (error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            app.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    //To Do 5: Create a function to create a Dashoboard sheet and then add a chart
    function createChart() {
        Excel.run(function (ctx) {
            // To Do 6: Create a worksheet
            var chartSheet = ctx.workbook.worksheets.add("Dashboard");
            chartSheet.activate();
            // To Do 7: Create a chart
            var dataRange = ctx.workbook.tables.getItem("InventoryTable").getDataBodyRange();
            var chart = chartSheet.charts.add(Excel.ChartType.columnClustered, dataRange);
            chart.title.text = "Inventory";
            chart.title.format.font.bold = true;
            return ctx.sync();

        })
        .catch(function (error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            app.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();