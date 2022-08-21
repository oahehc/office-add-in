import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require, fetch, Office, document  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

let dialog = null;

export default function App({ title, isOfficeInitialized }: AppProps) {
  const [isLoading, setLoading] = React.useState(false);
  async function click() {
    try {
      setLoading(true);
      const data = await fetch(
        "https://www.7timer.info/bin/astro.php?lon=113.2&lat=23.1&ac=0&unit=metric&output=json&tzshift=0"
      );
      const { dataseries } = await data.json();

      await Excel.run(async (context) => {
        const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();

        const range = context.workbook.getSelectedRange();
        range.load("address"); // Sheet1!J6
        range.format.fill.color = "yellow";
        await context.sync();
        activeWorksheet.getRange("A1:B2").set({
          numberFormat: [["0.00%"]],
          values: [
            [dataseries[0].temp2m, dataseries[1].temp2m],
            ["3", "4"],
          ],
          format: {
            fill: {
              color: "red",
            },
          },
        });
        await context.sync();

        // // didn't work?
        // const sheetName = "Sheet1";
        // const rangeAddress = "A1:B2";
        // const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
        // myRange.load("values");
        // await context.sync();
        // activeWorksheet.getRange("B1").set({
        //   numberFormat: [["0.00%"]],
        //   values: [[myRange.values]],
        //   format: {
        //     fill: {
        //       color: "red",
        //     },
        //   },
        // });
        // await context.sync();

        // const tableCount = context.workbook.tables.getCount();
        // await context.sync();
        // activeWorksheet.getRange("C1").set({
        //   numberFormat: [["0.00%"]],
        //   values: [[tableCount]], // empty?
        //   format: {
        //     fill: {
        //       color: "red",
        //     },
        //   },
        // });
        // await context.sync();

        // const sheet = context.workbook.worksheets.getItem("Sample");
        // const range2 = sheet.getRange("A2:E2");
        // range2.set({
        //   format: {
        //     fill: {
        //       color: "#4472C4",
        //     },
        //     font: {
        //       name: "Verdana",
        //       color: "white",
        //     },
        //   },
        // });
        // range2.format.autofitColumns();
        // await context.sync();

        let dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
        await context.sync();
        // weird linter error?
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add("Data");
        }
        dataSheet.position = 1;
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    } finally {
      setLoading(false);
    }
  }
  async function createTable() {
    await Excel.run(async (context) => {
      // table creation
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      // populate the table with data.
      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
      expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
      ]);

      // format the table.
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      await context.sync();
    }).catch((error) => {
      console.error(error);
    });
  }

  async function filterTable() {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const categoryFilter = expensesTable.columns.getItem("Category").filter;
      categoryFilter.applyValuesFilter(["Education", "Groceries"]);

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  async function sortTable() {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const sortFields = [
        {
          key: 1, // Merchant column
          ascending: false,
        },
      ];

      expensesTable.sort.apply(sortFields);

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  async function freezeHeader() {
    await Excel.run(async (context) => {
      // keep the header visible when the user scrolls.
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.freezePanes.freezeRows(1);

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  async function createChart() {
    await Excel.run(async (context) => {
      // get the range of data to be charted
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const dataRange = expensesTable.getDataBodyRange();

      // to create the chart and define its type
      const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");

      // position and format the chart.
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = "Value in \u20AC";

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  function openDialog() {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/popup.html",
      { height: 45, width: 55 },
      function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  }
  function processMessage(arg) {
    document.getElementById("user-name").innerHTML = arg.message;
    dialog.close();
  }

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome p8">
      <div>
        名前：<label id="user-name"></label>
      </div>
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        {isLoading ? "..." : "Run"}
      </DefaultButton>
      <div>
        <DefaultButton iconProps={{ iconName: "" }} onClick={createTable}>
          Create Table
        </DefaultButton>
        <DefaultButton iconProps={{ iconName: "" }} onClick={filterTable}>
          Filter Table
        </DefaultButton>
        <DefaultButton iconProps={{ iconName: "" }} onClick={sortTable}>
          Sort Table
        </DefaultButton>
        <DefaultButton iconProps={{ iconName: "" }} onClick={freezeHeader}>
          Freeze Header
        </DefaultButton>
      </div>
      <div>
        <DefaultButton iconProps={{ iconName: "" }} onClick={createChart}>
          Create Chart
        </DefaultButton>
      </div>
      <div>
        <DefaultButton iconProps={{ iconName: "" }} onClick={openDialog}>
          Open Dialog
        </DefaultButton>
      </div>
    </div>
  );
}
