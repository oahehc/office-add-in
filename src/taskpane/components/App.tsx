import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require, fetch  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

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
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
      <HeroList message="" items={[]}>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
          {isLoading ? "..." : "Run"}
        </DefaultButton>
      </HeroList>
    </div>
  );
}
