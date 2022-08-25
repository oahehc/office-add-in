import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";

/* global console, Excel, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

// filterName, filterColumn, values
type Settings = [string, string, string];

export default function App({ title, isOfficeInitialized }: AppProps) {
  const [table, setTable] = React.useState("");
  const [settings, setSettings] = React.useState<Settings[]>([]);

  async function loadSettings() {
    await Excel.run(async (context) => {
      const settingSheet = context.workbook.worksheets.getItem("settings");
      // TODO: error handling if settings sheet doesn't exist
      const table = settingSheet.getRange("A1:C10");
      table.load("values");
      await context.sync();

      for (const [key, val1, val2] of table.values) {
        if (key === "table" && val1) {
          setTable(val1);
        } else if (key && val1 && val2) {
          setSettings((s) => [...s, [key, val1, val2]]);
        }
      }

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  React.useEffect(() => {
    loadSettings();
  }, []);

  async function clearFilter() {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      // TODO: error handling if table doesn't exist
      const targetTable = currentWorksheet.tables.getItem(table);

      for (const [_, column] of settings) {
        const filter = targetTable.columns.getItem(column).filter;
        filter.clear();
      }

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  }

  const applyFilter = (column, values) => async () => {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const targetTable = currentWorksheet.tables.getItem(table);
      // targetTable.load("filter");
      // await context.sync();
      // TODO: error handling if column doesn't exist
      const filter = targetTable.columns.getItem(column).filter;
      filter.applyValuesFilter(values.split(","));

      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  if (!table) {
    return <div>Table Not Exist</div>;
  }

  if (settings.length === 0) {
    return <div>Not Filters</div>;
  }

  return (
    <div className="ms-welcome p8">
      <div className="pb4">
        <DefaultButton iconProps={{ iconName: "" }} onClick={clearFilter}>
          Clear Filter
        </DefaultButton>
      </div>
      {settings.map(([filterName, column, values]) => (
        <div key={filterName} className="pb4">
          <DefaultButton iconProps={{ iconName: "" }} onClick={applyFilter(column, values)}>
            {filterName}
          </DefaultButton>
        </div>
      ))}
    </div>
  );
}
