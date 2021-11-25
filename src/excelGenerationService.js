import ExcelJS from "exceljs";
import { substationKpis } from "./data";
import saveAs from "file-saver";
export const generateSubstaionKpiExcelFile  = ()=>{

    const columnHeaders = [
        { key: 'name', header: 'Name' },
        { key: 'quarter', header: 'Quarter' },
        { key: 'installedCapacity', header: 'Installed Capacity (MW)' },
        { key: 'peakLoad', header: 'Peak Load (MW)' },
        { key: 'totalCustomersServed', header: 'Customers Served' },
        { key: 'energyDelivered', header: 'Total Generation (MWh)' },
        { key: 'uptime', header: 'Total Operation Time (hrs)' },
        { key: 'totalCustomersInterrupted', header: 'Customers Interrupted' },
        { key: 'durationOfInterruption', header: 'Duration of Interruption' },
    ]

    const workbook = new ExcelJS.Workbook();    
    workbook.creator = 'Aaron Mineen';
    const worksheet = workbook.addWorksheet("Substation KPI");
    worksheet.columns= columnHeaders;
    substationKpis.forEach(element => {
        worksheet.addRow(element);
    });

    workbook.xlsx.writeBuffer().then(function(buffer) {
        saveAs(
          new Blob([buffer], { type: "application/octet-stream" }),
          `Substation Kpi.xlsx`
        );
      });

}