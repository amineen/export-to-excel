import ExcelJS from "exceljs";
import { substationKpis } from "./data";
import saveAs from "file-saver";

export const generateSubstaionKpiExcelFile = () => {

    const columnHeaders = [
        { key: 'name', heading: 'Name', width: "22" },
        { key: 'quarter', heading: 'Quarter', width: "12" },
        { key: 'installedCapacity', heading: 'Installed Capacity (MW)', width: "23" },
        { key: 'peakLoad', heading: 'Peak Load (MW)', width: "16" },
        { key: 'totalCustomersServed', heading: 'Customers Served', width: "22" },
        { key: 'energyDelivered', heading: 'Total Generation (MWh)', width: "24" },
        { key: 'uptime', heading: 'Total Operation Time (hrs)', width: "26" },
        { key: 'totalCustomersInterrupted', heading: 'Customers Interrupted', width: "23" },
        { key: 'durationOfInterruption', heading: 'Duration of Interruption', width: "23" },
    ]
    const totalUptime = substationKpis.reduce((n, { uptime }) => n + uptime, 0);
    const totalEnergyDelivered = substationKpis.reduce((n, { energyDelivered }) => n + energyDelivered, 0);
    const totalInterruptionDuration = substationKpis.reduce((n, { durationOfInterruption }) => n + durationOfInterruption, 0);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Aaron Mineen';
    const worksheet = workbook.addWorksheet("Substation KPI");
    worksheet.columns = columnHeaders;
    worksheet.mergeCells("A1:I1");
    worksheet.getCell("A1").value = "Substation KPI Report";
    worksheet.getCell("A1").font = { bold: true, size: 16 }
    worksheet.getRow(2).values = columnHeaders.map(ch => ch.heading);
    worksheet.getRow(2).font = { bold: true };
    // worksheet.getCell("A2:I2").fill = {bgColor:"#96C7C1"}
    substationKpis.forEach(element => {
        worksheet.addRow(element);
    });
    const lastRow = worksheet.lastRow.number;
    worksheet.getCell(`A${lastRow + 1}`).value = "TOTAL";
    worksheet.getCell(`F${lastRow + 1}`).value = totalEnergyDelivered;
    worksheet.getCell(`G${lastRow + 1}`).value = totalUptime;
    worksheet.getCell(`I${lastRow + 1}`).value = totalInterruptionDuration;
    worksheet.getRow(lastRow + 1).font = { bold: true };
    workbook.xlsx.writeBuffer().then(function (buffer) {
        saveAs(
            new Blob([buffer], { type: "application/octet-stream" }),
            `Substation Kpi.xlsx`
        );
    });

}