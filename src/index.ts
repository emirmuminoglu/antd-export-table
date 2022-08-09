import { TableColumnType } from "antd";
import xlsx from "better-xlsx";
import objectPath from "object-path";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from 'jspdf-autotable';

const useExport = <RecordType extends object>({
  columns,
  data,
  fileName,
  pdfTheme
}: {
  columns: TableColumnType<RecordType>[];
  data: RecordType[];
  fileName: string;
  pdfTheme?: "striped" | "grid" | "plain";
}) => {
  const onExcelPrint = () => {
    const file = new xlsx.File();
    const sheet = file.addSheet("Sheet1");
    const headerRow = sheet.addRow();
    columns.forEach(({ title, render }) => {
      if (render) return;
      headerRow.addCell(title);
    });
    data.forEach((record) => {
      const row = sheet.addRow();
      columns.forEach(({ dataIndex }) => {
        row.addCell(objectPath.get(record, dataIndex as objectPath.Path));
      });
    });

    file.saveAs('blob').then(blob => {
      saveAs(blob, `${fileName}.xlsx`);
    })
  };

  const onCsvPrint = () => {
    let csv = "";
    columns.forEach(({ title, render }, index) => {
      if (render) return;
      if (index !== 0) csv += ",";

      csv += `${title.replaceAll('"', '""')}`;
    });

    data.forEach((record) => {
      columns.forEach(({ dataIndex }, index) => {
        if (index !== 0) csv += ",";

        csv += `${objectPath.get(record, dataIndex as objectPath.Path).replaceAll('"', '""')}`;
      });
    });

    saveAs(csv, `${fileName}.csv`);
  }

  const onPdfPrint = () => {
    const doc = new jsPDF();

    autoTable(doc, {
      head: columns.map(c => c.title),
      body: data.map(r => columns.map(c => objectPath.get(r, c.dataIndex as objectPath.Path))),
      theme: pdfTheme,
    })

    doc.save(`${fileName}.pdf`);
  }

  return {
    onExcelPrint,
    onCsvPrint,
    onPdfPrint,
  }
};

export default useExport;