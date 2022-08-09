import { TableColumnType } from "antd";
import xlsx from "better-xlsx";
import objectPath from "object-path";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from 'jspdf-autotable';
import "./font";

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
      const cell = headerRow.addCell();
      cell.value = title;
    });
    data.forEach((record) => {
      const row = sheet.addRow();
      columns.forEach(({ dataIndex, render }) => {
        if (render) return;
        const cell = row.addCell();
        cell.value = objectPath.get(record, dataIndex as objectPath.Path);
      });
    });

    file.saveAs('blob').then((blob: Blob) => {
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

    csv += "\n";

    data.forEach((record) => {
      columns.forEach(({ dataIndex, render }, index) => {
        if (render) return;

        if (index !== 0) csv += ",";

        csv += `${objectPath.get(record, dataIndex as objectPath.Path).replaceAll('"', '""')}`;
      });
      csv += "\n";
    });

    saveAs(new Blob([csv]), `${fileName}.csv`);
  }

  const onPdfPrint = () => {
    const doc = new jsPDF();
    doc.setFont('FreeSans');

    autoTable(doc, {
      styles: { font: "FreeSans" },
      headStyles: { fontStyle: 'normal' },
      head: [columns.filter(c => !c.render).map(c => c.title)],
      body: data.map(r => columns.filter(c => !c.render).map(c => objectPath.get(r, c.dataIndex as objectPath.Path))),
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