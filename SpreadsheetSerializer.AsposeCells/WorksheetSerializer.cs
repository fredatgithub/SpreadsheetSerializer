﻿using System.Data;

namespace SpreadsheetSerializer.AsposeCells
{
    internal class WorksheetSerializer
    {
        private string WorksheetName { get; set; }
        private readonly DataTable dataTable;

        public WorksheetSerializer(DataTable dataTable, string worksheetName)
        {
            this.dataTable = dataTable;
            WorksheetName = worksheetName;
        }

        public Aspose.Cells.Worksheet CreateSheet(Aspose.Cells.Workbook workbook)
        {
            var sheet = workbook.Worksheets.Add(WorksheetName);
            var dataOptions = new Aspose.Cells.ImportTableOptions();

            sheet.Cells.ImportData(dataTable, 0, 0, dataOptions);
            sheet.AutoFitColumns();

            return sheet;
        }

        public void DisposeDataTable()
        {
            dataTable.Dispose();
        }
    }
}
