﻿using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookRetriever
    {
        public static AsposeWorkbook GetWorkbookFromFilePath(string filePath)
        {
            string workbookName = Path.GetFileNameWithoutExtension(filePath);
            var workbook = new AsposeWorkbook(filePath);
            workbook.WorkbookName = workbookName;
            return workbook;
        }

        public static AsposeWorkbook GetWorkbookFromStream(Stream stream, string workbookName)
        {
            var workbook = new AsposeWorkbook(stream);
            workbook.WorkbookName = workbookName;
            return workbook;
        }

        public static AsposeWorkbook GetWorkbookFromResource(string resourcePath)
        {
            string workbookName = resourcePath.Substring(resourcePath.LastIndexOf('.') + 1);
            workbookName = Path.GetFileNameWithoutExtension(workbookName);
            AsposeWorkbook workbook;
            var assembly = typeof(WorkbookRetriever).Assembly;
            using (Stream stream = assembly.GetManifestResourceStream(resourcePath))
            {
                workbook = new AsposeWorkbook(stream);
                workbook.WorkbookName = workbookName;
            }

            return workbook;
        }
    }
}
