﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
  public class WorkbookSerializer<T>
  {
    private string WorkbookName { get; set; } = string.Empty;
    private string FilePath { get; set; } = string.Empty;

    private List<WorksheetSerializer> worksheetSerializers = new List<WorksheetSerializer>();

    /// <summary>
    /// Writes out workbookClass object to an Excel Workbook file with a separate Worksheet for each List property having an Order attribute.
    /// Only properties that are Lists should be serialized.
    /// Defaults to current working directory and [nameof(T)].xlsx
    /// </summary>
    /// <param name="workbookClass"></param>
    /// <param name="filePath"></param>
    public void Serialize(T workbookClass, string filePath = "")
    {
      SetFileProperties(filePath);

      // Populate WorksheetSerializers from Workbook properties on workbookClass
      worksheetSerializers = GetWorksheetSerializers(workbookClass);
      // Create excel workbook
      WriteWorkbook();
      DisposeDataTables();
    }

    private List<WorksheetSerializer> GetWorksheetSerializers(T workbookClass)
    {
      List<WorksheetSerializer> serializers = new List<WorksheetSerializer>();
      var propertyInfos = typeof(T).GetProperties();
      foreach (var propertyInfo in propertyInfos)
      {
        string worksheetName = propertyInfo.Name;
        Type propertyType = propertyInfo.PropertyType;

        var genericListType = propertyType.GetGenericArguments()[0];
        var propertyListValue = (IList)propertyInfo.GetValue(workbookClass);

        var dataTableCreator = new DataTableConverter();
        var dataTable = dataTableCreator.CreateDataTableFor(propertyListValue, genericListType);

        var worksheetCreator = new WorksheetSerializer(dataTable, worksheetName);
        serializers.Add(worksheetCreator);
      }

      return serializers;
    }

    private void WriteWorkbook()
    {
      // Force creation through WorkbookCreator and WorkbookRetriever factory methods
      using (var workbook = WorkbookCreator.CreateWorkbookWithFilePath(FilePath))
      {
        CreateSheets(workbook);
        workbook.RemoveDefaultTab();
        workbook.Delete();
        workbook.Save();
      }
    }

    private void CreateSheets(AsposeWorkbook workbook)
    {
      foreach (var worksheetSerializer in worksheetSerializers)
      {
        worksheetSerializer.CreateSheet(workbook);
      }
    }

    private void DisposeDataTables()
    {
      foreach (var worksheetSerializer in worksheetSerializers)
      {
        worksheetSerializer.DisposeDataTable();
      }
    }

    private void SetFileProperties(string filePath)
    {
      FilePath = filePath;
      WorkbookName = Path.GetFileNameWithoutExtension(filePath);
      if (string.IsNullOrEmpty(WorkbookName))
      {
        WorkbookName = typeof(T).Name;
        FilePath = Path.Combine(FilePath, WorkbookName);
      }

      // if the file name does not have an extension, then add a default one for Excel
      if (!Path.HasExtension(FilePath))
      {
        FilePath += ".xlsx";
      }
    }
  }
}
