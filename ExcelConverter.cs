using Microsoft.VisualBasic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
namespace Duck.ExcelConverter
{
    public class ExcelConverter
    {
        public async Task<List<T>> ConvertToListAsync<T>(Stream stream, Dictionary<string, string> columnMappings, int headerRowIndex = 0, int skipRows = 0) where T : new()
        {
            var result = new List<T>();

            await Task.Run(() =>
            {
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                // Read header row to get column indexes
                IRow headerRow = sheet.GetRow(headerRowIndex);
                var headerIndexes = GetHeaderIndexes(headerRow, columnMappings.Keys.ToList());
                int startRowIndex = headerRowIndex + skipRows;
                for (int rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow sheetRow = sheet.GetRow(rowIndex);
                    if (sheetRow == null) continue;

                    T obj = new T();
                    foreach (var mapping in columnMappings)
                    {
                        var property = typeof(T).GetProperty(mapping.Value);
                        if (property != null && headerIndexes.ContainsKey(mapping.Key))
                        {
                            int colIndex = headerIndexes[mapping.Key];
                            var cell = sheetRow.GetCell(colIndex);
                            object cellValue = GetCellValue(cell, property.PropertyType);
                            
                            property.SetValue(obj, cellValue);
                        }
                    }

                    result.Add(obj);
                }
            });

            return result;
        }

        public async Task<List<T>> ConvertToListAsync<T>(byte[] byteStream, Dictionary<string, string> columnMappings) where T : new()
        {
            using (var memoryStream = new MemoryStream(byteStream))
            {
                return await ConvertToListAsync<T>(memoryStream, columnMappings);
            }
        }

        private Dictionary<string, int> GetHeaderIndexes(IRow headerRow, List<string> requiredHeaders)
        {
            var headerIndexes = new Dictionary<string, int>();

            for (int colIndex = 0; colIndex < headerRow.LastCellNum; colIndex++)
            {
                var cell = headerRow.GetCell(colIndex);
                if (cell != null && requiredHeaders.Contains(cell.ToString()))
                {
                    headerIndexes[cell.ToString()] = colIndex;
                }
            }

            return headerIndexes;
        }

        private object GetCellValue(ICell cell, Type propertyType)
        {
            if (cell == null) return null;

            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    if (propertyType == typeof(DateTime))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Boolean:
                            return cell.BooleanCellValue;
                        case CellType.Numeric:
                            return cell.NumericCellValue;
                        case CellType.String:
                            return cell.StringCellValue;
                    }
                    break;
            }
            return null;
        }
    }

}