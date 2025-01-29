using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using EPPlusLicenseContext = OfficeOpenXml.LicenseContext;
using OfficeOpenXml;
using System.Data;


namespace Sebtum.Models
{
    public class Excel
    {
        public DataTable ReadExcelFile(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = EPPlusLicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    string columnName = Convert.ToString(worksheet.Cells[1, col].Value);
                    dataTable.Columns.Add(columnName);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue = Convert.ToString(worksheet.Cells[row, col].Value);
                        dataRow[col - 1] = cellValue;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        public void SaveExel(DataTable dataTable, string filePath)
        {
            // Создаем новый пакет Excel
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Записываем заголовки столбцов
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                // Записываем данные из DataTable
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                // Сохраняем файл
                package.SaveAs(new FileInfo(filePath));
            }


        }
    }
}
