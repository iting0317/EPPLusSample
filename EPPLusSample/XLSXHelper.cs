using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLusSample
{
    internal class XLSXHelper
    {
        /// <summary>
        /// 檢查格式
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public (bool IsValid, string Msg) ValidFile(string fullPath) 
        {
            if (!File.Exists(fullPath))
                return (false, $"{fullPath}檔案不存在");

            var supportedFormat = new string[] { ".xlsx" }; //注意：EPPlus不支援.xls格式
            if (!supportedFormat.Contains(Path.GetExtension(fullPath)))
                return (false, $"僅支援{string.Join(',', supportedFormat)}格式");

            return (true, "檢查通過");
        }

        /// <summary>
        /// xlsx轉csv字串
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public List<string> ReadExcelToStringList(Stream inputStream)
        {
            var list = new List<string>();
            StringBuilder sb = new StringBuilder();
            string cellValue;

            using (ExcelPackage excelPackage = new ExcelPackage(inputStream))
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.First();

                int startRow = ws.Dimension.Start.Row;      //起始列
                int endRow = ws.Dimension.End.Row;          //結束列
                int startColumn = ws.Dimension.Start.Column;//起始欄
                int endColumn = ws.Dimension.End.Column;    //結束欄

                for (int row = startRow; row <= endRow; row++)
                {
                    sb.Clear();

                    //跳過空白列
                    var rowRange = ws.Cells[row, startColumn, row, endColumn];
                    if (rowRange.Any(t => !string.IsNullOrEmpty(t.Text)) == false)
                        continue;

                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        cellValue = string.Empty;   //init

                        //處理excel的null cell
                        cellValue = (ws.Cells[row, col].Text ?? string.Empty).ToString().Replace("\"", "");

                        sb.Append(cellValue);
                        sb.Append(",");
                    }

                    sb.Remove(sb.Length - 1, 1);
                    list.Add(sb.ToString());
                }
            }

            return list;
        }

        /// <summary>
        /// xlsx轉DataTable
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public DataTable ReadExcelToDataTable(Stream stream)
        {
            var table = new DataTable();

            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.First();

                int startRow = ws.Dimension.Start.Row;      //起始列
                int endRow = ws.Dimension.End.Row;          //結束列
                int startColumn = ws.Dimension.Start.Column;//起始欄
                int endColumn = ws.Dimension.End.Column;    //結束欄

                //第一列取標題做為欄位名稱
                for (int col = startColumn; col <= endColumn; col++)
                    table.Columns.Add(new DataColumn(ws.Cells[1, col].Text));

                //略過第一列(標題列)
                for (int row = (startRow + 1); row <= endRow; row++)
                {
                    //跳過空白列
                    var rowRange = ws.Cells[row, startColumn, row, endColumn];
                    if (rowRange.Any(t => !string.IsNullOrEmpty(t.Text)) == false)
                        continue;

                    var dataRow = table.NewRow();
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        //需減1，因為EPPLUS的index從1開始
                        dataRow[col - 1] = (ws.Cells[row, col].Text ?? string.Empty).ToString().Replace("\"", "");
                    }

                    table.Rows.Add(dataRow);
                }
            }

            return table;
        }
    }
}
