using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FillTemplate.Office
{
    public class ExcelUtility
    {
        public static async Task ExportDataToExcel(DataTable dt)
        {
            if (dt == null || dt.Rows.Count <= 0 || dt.Columns.Count <= 0) return;
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                Filter = "Excel|*.xlsx",
                Title = "文件保存路径",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                var savefilepath = saveFileDialog.FileName;
                await Task.Run(() =>
                {
                    var workbook = new XSSFWorkbook();
                    var sheet = workbook.CreateSheet("模板数据");
                    var rowhead = sheet.CreateRow(0);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        var cell = rowhead.CreateCell(i);
                        cell.SetCellValue(dt.Columns[i].ColumnName);
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            row.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (var fs = new FileStream(savefilepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        workbook.Write(fs);
                        workbook.Close();
                    }
                });
            }
        }

        public static async Task<DataTable> InputExcelToData(DataTable DT)
        {
            if (DT == null || DT.Columns.Count <= 0) return DT;
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Multiselect = false,
                Filter = "Excel|*.xlsx",
                Title = "选择模板数据文件",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (openFileDialog.ShowDialog() == true)
            {
                var filepath = openFileDialog.FileName;
                return await Task.Run(() =>
                {
                    using (var stream = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                    {
                        var workbook = new XSSFWorkbook(stream);
                        var sheetcount = workbook.NumberOfSheets;
                        if (sheetcount <= 0) return DT;
                        ISheet sheet;
                        if (sheetcount > 1)
                        {
                            List<string> sheetnames = new List<string>();
                            for (int i = 0; i < sheetcount; i++) sheetnames.Add(workbook[i].SheetName);
                            Pages.SheetSelect sheetSelect = new Pages.SheetSelect() { Sheets = sheetnames };
                            if (sheetSelect.ShowDialog() != true) return DT;
                            sheet = workbook.GetSheet(sheetSelect.Sheet);
                        }
                        else sheet = workbook[0];
                        if (sheet.LastRowNum <= 0) return DT;
                        var rowhead = sheet.GetRow(0);
                        if (rowhead.LastCellNum <= 0) return DT;

                        Dictionary<string, int> indexs = new Dictionary<string, int>();
                        for (int i = 0; i < rowhead.LastCellNum; i++)
                        {
                            var head = rowhead.GetCell(i).StringCellValue;
                            if (DT.Columns.Contains(head)) indexs[head] = i;
                        }
                        for (int i = 1; i <= sheet.LastRowNum; i++)
                        {
                            var dtrow = DT.NewRow();
                            var row = sheet.GetRow(i);
                            foreach (var item in indexs) dtrow[item.Key] = row.GetCell(item.Value).StringCellValue;
                            DT.Rows.Add(dtrow);
                        }
                        return DT;
                    }
                });
            }
            else return DT;
        }
    }
}
