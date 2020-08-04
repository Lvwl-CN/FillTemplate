using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Stylet;
using System.IO;
using Microsoft.Office.Interop.Word;
using NPOI.XSSF.UserModel;

namespace FillTemplate.Pages
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ShellView : System.Windows.Window
    {
        public ShellView()
        {
            InitializeComponent();
        }

        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }
    }

    public class ShellViewModel : Screen
    {
        #region word
        public string TemplateFilePath { get; set; }

        /// <summary>
        /// 选择word模板文件
        /// </summary>
        public void SelectTemplate()
        {
            OpenFileDialog open = new OpenFileDialog()
            {
                Multiselect = false,
                Filter = "Word(*.docx)|*.docx",
                Title = "选择Word模板文件",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (open.ShowDialog() == true)
            {
                TemplateFilePath = open.FileName;
                ReadTemplateBookMarkAndCreateDataTable();
            }
        }


        public Visibility LoadingVisibility { get; set; } = Visibility.Collapsed;

        /// <summary>
        /// 读取模板书签并创建数据表
        /// </summary>
        public async void ReadTemplateBookMarkAndCreateDataTable()
        {
            LoadingVisibility = Visibility.Visible;
            var dt = DT;
            DT = null;
            await System.Threading.Tasks.Task.Run(() =>
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application() { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                var doc = app.Documents.OpenNoRepairDialog(TemplateFilePath, false, true);
                try
                {
                    List<string> list = new List<string>();
                    var bookmarks = doc.Bookmarks;
                    foreach (Bookmark bookmark in bookmarks)
                    {
                        list.Add(bookmark.Name);
                    }
                    DT = CreateDT(list, dt);
                }
                catch (Exception e)
                {

                }
                finally
                {
                    doc.Close(false);
                    app.Quit(false);
                }
            });
            LoadingVisibility = Visibility.Collapsed;
        }
        #endregion

        #region 列表

        public System.Data.DataTable DT { get; set; }

        /// <summary>
        /// 创建数据表
        /// </summary>
        /// <param name="bookmarks"></param>
        private System.Data.DataTable CreateDT(List<string> bookmarks, System.Data.DataTable dt)
        {
            if (bookmarks == null) return dt;
            Bookmarks = bookmarks;
            if (dt == null)
            {
                dt = new System.Data.DataTable();
                foreach (var bookmark in bookmarks)
                {
                    dt.Columns.Add(bookmark, typeof(string));
                }
            }
            else
            {
                foreach (var bookmark in bookmarks)
                {
                    if (!dt.Columns.Contains(bookmark))
                    {
                        dt.Columns.Add(bookmark, typeof(string));
                    }
                }
                foreach (DataColumn column in dt.Columns)
                {
                    if (!bookmarks.Contains(column.ColumnName)) dt.Columns.Remove(column);
                }
            }
            return dt;
        }
        public List<string> Bookmarks { get; set; }
        public string SelectedBookmark { get; set; }

        public int RowAddCount { get; set; } = 1;


        public bool CanAddRow { get { return DT != null && DT.Columns.Count > 0; } }
        public void AddRow()
        {
            for (int i = 0; i < RowAddCount; i++)
            {
                DT.Rows.Add(DT.NewRow());
            }
            var dt = DT;
            DT = null;
            DT = dt;
        }

        public bool CanClearAll { get { return DT != null && DT.Rows.Count > 0; } }
        public void ClearAll()
        {
            DT.Rows.Clear();
            var dt = DT;
            DT = null;
            DT = dt;
        }

        public bool CanClearEmpty { get { return DT != null && DT.Rows.Count > 0; } }
        public void ClearEmpty()
        {
            for (int i = DT.Rows.Count - 1; i >= 0; i--)
            {
                var row = DT.Rows[i];
                bool isNotEmpty = false;
                foreach (DataColumn column in DT.Columns)
                {
                    var temp = row[column];
                    if (temp == null || temp == DBNull.Value || string.IsNullOrEmpty(temp.ToString()))
                    {

                    }
                    else
                    {
                        isNotEmpty = true;
                        break;
                    }
                }
                if (!isNotEmpty) DT.Rows.Remove(row);
            }
            var dt = DT;
            DT = null;
            DT = dt;
        }

        /// <summary>
        /// 删除一行数据
        /// </summary>
        /// <param name="row"></param>
        public void DeleteRow(DataRowView rowView)
        {
            var row = rowView.Row;
            if (DT.Rows.IndexOf(row) >= 0)
            {
                DT.Rows.Remove(row);
                var dt = DT;
                DT = null;
                DT = dt;
            }
        }

        #endregion

        #region 底部操作
        public List<string> PrintTypes { get; set; } = new List<string>() { "单面", "双面" };
        public string SelectedPrintType { get; set; } = "单面";
        public int Copies { get; set; } = 1;

        public bool CanPrint { get { return DT != null && DT.Rows.Count > 0 && DT.Columns.Count > 0; } }
        /// <summary>
        /// 打印文件
        /// </summary>
        public async void Print()
        {
            LoadingVisibility = Visibility.Visible;
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                string printname = printDialog.PrintQueue.Name;
                bool manualduplexprint = "双面".Equals(SelectedPrintType);
                await System.Threading.Tasks.Task.Run(() =>
                {
                    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application() { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                    app.ActivePrinter = printname;
                    try
                    {
                        foreach (DataRow row in DT.Rows)
                        {
                            var doc = app.Documents.OpenNoRepairDialog(TemplateFilePath, false, true);
                            var bookmarks = doc.Bookmarks;
                            foreach (Bookmark bookmark in bookmarks)
                            {
                                if (DT.Columns.Contains(bookmark.Name))
                                {
                                    var text = row[bookmark.Name].ToString();
                                    bookmark.Range.Text = text;
                                }
                            }
                            doc.PrintOut(false, true, Copies: Copies, ManualDuplexPrint: manualduplexprint);
                            doc.Close(false);
                        }
                    }
                    catch (Exception e)
                    {

                    }
                    finally
                    {
                        app.Quit(false);
                    }
                });
            }
            LoadingVisibility = Visibility.Collapsed;
        }


        public bool CanExport { get { return DT != null && DT.Rows.Count > 0 && DT.Columns.Count > 0; } }
        /// <summary>
        /// 导出文件
        /// </summary>
        public async void Export()
        {
            LoadingVisibility = Visibility.Visible;
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var folderPath = folderBrowserDialog.SelectedPath;
                await System.Threading.Tasks.Task.Run(() =>
                {
                    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application() { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                    try
                    {
                        for (int i = 0; i < DT.Rows.Count; i++)
                        {
                            var row = DT.Rows[i];
                            var doc = app.Documents.OpenNoRepairDialog(TemplateFilePath, false, true);
                            var bookmarks = doc.Bookmarks;
                            foreach (Bookmark bookmark in bookmarks)
                            {
                                if (DT.Columns.Contains(bookmark.Name))
                                {
                                    var text = row[bookmark.Name].ToString();
                                    bookmark.Range.Text = text;
                                }
                            }
                            string filepath = string.Concat(folderPath, @"\", (!string.IsNullOrEmpty(SelectedBookmark) && Bookmarks.Contains(SelectedBookmark)) ? row[SelectedBookmark].ToString() : i.ToString(), ".docx");
                            doc.SaveAs2(filepath);
                            doc.Close(false);
                        }
                    }
                    catch (Exception e)
                    {

                    }
                    finally
                    {
                        app.Quit(false);
                    }
                });
            }
            LoadingVisibility = Visibility.Collapsed;
        }


        public bool CanExportData { get { return DT != null && DT.Rows.Count > 0 && DT.Columns.Count > 0; } }

        /// <summary>
        /// 导出数据
        /// </summary>
        public async void ExportData()
        {
            LoadingVisibility = Visibility.Visible;
            await Office.ExcelUtility.ExportDataToExcel(DT);
            LoadingVisibility = Visibility.Collapsed;
        }

        public bool CanInportData { get { return DT != null && DT.Columns.Count > 0; } }
        /// <summary>
        /// 导入数据
        /// </summary>
        public async void InportData()
        {
            LoadingVisibility = Visibility.Visible;
            var dt = DT;
            DT = null;
            DT = await Office.ExcelUtility.InputExcelToData(dt);
            LoadingVisibility = Visibility.Collapsed;
        }

        #endregion

    }

}
