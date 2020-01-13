using System;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using K12.Data;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using System.Linq;

namespace K12.缺曠通知單2015
{
    internal static class CommonMethods
    {
        //Excel報表
        public static void ExcelReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    #region 列印

                    string reportName;
                    string path;
                    Workbook wb;

                    object[] result = (object[])e.Result;
                    reportName = (string)result[0];
                    path = (string)result[1];
                    wb = (Workbook)result[2];

                    if (File.Exists(path))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                            if (!File.Exists(newPath))
                            {
                                path = newPath;
                                break;
                            }
                        }
                    }

                    try
                    {
                        wb.Save(path, FileFormatType.Docx);
                        FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");
                        System.Diagnostics.Process.Start(path);
                    }
                    catch
                    {
                        SaveFileDialog sd = new SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".xls";
                        sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                wb.Save(sd.FileName, FileFormatType.Docx);
                            }
                            catch
                            {
                                MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }

                    #endregion
                }
                else
                {
                    MsgBox.Show("發生錯誤:\n" + e.Error.Message);
                }
            }
            else
            {
                MsgBox.Show("已取消!!");
            }
        }

        //Word報表
        public static void WordReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    #region 列印

                    object[] result = (object[])e.Result;

                    string reportName = (string)result[0];
                    string path = (string)result[1];
                    Aspose.Words.Document doc = (Aspose.Words.Document)result[2];
                    string path2 = (string)result[3];
                    bool PrintStudetnList = (bool)result[4];
                    Aspose.Cells.Workbook wb = (Aspose.Cells.Workbook)result[5];

                    DataTable dt = (DataTable)result[6];

                    if (File.Exists(path))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                            if (!File.Exists(newPath))
                            {
                                path = newPath;
                                break;
                            }
                        }
                    }

                    if (File.Exists(path2))
                    {
                        int i = 1;
                        while (true)
                        {
                            string newPath = Path.GetDirectoryName(path2) + "\\" + Path.GetFileNameWithoutExtension(path2) + (i++) + Path.GetExtension(path2);
                            if (!File.Exists(newPath))
                            {
                                path2 = newPath;
                                break;
                            }
                        }
                    }

                    try
                    {
                        if (PrintStudetnList)
                        {
                            doc.Save(path, Aspose.Words.SaveFormat.Docx);
                            wb.Save(path2);
                            FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");

                            #region 產生CSV 檔
                            DateTime now = DateTime.Now;

                            String workingFolder = $"{System.Windows.Forms.Application.StartupPath}\\Reports";
                            if (!Directory.Exists(workingFolder))
                            {
                                Directory.CreateDirectory(workingFolder);
                            }
                            string csvfilePath = $"{workingFolder}\\缺曠通知單(弘文ePost)_{now.ToString("yyyyMMdd-HHmmss")}.txt";

                            exportToCSV(dt, csvfilePath);

                            //this.circularProgress1.Visible = false;
                            //this.circularProgress1.IsRunning = false;

                            if (MessageBox.Show("已成功匯出 .txt 檔案，是否要開啟檔案？", "完成", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                System.Diagnostics.Process.Start(csvfilePath);
                            } 
                            #endregion

                            System.Diagnostics.Process.Start(path);
                            System.Diagnostics.Process.Start(path2);
                        }
                        else
                        {
                            #region 產生CSV 檔
                            DateTime now = DateTime.Now;

                            String workingFolder = $"{System.Windows.Forms.Application.StartupPath}\\Reports";
                            if (!Directory.Exists(workingFolder))
                            {
                                Directory.CreateDirectory(workingFolder);
                            }
                            string csvfilePath = $"{workingFolder}\\缺曠通知單(弘文ePost)_{now.ToString("yyyyMMdd-HHmmss")}.txt";

                            exportToCSV(dt, csvfilePath);

                            //this.circularProgress1.Visible = false;
                            //this.circularProgress1.IsRunning = false;

                            if (MessageBox.Show("已成功匯出 .txt 檔案，是否要開啟檔案？", "完成", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                System.Diagnostics.Process.Start(csvfilePath);
                            }
                            #endregion

                            doc.Save(path, Aspose.Words.SaveFormat.Docx);
                            FISCA.Presentation.MotherForm.SetStatusBarMessage(reportName + "產生完成");
                            System.Diagnostics.Process.Start(path);
                        }
                    }
                    catch
                    {
                        SaveFileDialog sd = new SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".docx";
                        sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);

                            }
                            catch
                            {
                                MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }

                    #endregion
                }
                else
                {
                    MsgBox.Show("發生錯誤:\n" + e.Error.Message);
                }
            }
            else
            {
                MsgBox.Show("列印失敗,未取得缺曠資料!");
            }
        }

        /// <summary>
        ///  產生csv 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="csvFilePath"></param>
        private static void exportToCSV(DataTable dt, string csvFilePath)
        {
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            // Big5 編碼
            //File.WriteAllText(csvFilePath, sb.ToString(), Encoding.GetEncoding("Big5"));

            // Unicode 編碼  
            File.WriteAllText(csvFilePath, sb.ToString(), Encoding.GetEncoding("Unicode"));
        }

        //回報進度
        public static void Report_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState + "產生中...", e.ProgressPercentage);
        }

        internal static string GetChineseDayOfWeek(DateTime date)
        {
            string dayOfWeek = "";

            switch (date.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    dayOfWeek = "一";
                    break;
                case DayOfWeek.Tuesday:
                    dayOfWeek = "二";
                    break;
                case DayOfWeek.Wednesday:
                    dayOfWeek = "三";
                    break;
                case DayOfWeek.Thursday:
                    dayOfWeek = "四";
                    break;
                case DayOfWeek.Friday:
                    dayOfWeek = "五";
                    break;
                case DayOfWeek.Saturday:
                    dayOfWeek = "六";
                    break;
                case DayOfWeek.Sunday:
                    dayOfWeek = "日";
                    break;
            }

            return dayOfWeek;
        }

        //依班級座號排序
        public static int ClassSeatNoComparer(StudentRecord x, StudentRecord y)
        {
            string xx1 = (string.IsNullOrEmpty(x.RefClassID) ? "" : x.Class.Name) + "::";
            string xx2 = x.SeatNo.HasValue ? x.SeatNo.Value.ToString().PadLeft(2, '0') : "";
            string yy1 = (string.IsNullOrEmpty(y.RefClassID) ? "" : y.Class.Name) + "::";
            string yy2 = y.SeatNo.HasValue ? y.SeatNo.Value.ToString().PadLeft(2, '0') : "";
            xx1 += xx2;
            yy1 += yy2;
            return xx1.CompareTo(yy1);
        }
    }
}
