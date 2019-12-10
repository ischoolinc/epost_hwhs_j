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

namespace hwhs.epost.定期評量通知單
{
    internal static class CommonMethods
    {
        
        //Word報表
        public static void WordReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    #region 列印

                    object[] result = (object[])e.Result;

                    DataTable dt = (DataTable)result[0];
                    
                    try
                    {
                        #region 產生CSV 檔
                        DateTime now = DateTime.Now;

                        String workingFolder = $"{System.Windows.Forms.Application.StartupPath}\\Reports";
                        if (!Directory.Exists(workingFolder))
                        {
                            Directory.CreateDirectory(workingFolder);
                        }
                        string csvfilePath = $"{workingFolder}\\定期評量通知單(弘文ePost)_{now.ToString("yyyyMMdd-HHmmss")}.csv";

                        exportToCSV(dt, csvfilePath);

                        if (MessageBox.Show("已成功匯出 CSV 檔案，是否要開啟檔案？", "完成", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(csvfilePath);
                        }
                        #endregion

                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
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

            File.WriteAllText(csvFilePath, sb.ToString(), Encoding.GetEncoding("Big5"));
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
