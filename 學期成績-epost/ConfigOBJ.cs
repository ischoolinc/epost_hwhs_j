﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace hwhs_epost_semester
{
    class ConfigOBJ
    {
        public ConfigOBJ()
        {
            userDefinedConfig = new Dictionary<string, List<string>>();
        }
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartDate { get; set; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndDate { get; set; }
        /// <summary>
        /// 沒有資料即不印
        /// </summary>
        public bool PrintHasRecordOnly { get; set; }
        /// <summary>
        /// 範本
        /// </summary>
        public MemoryStream Template { get; set; }

        /// <summary>
        /// 選擇的設定檔
        /// </summary>
        public Dictionary<string, List<string>> userDefinedConfig { get; set; }

        /// <summary>
        /// 寄件人姓名
        /// </summary>
        public string ReceiveName { get; set; }
        /// <summary>
        /// 寄件人地址
        /// </summary>
        public string ReceiveAddress { get; set; }
        /// <summary>
        /// 缺曠名稱
        /// </summary>
        public string ConditionName { get; set; }
        /// <summary>
        /// 缺曠支數
        /// </summary>
        public string ConditionNumber { get; set; }
        /// <summary>
        /// 缺曠名稱2
        /// </summary>
        public string ConditionName2 { get; set; }
        /// <summary>
        /// 缺曠支數2
        /// </summary>
        public string ConditionNumber2 { get; set; }

        /// <summary>
        /// 是否列印學生清單
        /// </summary>
        public bool PrintStudentList { get; set; }


        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 試別名稱
        /// </summary>
        public string ExamName { get; set; }

        /// <summary>
        /// 試別ID
        /// </summary>
        public string ExamID { get; set; }

        /// <summary>
        /// 使用者選擇的科目名稱
        /// </summary>
        public List<string> SelSubjNameList { get; set; }

    }
}
