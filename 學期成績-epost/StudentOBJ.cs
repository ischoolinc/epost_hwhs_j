using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace hwhs_epost_semester
{
    class StudentOBJ
    {
        public StudentOBJ()
        {
            //取得缺曠設定檔

            //以Dic建立儲存計數的物件
            studentAbsence = new Dictionary<string, int>();

            studentSemesterAbsence = new Dictionary<string, int>();

            studentMerit = new Dictionary<string, int>();

            studentSemesterMerit = new Dictionary<string, int>();

            studentAbsenceDetail = new Dictionary<string, Dictionary<string, string>>();
        }
        /// <summary>
        /// 學生物件
        /// </summary>
        public StudentRecord student { get; set; }

        /// <summary>
        /// 缺曠統計
        /// </summary>
        public Dictionary<string, int> studentAbsence = new Dictionary<string, int>();

        /// <summary>
        /// 缺曠學期統計
        /// </summary>
        public Dictionary<string, int> studentSemesterAbsence = new Dictionary<string, int>();

        /// <summary>
        /// 獎懲統計
        /// </summary>
        public Dictionary<string, int> studentMerit = new Dictionary<string, int>();

        /// <summary>
        /// 獎懲學期統計
        /// </summary>
        public Dictionary<string, int> studentSemesterMerit = new Dictionary<string, int>();

        /// <summary>
        /// 缺曠明細內容
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> studentAbsenceDetail = new Dictionary<string, Dictionary<string, string>>();

        public string TeacherName { get; set; }
        public string ClassName { get; set; }
        public string SeatNo { get; set; }
        public string StudentNumber { get; set; }

        //收件人地址
        public string address { get; set; }
        //郵遞區號
        public string ZipCode { get; set; }
        //郵遞區號第一碼
        public string ZipCode1 { get; set; }
        //郵遞區號第二碼
        public string ZipCode2 { get; set; }
        //郵遞區號第三碼
        public string ZipCode3 { get; set; }
        //郵遞區號第四碼
        public string ZipCode4 { get; set; }
        //郵遞區號第五碼
        public string ZipCode5 { get; set; }

        //監護人
        public string CustodianName { get; set; }
        //父親
        public string FatherName { get; set; }
        //母親
        public string MotherName { get; set; }

        //收件人地址
        public string ParentCode { get; set; }
    }
}
