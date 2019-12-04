using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace hwhs.epost.定期評量通知單
{
    class AttendanceOBJ
    {
        /// <summary>
        /// 缺曠類別(一般/集會)
        /// </summary>
        public string AttendanceType { get; set; }

        /// <summary>
        /// 缺曠名稱
        /// </summary>
        public string AttendanceName { get; set; }

        /// <summary>
        /// 缺曠計數
        /// </summary>
        public int AttendanceCount { get; set; }
    }
}
