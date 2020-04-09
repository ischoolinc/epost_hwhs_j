using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace hwhs.epost.學期成績通知單
{
    class Permissions
    {
        public static string 學生學期成績通知單 { get { return "Student.SemesterGradeNotification_epost.hwhs.j.2019"; } }
        public static string 班級學期成績通知單 { get { return "Class.SemesterGradeNotification_epost.hwhs.j.2019"; } }

        public static bool 學生學期成績通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生學期成績通知單].Executable;
            }
        }

        public static bool 班級學期成績通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級學期成績通知單].Executable;
            }
        }
    }
}
