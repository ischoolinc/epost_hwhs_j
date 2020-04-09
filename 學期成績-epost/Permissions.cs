using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace hwhs.epost.定期評量通知單
{
    class Permissions
    {
        public static string 學生定期評量通知單 { get { return "Student.ExamNotification.hwhs.j.2019"; } }
        public static string 班級定期評量通知單 { get { return "Class..ExamNotification.hwhs.j.2019"; } }

        public static bool 學生定期評量通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生定期評量通知單].Executable;
            }
        }

        public static bool 班級定期評量通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級定期評量通知單].Executable;
            }
        }
    }
}
