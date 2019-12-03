using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.缺曠通知單2015
{
    class Permissions
    {
        public static string 學生缺曠通知單 { get { return "K12.Behavior.Student.AbsenceNotificationSelectDateRangeForm.2013"; } }
        public static string 班級缺曠通知單 { get { return "K12.Behavior.Class.AbsenceNotificationSelectDateRangeForm.2013"; } }

        public static bool 學生缺曠通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生缺曠通知單].Executable;
            }
        }

        public static bool 班級缺曠通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級缺曠通知單].Executable;
            }
        }
    }
}
