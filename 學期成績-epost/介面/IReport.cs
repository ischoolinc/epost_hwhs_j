
using System.Collections.Generic;

namespace hwhs.epost.學期成績通知單
{
    internal interface IReport
    {
        void Print(List<string> StudIDList);
    }
}
