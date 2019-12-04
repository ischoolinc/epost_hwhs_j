
using System.Collections.Generic;

namespace hwhs.epost.定期評量通知單
{
    internal interface IReport
    {
        void Print(List<string> StudIDList);
    }
}
