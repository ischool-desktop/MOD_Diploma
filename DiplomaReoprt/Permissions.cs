using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiplomaReport
{
    class Permissions
    {
        public static string 列印畢業證書 { get { return "DiplomaReoprt.Report.000001"; } }
        public static bool 列印畢業證書權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[列印畢業證書].Executable;
            }
        }
    }
}
