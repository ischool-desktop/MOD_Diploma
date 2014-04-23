using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiplomaReport
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            RibbonBarItem check = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"];
            check["報表"]["學籍相關報表"]["畢業證書(高中,高職,進校)"].Enable = Permissions.列印畢業證書權限;
            check["報表"]["學籍相關報表"]["畢業證書(高中,高職,進校)"].Click += delegate
            {
                DiplomaForm insert = new DiplomaForm();
                insert.ShowDialog();
            };

            Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.列印畢業證書, "畢業證書(高中,高職,進校)"));
        }
    }
}
