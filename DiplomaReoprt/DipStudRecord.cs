using FISCA.DSAUtil;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;

namespace DiplomaReport
{
    class DipStudRecord
    {
        public DipStudRecord(DataRow row)
        {
            id = "" + row["studentid"];
            id_number = "" + row["id_number"];
            name = "" + row["name"];
            english_name = "" + row["english_name"];

            if (!string.IsNullOrEmpty("" + row["diploma_number"]))
            {
                string xmlName = "<xml>" + row["diploma_number"] + "</xml>";
                XmlElement xml = DSXmlHelper.LoadXml(xmlName);
                diploma_number = xml.InnerText;
            }

            if (!string.IsNullOrEmpty("" + row["birthdate"]))
            {
                birthdate = tool.GetBirthdate("" + row["birthdate"]);
            }

            if (string.IsNullOrEmpty("" + row["studentdept"]))
            {
                department = "" + row["classdept"];
            }
            else
            {
                department = "" + row["studentdept"];
            }
        }

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string id { get; set; }

        /// <summary>
        /// 身分證字號
        /// </summary>
        public string id_number { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string name { get; set; }

        /// <summary>
        /// 英文姓名
        /// </summary>
        public string english_name { get; set; }

        /// <summary>
        /// 畢業證書字號
        /// </summary>
        public string diploma_number { get; set; }

        /// <summary>
        /// 生日
        /// </summary>
        public DateTime birthdate { get; set; }

        /// <summary>
        /// 科別
        /// </summary>
        public string department { get; set; }

    }
}
