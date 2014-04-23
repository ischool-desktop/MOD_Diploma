using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace DiplomaReport
{
    static class tool
    {
        static public QueryHelper _Q = new QueryHelper();

        /// <summary>
        /// 將字串Parse為DateTime
        /// </summary>
        /// <param name="birthdate">傳入欲Parse的內容</param>
        /// <returns>回傳DateTime</returns>
        static public DateTime GetBirthdate(string birthdate)
        {
            DateTime dt;
            DateTime.TryParse(birthdate, out dt);

            return dt;
        }

        /// <summary>
        /// 傳入XML,回傳要取得的Element
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="ElementName"></param>
        /// <returns>Element內容</returns>
        static public string GetSchoolInfo(XmlElement xml, string ElementName)
        {
            if (K12.Data.School.Configuration["學校資訊"].PreviousData != null)
            {
                if (K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode(ElementName) != null)
                {
                    return K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode(ElementName).InnerText;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
    }
}
