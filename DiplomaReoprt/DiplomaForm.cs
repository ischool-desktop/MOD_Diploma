using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace DiplomaReport
{
    public partial class DiplomaForm : BaseForm
    {

        BackgroundWorker BGW = new BackgroundWorker();

        byte[] TemplateByte = Properties.Resources.改_畢業證書格式_高中;

        Dictionary<string, string> _PhotoGDict = new Dictionary<string, string>();

        //預設設定檔為 - 高中畢業證書(DiplomaReport00000)
        string SetupConfig = "DiplomaReport00000";

        public DiplomaForm()
        {
            InitializeComponent();

            BGW.RunWorkerCompleted += BGW_RunWorkerCompleted;
            BGW.DoWork += BGW_DoWork;

            cboxSelectNow.SelectedIndex = 0;
        }

        bool NowIsLock
        {
            set
            {
                cboxSelectNow.Enabled = value;
                linkChange.Enabled = value;
                linkReportNameList.Enabled = value;
                btnPrint.Enabled = value;
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //開始套表列印
            if (!BGW.IsBusy)
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    NowIsLock = false;
                    BGW.RunWorkerAsync();
                }
                else
                {
                    MsgBox.Show("未選擇學生!");
                }
            }
            else
            {
                MsgBox.Show("系統忙碌中,請稍後!!");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {

            #region 範本
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(SetupConfig);
            Aspose.Words.Document Template;
            if (ConfigurationInCadre.Template == null)
            {
                //如果範本為空,則建立一個預設範本
                Campus.Report.ReportConfiguration ConfigurationInCadre_1 = new Campus.Report.ReportConfiguration(SetupConfig);
                ConfigurationInCadre_1.Template = new Campus.Report.ReportTemplate(TemplateByte, Campus.Report.TemplateType.Word);
                //ConfigurationInCadre_1.Template = new Campus.Report.ReportTemplate(Properties.Resources.社團點名表_合併欄位總表, Campus.Report.TemplateType.Word);
                Template = ConfigurationInCadre_1.Template.ToDocument();
            }
            else
            {
                //如果已有範本,則取得樣板
                Template = new Document(new MemoryStream(ConfigurationInCadre.Template.ToBinary(), false));
            }
            #endregion

            List<string> StudentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;

            //取得學生照片
            if (StudentIDList.Count > 0)
            {
                _PhotoGDict.Clear();
                _PhotoGDict = K12.Data.Photo.SelectGraduatePhoto(StudentIDList);
            }

            DataTable MergeTable = new DataTable();
            MergeTable.Columns.Add("學校名稱");
            MergeTable.Columns.Add("學校英文名稱");
            MergeTable.Columns.Add("校長姓名");
            MergeTable.Columns.Add("校長英文姓名");
            MergeTable.Columns.Add("學年度");
            MergeTable.Columns.Add("學生");
            MergeTable.Columns.Add("身分證號");
            MergeTable.Columns.Add("畢業證書字號");
            MergeTable.Columns.Add("英文姓名");
            MergeTable.Columns.Add("科別");
            MergeTable.Columns.Add("科別英文");

            MergeTable.Columns.Add("照片1吋");
            MergeTable.Columns.Add("照片2吋");
            MergeTable.Columns.Add("生日");
            MergeTable.Columns.Add("生日年");
            MergeTable.Columns.Add("生日月");
            MergeTable.Columns.Add("生日日");
            MergeTable.Columns.Add("列印日期");
            MergeTable.Columns.Add("列印年");
            MergeTable.Columns.Add("列印月");
            MergeTable.Columns.Add("列印日");
            //姓名 ,  英文姓名 , 生日 , 科別 , 畢業證書字號 , 畢業照片
            //學年度 , 校長姓名 , 校長英文姓名 , 學校名稱 , 學校英文姓名

            StringBuilder sb = new StringBuilder();
            sb.Append("select student.id as studentid,student.id_number,student.name,student.english_name,student.diploma_number,student.birthdate,student.ref_dept_id as studentdept,class.ref_dept_id as classdept from student ");
            sb.Append("left join class on class.id=student.ref_class_id ");
            sb.Append("where student.id in ('" + string.Join("','", StudentIDList) + "')");

            DataTable td = tool._Q.Select(sb.ToString());

            DataTable tdDept = tool._Q.Select("select * from dept");
            Dictionary<string, string> DeptNameDic = new Dictionary<string, string>();
            Dictionary<string, string> DeptEnNameDic = new Dictionary<string, string>();

            foreach (DataRow row in tdDept.Rows)
            {
                string deptid = "" + row["id"];
                string deptname = "" + row["name"];
                if (string.IsNullOrEmpty(deptid))
                    continue;

                if (!DeptNameDic.ContainsKey(deptid))
                {
                    DeptNameDic.Add(deptid, deptname);
                }

                if (!DeptEnNameDic.ContainsKey(deptid))
                {
                    DeptEnNameDic.Add(deptname, "");
                }
            }

            XmlElement Data = SmartSchool.Customization.Data.SystemInformation.Configuration["科別中英文對照表"];
            foreach (XmlElement var in Data)
            {
                string chinese = var.GetAttribute("Chinese");
                string english = var.GetAttribute("English");
                if (DeptEnNameDic.ContainsKey(chinese))
                {
                    DeptEnNameDic[chinese] = english;
                }
            }

            foreach (DataRow row in td.Rows)
            {
                //學生基本資料
                DipStudRecord dsr = new DipStudRecord(row);

                //Merge用的Row
                DataRow MergeRow = MergeTable.NewRow();

                MergeRow["學校名稱"] = School.ChineseName;
                MergeRow["學校英文名稱"] = School.EnglishName;
                MergeRow["校長姓名"] = tool.GetSchoolInfo(K12.Data.School.Configuration["學校資訊"].PreviousData, "ChancellorChineseName");
                MergeRow["校長英文姓名"] = tool.GetSchoolInfo(K12.Data.School.Configuration["學校資訊"].PreviousData, "ChancellorEnglishName");
                MergeRow["學年度"] = School.DefaultSchoolYear;
                MergeRow["學生"] = dsr.name;
                MergeRow["身分證號"] = dsr.name;
                MergeRow["畢業證書字號"] = dsr.diploma_number;
                MergeRow["英文姓名"] = dsr.english_name;
                if (DeptNameDic.ContainsKey(dsr.department))
                {
                    string deptName = DeptNameDic[dsr.department];
                    MergeRow["科別"] = deptName;

                    if (!string.IsNullOrEmpty(deptName))
                    {
                        if (DeptEnNameDic.ContainsKey(deptName))
                        {
                            MergeRow["科別英文"] = DeptEnNameDic[deptName];
                        }
                        else
                        {
                            MergeRow["科別英文"] = string.Empty;
                        }
                    }
                    else
                    {
                        MergeRow["科別英文"] = string.Empty;
                    }
                }
                else
                {
                    MergeRow["科別"] = string.Empty;
                }

                if (_PhotoGDict.ContainsKey(dsr.id))
                    MergeRow["照片1吋"] = _PhotoGDict[dsr.id];
                else
                    MergeRow["照片1吋"] = string.Empty;

                if (_PhotoGDict.ContainsKey(dsr.id))
                    MergeRow["照片2吋"] = _PhotoGDict[dsr.id];
                else
                    MergeRow["照片2吋"] = string.Empty;

                if (dsr.birthdate != null)
                {
                    MergeRow["生日"] = dsr.birthdate.ToShortDateString();
                    MergeRow["生日年"] = dsr.birthdate.Year.ToString();
                    MergeRow["生日月"] = dsr.birthdate.Month.ToString();
                    MergeRow["生日日"] = dsr.birthdate.Day.ToString();
                }
                else
                {
                    MergeRow["生日"] = string.Empty;
                    MergeRow["生日年"] = string.Empty;
                    MergeRow["生日月"] = string.Empty;
                    MergeRow["生日日"] = string.Empty;
                }

                MergeRow["列印日期"] = DateTime.Today.ToShortDateString();
                MergeRow["列印年"] = DateTime.Today.Year.ToString();
                MergeRow["列印月"] = DateTime.Today.Month.ToString();
                MergeRow["列印日"] = DateTime.Today.Day.ToString();

                MergeTable.Rows.Add(MergeRow);
            }

            Document PageOne = (Document)Template.Clone(true);
            PageOne.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
            PageOne.MailMerge.Execute(MergeTable);
            PageOne.MailMerge.DeleteFields();
            e.Result = PageOne;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            NowIsLock = true;
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    //列印完成
                    Document PageOne = (Document)e.Result;

                    try
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                        SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        string SelectName = ((DevComponents.Editors.ComboItem)cboxSelectNow.SelectedItem).Text;
                        SaveFileDialog1.FileName = "畢業證書(" + SelectName + ")";

                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            PageOne.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                            return;
                        }
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                        return;
                    }

                    this.Close();

                }
                else
                {
                    MsgBox.Show("系統發生錯誤:\n" + e.Error.Message);
                }
            }
            else
            {
                MsgBox.Show("列印已取消!!");
            }
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldName == "照片1吋" || e.FieldName == "照片2吋")
            {
                #region 畢業照片
                if (!string.IsNullOrEmpty(e.FieldValue.ToString()))
                {
                    byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                    if (photo != null && photo.Length > 0)
                    {
                        DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                        photoBuilder.MoveToField(e.Field, true);
                        e.Field.Remove();
                        //Paragraph paragraph = photoBuilder.InsertParagraph();// new Paragraph(e.Document);
                        Shape photoShape = new Shape(e.Document, ShapeType.Image);
                        photoShape.ImageData.SetImage(photo);
                        photoShape.WrapType = WrapType.Inline;
                        //Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                        //cell.CellFormat.LeftPadding = 0;
                        //cell.CellFormat.RightPadding = 0;
                        if (e.FieldName == "照片1吋")
                        {
                            // 1吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(25);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(35);
                        }
                        else
                        {
                            //2吋
                            photoShape.Width = ConvertUtil.MillimeterToPoint(35);
                            photoShape.Height = ConvertUtil.MillimeterToPoint(45);
                        }
                        //paragraph.AppendChild(photoShape);
                        photoBuilder.InsertNode(photoShape);
                    }
                }
                #endregion
            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //樣版說明
            string SelectName = ((DevComponents.Editors.ComboItem)cboxSelectNow.SelectedItem).Text;
            linkChange.Text = "修改樣版(" + SelectName + ")";

            if (cboxSelectNow.SelectedIndex == 0)
            {
                TemplateByte = Properties.Resources.改_畢業證書格式_高中;
                SetupConfig = "DiplomaReport00000";
            }
            else if (cboxSelectNow.SelectedIndex == 1)
            {
                SetupConfig = "DiplomaReport00001";
                TemplateByte = Properties.Resources.改_畢業證書格式_高中含照片;
            }
            else if (cboxSelectNow.SelectedIndex == 2)
            {
                SetupConfig = "DiplomaReport00002";
                TemplateByte = Properties.Resources.改_畢業證書格式_高職;
            }
            else if (cboxSelectNow.SelectedIndex == 3)
            {
                SetupConfig = "DiplomaReport00003";
                TemplateByte = Properties.Resources.改_畢業證書格式_高職含照片;
            }
            else if (cboxSelectNow.SelectedIndex == 4)
            {
                SetupConfig = "DiplomaReport00004";
                TemplateByte = Properties.Resources.改_畢業證書格式_進校;
            }
            else if (cboxSelectNow.SelectedIndex == 5)
            {
                SetupConfig = "DiplomaReport00005";
                TemplateByte = Properties.Resources.改_畢業證書格式_進校含照片;
            }
        }

        private void linkReportNameList_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "畢業證書_合併欄位總表.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.功能變數總表, 0, Properties.Resources.功能變數總表.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkChange_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(SetupConfig);
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm;
            if (ConfigurationInCadre.Template != null)
            {
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(TemplateByte, Campus.Report.TemplateType.Word));
            }
            else
            {
                ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(TemplateByte, Campus.Report.TemplateType.Word);
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(TemplateByte, Campus.Report.TemplateType.Word));
            }
            //預設名稱
            string SelectName = ((DevComponents.Editors.ComboItem)cboxSelectNow.SelectedItem).Text;
            TemplateForm.DefaultFileName = "畢業證書_範本_" + SelectName;
            //如果回傳為OK
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                //設定後樣試,回傳
                ConfigurationInCadre.Template = TemplateForm.Template;
                //儲存
                ConfigurationInCadre.Save();
            }
        }
    }
}
