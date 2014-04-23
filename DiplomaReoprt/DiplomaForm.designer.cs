namespace DiplomaReport
{
    partial class DiplomaForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.cboxSelectNow = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboItem1 = new DevComponents.Editors.ComboItem();
            this.comboItem2 = new DevComponents.Editors.ComboItem();
            this.comboItem3 = new DevComponents.Editors.ComboItem();
            this.comboItem4 = new DevComponents.Editors.ComboItem();
            this.comboItem5 = new DevComponents.Editors.ComboItem();
            this.comboItem6 = new DevComponents.Editors.ComboItem();
            this.lbSelectNow = new DevComponents.DotNetBar.LabelX();
            this.linkChange = new System.Windows.Forms.LinkLabel();
            this.lbHelp = new DevComponents.DotNetBar.LabelX();
            this.linkReportNameList = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(352, 137);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.ImeMode = System.Windows.Forms.ImeMode.On;
            this.btnExit.Location = new System.Drawing.Point(437, 137);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // cboxSelectNow
            // 
            this.cboxSelectNow.DisplayMember = "Text";
            this.cboxSelectNow.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboxSelectNow.FormattingEnabled = true;
            this.cboxSelectNow.ItemHeight = 19;
            this.cboxSelectNow.Items.AddRange(new object[] {
            this.comboItem1,
            this.comboItem2,
            this.comboItem3,
            this.comboItem4,
            this.comboItem5,
            this.comboItem6});
            this.cboxSelectNow.Location = new System.Drawing.Point(122, 14);
            this.cboxSelectNow.Name = "cboxSelectNow";
            this.cboxSelectNow.Size = new System.Drawing.Size(281, 25);
            this.cboxSelectNow.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboxSelectNow.TabIndex = 2;
            this.cboxSelectNow.SelectedIndexChanged += new System.EventHandler(this.comboBoxEx1_SelectedIndexChanged);
            // 
            // comboItem1
            // 
            this.comboItem1.Text = "高中";
            // 
            // comboItem2
            // 
            this.comboItem2.Text = "高中(含照片)";
            // 
            // comboItem3
            // 
            this.comboItem3.Text = "高職";
            // 
            // comboItem4
            // 
            this.comboItem4.Text = "高職(含照片)";
            // 
            // comboItem5
            // 
            this.comboItem5.Text = "進校";
            // 
            // comboItem6
            // 
            this.comboItem6.Text = "進校(含照片)";
            // 
            // lbSelectNow
            // 
            this.lbSelectNow.AutoSize = true;
            this.lbSelectNow.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbSelectNow.BackgroundStyle.Class = "";
            this.lbSelectNow.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbSelectNow.Location = new System.Drawing.Point(15, 16);
            this.lbSelectNow.Name = "lbSelectNow";
            this.lbSelectNow.Size = new System.Drawing.Size(101, 21);
            this.lbSelectNow.TabIndex = 3;
            this.lbSelectNow.Text = "請選擇列印種類";
            // 
            // linkChange
            // 
            this.linkChange.AutoSize = true;
            this.linkChange.BackColor = System.Drawing.Color.Transparent;
            this.linkChange.Location = new System.Drawing.Point(12, 145);
            this.linkChange.Name = "linkChange";
            this.linkChange.Size = new System.Drawing.Size(94, 17);
            this.linkChange.TabIndex = 4;
            this.linkChange.TabStop = true;
            this.linkChange.Text = "修改樣版(高中)";
            this.linkChange.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkChange_LinkClicked);
            // 
            // lbHelp
            // 
            this.lbHelp.AutoSize = true;
            this.lbHelp.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp.BackgroundStyle.Class = "";
            this.lbHelp.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp.Location = new System.Drawing.Point(27, 62);
            this.lbHelp.Name = "lbHelp";
            this.lbHelp.Size = new System.Drawing.Size(329, 56);
            this.lbHelp.TabIndex = 5;
            this.lbHelp.Text = "1.列印的種類分別為高中/高職/進校\r\n2.單一列印作業請勿超過500名學生\r\n3.單檔列印:每一名學生將以學號姓名編號儲存單一檔案";
            // 
            // linkReportNameList
            // 
            this.linkReportNameList.AutoSize = true;
            this.linkReportNameList.BackColor = System.Drawing.Color.Transparent;
            this.linkReportNameList.Location = new System.Drawing.Point(186, 145);
            this.linkReportNameList.Name = "linkReportNameList";
            this.linkReportNameList.Size = new System.Drawing.Size(86, 17);
            this.linkReportNameList.TabIndex = 6;
            this.linkReportNameList.TabStop = true;
            this.linkReportNameList.Text = "功能變數總表";
            this.linkReportNameList.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkReportNameList_LinkClicked);
            // 
            // DiplomaForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 169);
            this.Controls.Add(this.linkReportNameList);
            this.Controls.Add(this.lbHelp);
            this.Controls.Add(this.linkChange);
            this.Controls.Add(this.lbSelectNow);
            this.Controls.Add(this.cboxSelectNow);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.DoubleBuffered = true;
            this.Name = "DiplomaForm";
            this.Text = "畢業證書_高中,高職,進校(For Office 2007以上版本Word)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboxSelectNow;
        private DevComponents.Editors.ComboItem comboItem1;
        private DevComponents.Editors.ComboItem comboItem2;
        private DevComponents.Editors.ComboItem comboItem3;
        private DevComponents.Editors.ComboItem comboItem4;
        private DevComponents.Editors.ComboItem comboItem5;
        private DevComponents.DotNetBar.LabelX lbSelectNow;
        private System.Windows.Forms.LinkLabel linkChange;
        private DevComponents.DotNetBar.LabelX lbHelp;
        private System.Windows.Forms.LinkLabel linkReportNameList;
        private DevComponents.Editors.ComboItem comboItem6;
    }
}

