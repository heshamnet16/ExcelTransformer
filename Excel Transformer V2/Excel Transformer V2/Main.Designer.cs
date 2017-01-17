namespace Excel_Transformer_V2
{
    partial class Main : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Main()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.txtColumnFrom = this.Factory.CreateRibbonEditBox();
            this.txtColumnTo = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.txtRowFrom = this.Factory.CreateRibbonEditBox();
            this.txtRowTo = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.txtFileName = this.Factory.CreateRibbonEditBox();
            this.chkOnefile = this.Factory.CreateRibbonCheckBox();
            this.chkUseCom = this.Factory.CreateRibbonCheckBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnStart = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.grpStat = this.Factory.CreateRibbonGroup();
            this.lblStatu = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.grpStat.SuspendLayout();
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "اختر المجلد الذي ترغب بحفظ الملفات فيه";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "ملفات وورد|*.doc;*.docx";
            this.saveFileDialog1.RestoreDirectory = true;
            this.saveFileDialog1.Title = "اختر اسم الملف النهائي";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "ملفات وورد|*.doc;*.docx";
            this.openFileDialog1.RestoreDirectory = true;
            this.openFileDialog1.Title = "اختر نمودج الوورد";
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.grpStat);
            this.tab1.Label = "التحويل إلى وورد";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.txtColumnFrom);
            this.group1.Items.Add(this.txtColumnTo);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.txtRowFrom);
            this.group1.Items.Add(this.txtRowTo);
            this.group1.Label = "تحديد";
            this.group1.Name = "group1";
            // 
            // txtColumnFrom
            // 
            this.txtColumnFrom.Label = "عمود البداية";
            this.txtColumnFrom.Name = "txtColumnFrom";
            this.txtColumnFrom.ScreenTip = "اسم العمود الذي سيبدأ التصدير من عنده";
            this.txtColumnFrom.SuperTip = "اسم العمود الذي سيبدأ التصدير من عنده";
            this.txtColumnFrom.Text = "A";
            // 
            // txtColumnTo
            // 
            this.txtColumnTo.Label = "عمود النهاية";
            this.txtColumnTo.Name = "txtColumnTo";
            this.txtColumnTo.ScreenTip = "اسم العمود الذي سيتوقف التصدير عنده";
            this.txtColumnTo.SuperTip = "اسم العمود الذي سيتوقف التصدير عنده";
            this.txtColumnTo.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // txtRowFrom
            // 
            this.txtRowFrom.Label = "صف البداية";
            this.txtRowFrom.Name = "txtRowFrom";
            this.txtRowFrom.ScreenTip = "رقم الصف الذي سيبدأ البرنامج من عنده";
            this.txtRowFrom.SuperTip = "رقم الصف الذي سيبدأ البرنامج من عنده";
            this.txtRowFrom.Text = "1";
            // 
            // txtRowTo
            // 
            this.txtRowTo.Label = "صف النهاية";
            this.txtRowTo.Name = "txtRowTo";
            this.txtRowTo.ScreenTip = "رقم الصف الذي سيتوقف التصدير عنده";
            this.txtRowTo.SuperTip = "رقم الصف الذي سيتوقف التصدير عنده";
            this.txtRowTo.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.txtFileName);
            this.group2.Items.Add(this.chkOnefile);
            this.group2.Items.Add(this.chkUseCom);
            this.group2.Items.Add(this.separator2);
            this.group2.Items.Add(this.btnStart);
            this.group2.Label = "تنفيذ";
            this.group2.Name = "group2";
            // 
            // txtFileName
            // 
            this.txtFileName.Label = "اسم الملف";
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.ScreenTip = "اسم الملف الناتج في حال عدم التجميع و يمكنك ادخال اسم العمود بين قوسين مجموعة [] " +
    "لكي يقوم البرنامج بأخذ اسماء الملفات من بينات هذا العمود مثلا [A]";
            this.txtFileName.SuperTip = "اسم الملف الناتج في حال عدم التجميع و يمكنك ادخال اسم العمود بين قوسين مجموعة [] " +
    "لكي يقوم البرنامج بأخذ اسماء الملفات من بينات هذا العمود مثلا [A]";
            this.txtFileName.Text = null;
            // 
            // chkOnefile
            // 
            this.chkOnefile.Label = "ملف واحد";
            this.chkOnefile.Name = "chkOnefile";
            this.chkOnefile.ScreenTip = "دمج في ملف واحد";
            this.chkOnefile.SuperTip = "دمج في ملف واحد";
            // 
            // chkUseCom
            // 
            this.chkUseCom.Checked = true;
            this.chkUseCom.Label = "Com استخدام";
            this.chkUseCom.Name = "chkUseCom";
            this.chkUseCom.ScreenTip = "ادوات Com تحاكي عملية الإدخال اليدوية تماما ومن عيوبها البطئ في التنفيذ";
            this.chkUseCom.SuperTip = "ادوات Com تحاكي عملية الإدخال اليدوية تماما ومن عيوبها البطئ في التنفيذ";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnStart
            // 
            this.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStart.Label = "ابدأ";
            this.btnStart.Name = "btnStart";
            this.btnStart.OfficeImageId = "ExportWord";
            this.btnStart.ScreenTip = "اضغط للبدء";
            this.btnStart.ShowImage = true;
            this.btnStart.SuperTip = "اضغط للبدء";
            this.btnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStart_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnAbout);
            this.group3.Label = "حول";
            this.group3.Name = "group3";
            // 
            // btnAbout
            // 
            this.btnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAbout.Label = "حول";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.OfficeImageId = "TentativeAcceptInvitation";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // grpStat
            // 
            this.grpStat.Items.Add(this.lblStatu);
            this.grpStat.Label = "التقدم";
            this.grpStat.Name = "grpStat";
            this.grpStat.Visible = false;
            // 
            // lblStatu
            // 
            this.lblStatu.Label = " ";
            this.lblStatu.Name = "lblStatu";
            // 
            // Main
            // 
            this.Name = "Main";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Main_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.grpStat.ResumeLayout(false);
            this.grpStat.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtColumnFrom;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtColumnTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtRowFrom;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtRowTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkOnefile;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkUseCom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblStatu;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpStat;
    }

    partial class ThisRibbonCollection
    {
        internal Main Main
        {
            get { return this.GetRibbon<Main>(); }
        }
    }
}
