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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.editBox4 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "التحويل إلى وورد";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.editBox3);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.editBox2);
            this.group1.Items.Add(this.editBox4);
            this.group1.Label = "تحديد";
            this.group1.Name = "group1";
            // 
            // editBox1
            // 
            this.editBox1.Label = "editBox1";
            this.editBox1.Name = "editBox1";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.CheckFileExists = true;
            this.saveFileDialog1.Filter = "ملفات وورد|*.doc,*.docx";
            this.saveFileDialog1.RestoreDirectory = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "ملفات وورد|*.doc,*.docx";
            this.openFileDialog1.RestoreDirectory = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.checkBox1);
            this.group2.Items.Add(this.checkBox2);
            this.group2.Items.Add(this.button1);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // editBox2
            // 
            this.editBox2.Label = "editBox2";
            this.editBox2.Name = "editBox2";
            // 
            // editBox3
            // 
            this.editBox3.Label = "editBox3";
            this.editBox3.Name = "editBox3";
            // 
            // editBox4
            // 
            this.editBox4.Label = "editBox4";
            this.editBox4.Name = "editBox4";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "checkBox1";
            this.checkBox1.Name = "checkBox1";
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "checkBox2";
            this.checkBox2.Name = "checkBox2";
            // 
            // button1
            // 
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Main Main
        {
            get { return this.GetRibbon<Main>(); }
        }
    }
}
