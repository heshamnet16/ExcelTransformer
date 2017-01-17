using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel_Transformer;

namespace Excel_Transformer_V2
{
    public partial class Main
    {
        WordPage wp = null;
        private void Main_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private string ShowFolder()
        {
            string ret=null;
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ret = folderBrowserDialog1.SelectedPath;
            }
            else
                ret = null;
            return ret;
        }
        private string ShowFileOpen()
        {
            string ret = null;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ret = openFileDialog1.FileName;
            }
            else
                ret = null;
            return ret;
        }
        private string ShowFileSave()
        {
            string ret = null;
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ret = saveFileDialog1.FileName;
            }
            else
                ret = null;
            return ret;
        }

        private bool IsNumber(string str)
        {
            foreach (char c in str)
            {
                if (!char.IsNumber(c))
                    return false;
            }
            return true;
        }
        private bool IsLetter(string str)
        {
            foreach (char c in str)
            {
                if (!char.IsLetter(c))
                    return false;
            }
            return true;
        }
        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            if (!IsNumber(txtRowFrom.Text) || !IsNumber(txtRowTo.Text))
            {
                System.Windows.Forms.MessageBox.Show("استخدم الارقام فقط لإدخال رقم الصف", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            if (!IsLetter(txtColumnFrom.Text) || !IsLetter(txtColumnTo.Text))
            {
                System.Windows.Forms.MessageBox.Show("استخدم اسماء الأعمدة بشكل صحيح", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            if (ExcelColumnNameToNumber(txtColumnFrom.Text) > ExcelColumnNameToNumber(txtColumnTo.Text))
            {
                System.Windows.Forms.MessageBox.Show("استخدم اسماء الأعمدة بشكل صحيح", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            if (int.Parse(txtRowFrom.Text) > int.Parse(txtRowTo.Text))
            {                
                System.Windows.Forms.MessageBox.Show("ادخل مجال الصفوف بشكل صحيح", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            if(HasMark(txtFileName.Text))
                if (!IsLetter(GetMarksValue(txtFileName.Text)))
                {
                    System.Windows.Forms.MessageBox.Show("ادخل اسم العمود بشكل صحيح", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

            if (btnStart.Label == "توقف")
            {
                if (wp != null)
                    wp.Stop();
                btnStart.Label = "ابدأ";
                btnStart.OfficeImageId = "ExportWord";
                grpStat.Visible = false;
                backgroundWorker1.CancelAsync();
                return;
            }
            try
            {
                string Sfile = ShowFileOpen();
                if (Sfile == null) return;
                    if (!chkOnefile.Checked)
                    {
                        //To Folder
                        string ofld = ShowFolder();
                        if (ofld != null)
                        {
                            grpStat.Visible = true;
                            wp = new WordPage(Sfile, "", GetAllDataM(ofld));
                            wp.OneFile = chkOnefile.Checked;
                            wp.UseCom = chkUseCom.Checked;
                            wp.Progress += wp_Progress;
                            btnStart.Label = "توقف";
                            btnStart.OfficeImageId = "PrintPreviewClose";
                            System.Windows.Forms.Application.DoEvents();
                            backgroundWorker1.RunWorkerAsync(wp);
                            //wp.Start();
                            //System.Windows.Forms.MessageBox.Show("انتهى", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        //To File
                        grpStat.Visible = true;
                        string Dfile = ShowFileSave();
                        if (Dfile == null) return;
                        wp = new WordPage(Sfile, Dfile, GetAllDataM(System.IO.Path.GetDirectoryName(Dfile)));
                        wp.OneFile = chkOnefile.Checked;
                        wp.UseCom = chkUseCom.Checked;
                        wp.Progress += wp_Progress;
                        btnStart.Label = "توقف";
                        btnStart.OfficeImageId = "PrintPreviewClose";
                        System.Windows.Forms.Application.DoEvents();
                        backgroundWorker1.RunWorkerAsync(wp);
                        //wp.Start();
                        //System.Windows.Forms.MessageBox.Show("انتهى", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }              
                //btnStart.Label = "ابدأ";
                //btnStart.OfficeImageId = "ExportWord";
                //grpStat.Visible = false;
            }
            catch (Exception ex)
            {
                btnStart.Label = "ابدأ";
                grpStat.Visible = false;
                throw ex;
            }        
        }

        void wp_Progress(int value)
        {            
            backgroundWorker1.ReportProgress(value);
            //lblStatu.Label = value.ToString() + "%";            
            //System.Windows.Forms.Application.DoEvents();
            //this.ResumeLayout();            
        }
        public static string GetExcelColumnName(int columnNumber)
        {
            unsafe { 
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
            }
        }
        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
        private string GetMarksValue(string marks)
        {
            if (HasMark(marks))
            {
                string colName = marks.Substring(marks.IndexOf('[') + 1, marks.IndexOf(']') - marks.IndexOf('[') - 1);
                return colName;
            }
            else
                return null;
        }
        private bool HasMark(string str)
        {
            if (txtFileName.Text.Contains("[") && txtFileName.Text.Contains("]"))
                return true;
            else
                return false;
        }
        public Dictionary<string, Dictionary<string, string>> GetAllDataM(string dFolder)
        {
            if (dFolder == null) return null;
            else
                if (!dFolder.EndsWith("\\")) dFolder += "\\";
            Application eApp = Globals.ThisAddIn.Application;
            Workbook eWbook = eApp.ActiveWorkbook;
            Worksheet eSheet = (Worksheet)eApp.ActiveSheet;
            int i, j;
            int EndRows = int.Parse(txtRowTo.Text);
            int EndCol = ExcelColumnNameToNumber(txtColumnTo.Text);
            Dictionary<string, Dictionary<string, string>> AllData = new Dictionary<string, Dictionary<string, string>>();
            for (i = int.Parse(txtRowFrom.Text); i <= EndRows; i++)
            {
                Dictionary<string, string> oneRow = new Dictionary<string, string>();
                for (j = ExcelColumnNameToNumber(txtColumnFrom.Text); j <= EndCol; j++)
                {
                    object oval = (eSheet.Cells[i, j] as Range).Value2;
                    if (oval != null)
                    {
                        string val = oval.ToString();
                        string ColName = GetExcelColumnName(j);
                        oneRow.Add(ColName, val);
                    }
                    //oneRow.Add()
                }
                string FileN;
                if (txtFileName.Text.Length > 0)
                    if (txtFileName.Text.Contains("[") && txtFileName.Text.Contains("]"))
                    {
                        string FinalFileName = txtFileName.Text;
                        string colName = FinalFileName.Substring(FinalFileName.IndexOf('[') + 1, FinalFileName.IndexOf(']') - FinalFileName.IndexOf('[')-1);
                        int colN = ExcelColumnNameToNumber(colName);
                        object oval = (eSheet.Cells[i, colN] as Range).Value2;
                        if (oval != null)
                        {
                            string val = oval.ToString();
                            FinalFileName = FinalFileName.Replace("[" + colName + "]", val);
                            FileN = dFolder + FinalFileName + (FinalFileName.ToLower().EndsWith(".docx") ? "" : ".docx");

                        }
                        else
                        {
                            FileN = dFolder + i.ToString() + txtFileName.Text + (txtFileName.Text.ToLower().EndsWith(".docx") ? "" : ".docx");
                        }
                    }
                    else
                    {
                        FileN = dFolder + i.ToString() + txtFileName.Text+(txtFileName.Text.ToLower().EndsWith(".docx") ? "" : ".docx");
                    }
                else
                    FileN = dFolder + i.ToString() + ".docx";
                if (!System.IO.Directory.Exists(dFolder))
                {
                    System.IO.Directory.CreateDirectory(dFolder);
                }
                AllData.Add(FileN, oneRow);
            }//Row for
            return AllData;
        }
        public Dictionary<string, Dictionary<string, string>> GetAllData()
        {
            Application eApp = Globals.ThisAddIn.Application;
            Workbook eWbook = eApp.ActiveWorkbook;
            Worksheet eSheet =(Worksheet)eApp.ActiveSheet;
            int i, j;
            int EndRows = int.Parse(txtRowTo.Text);
            int EndCol = ExcelColumnNameToNumber(txtColumnTo.Text);

            Dictionary<string, Dictionary<string, string>> AllData = new Dictionary<string, Dictionary<string, string>>();
            for (i = int.Parse(txtRowFrom.Text); i <= EndRows; i++)
            {
                Dictionary<string, string> oneRow = new Dictionary<string, string>();
                for (j = ExcelColumnNameToNumber(txtColumnFrom.Text); j <= EndCol; j++)
                {
                    object oval = (eSheet.Cells[i, j] as Range).Value2;
                    if (oval != null)
                    {
                        string val = oval.ToString();
                        string ColName = GetExcelColumnName(j);
                        oneRow.Add(ColName, val);
                    }
                    //oneRow.Add()
                }
                AllData.Add(i.ToString(), oneRow);
            }//Row for
            return AllData;
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            WordPage wp1 = (WordPage)e.Argument;
            wp.Progress += wp_Progress;
            wp1.Start();
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            lblStatu.Label = e.ProgressPercentage.ToString() + "%";
            System.Windows.Forms.Application.DoEvents();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("انتهى", "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            btnStart.Label = "ابدأ";
            btnStart.OfficeImageId = "ExportWord";
            grpStat.Visible = false;
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            string msg = "المحول للورد" + Environment.NewLine + "يقوم البرنامج بتحويل كل سطر من ملف الاكسل إلى صفحة وورد وذلك بتسمية الاشارات المرجعية باسماء اعمدة ملف الاكسل.";
            msg += Environment.NewLine + "تمت برمجته لصالح شعبة الهلال الاحمر العربي السوري في تلبيسة";
            msg += Environment.NewLine + "فكرة و برمجة و تدقيق فريق عمل الشعبة لعام 2012";
            System.Windows.Forms.MessageBox.Show(msg, "", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
