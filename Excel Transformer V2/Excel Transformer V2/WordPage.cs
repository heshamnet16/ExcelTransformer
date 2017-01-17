using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.SmartTag;
using System.Runtime.InteropServices;
using System.Collections;
namespace Excel_Transformer
{
    public class WordPage
    {
        public delegate void delProgress (int value);
        public event delProgress Progress;

        #region "Properites"

        private Dictionary<string, Dictionary<string, string>> _DicData;
        public Dictionary<string, Dictionary<string, string>> TransformedData
        {
            get { return _DicData; }
            set { _DicData = value; }
        }

        private string _SourceFileName;
        public string SourceFileName
        {
            get { return _SourceFileName; }
            set { _SourceFileName = value; }
        }
        private string _DestinationFileName;
        public string DestinationFileName
        {
            get { return _DestinationFileName; }
            set { _DestinationFileName = value; }
        }
        private bool _OneFile;

        public bool OneFile
        {
            get { return _OneFile; }
            set { _OneFile = value; }
        }

        private bool _UseCom;
        public bool UseCom
        {
            get { return _UseCom; }
            set { _UseCom = value; }
        }
        #endregion

        private bool _STOP = false;
        public WordPage(string SourceFile, string DestFile, Dictionary<string, Dictionary<string, string>> data)
        {
            this._DestinationFileName = DestFile;
            this._SourceFileName = SourceFile;
            this._DicData = data;
        }
        public void Stop()
        {
            _STOP = true;
        }
        public void Start()
        {
            _STOP = false;
            if (this._UseCom)
                if (this.OneFile)
                    WriteToAllOneFile();
                else
                    WriteToFile();
            else
                if (this.OneFile)
                    WriteToAllOneFileOpenXml();
                else
                    WriteToFileOpenXml();

        }
        private bool WriteToAllOneFile()
        {
            Application wSapp = new Application();
            Application wDapp = new Application();
            //wSapp.Visible = true;
            //wDapp.Visible = true;
            object dontSave = WdSaveOptions.wdDoNotSaveChanges;
            if (System.IO.File.Exists(_DestinationFileName))
            {
                System.IO.File.Delete(_DestinationFileName);
            }
            System.IO.File.Copy(_SourceFileName, _DestinationFileName);
            object omissing = Type.Missing;
            object ofilename = _SourceFileName;
            object oDfilename = _DestinationFileName;
            bool FirstPage = true;
            Document wDdoc = wDapp.Documents.Open(ref oDfilename, ref omissing, ref omissing, ref omissing, ref omissing, ref omissing,
                ref omissing, ref omissing, ref omissing, ref omissing, ref omissing);
            float  i = 0f;
            float All = (float)_DicData.Keys.Count;
            foreach (string row in _DicData.Keys)
            {
                Document wSDoc = wSapp.Documents.Open(ref ofilename, ref omissing, ref omissing, ref omissing, ref omissing, ref omissing,
                        ref omissing, ref omissing, ref omissing, ref omissing, ref omissing);
                i++;
                if (_STOP)
                {
                    break;
                }
                foreach (string key in _DicData[row].Keys)
                {
                    if (_STOP)
                    {
                        break;
                    }
                    foreach (Bookmark bk in wSDoc.Bookmarks)
                    {
                        if (_STOP)
                        {
                            break;
                        }
                        if (bk.Name.ToLower() == key.ToString().ToLower())
                        {
                            bk.Select();
                            wSapp.Selection.Text = _DicData[row][key];
                        }
                    }//End Of Bookmarks foreach
                }//End of keys foreach
                wSapp.ActiveDocument.Content.Select();
                wSapp.Selection.Copy();
                if (!FirstPage)
                {
                    wDdoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                    wDapp.Selection.EndKey(WdUnits.wdStory, ref omissing);
                }
                else
                {
                    wDapp.ActiveDocument.Content.Select();
                    wDapp.Selection.Delete();
                    FirstPage = false;
                }
                wDapp.Selection.Paste();
                int prog = (int)((i / All) * 100);
                if (prog < 0) prog = 0;
                if (prog > 100) prog = 100;
                if(Progress != null)
                       Progress(prog);
                wSDoc.Close(dontSave, ref omissing, ref omissing);
            }//end of foreach rows
            //clean up
            wDdoc.Save();
            wDdoc.Close(dontSave,ref omissing, ref omissing);
            wDapp.Quit();
            wSapp.Quit(ref dontSave);
            Marshal.FinalReleaseComObject(wDapp);
            Marshal.FinalReleaseComObject(wDdoc);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return true;
        }

        private bool WriteToFile()
        {
            Application wSapp = new Application();
            //wSapp.Visible = true;
            //wDapp.Visible = true;
            object dontSave = WdSaveOptions.wdDoNotSaveChanges;
            object omissing = Type.Missing;
            float i = 0f;
            float All = (float)_DicData.Keys.Count;
            foreach (string row in _DicData.Keys)
            {
                if (System.IO.File.Exists(row))
                {
                    System.IO.File.Delete(row);
                }
                System.IO.File.Copy(_SourceFileName, row);
                object oDfilename = row;
                Document wSDoc = wSapp.Documents.Open(ref oDfilename, ref omissing, ref omissing, ref omissing, ref omissing, ref omissing,
                        ref omissing, ref omissing, ref omissing, ref omissing, ref omissing);
                i++;
                if (_STOP)
                {
                    break;
                }
                foreach (string key in _DicData[row].Keys)
                {
                    if (_STOP)
                    {
                        break;
                    }
                    foreach (Bookmark bk in wSDoc.Bookmarks)
                    {
                        if (_STOP)
                        {
                            break;
                        }
                        if (bk.Name.ToLower() == key.ToString().ToLower())
                        {
                            bk.Select();
                            wSapp.Selection.Text = _DicData[row][key];
                        }
                    }//End Of Bookmarks foreach
                }//End of keys foreach
                int prog = (int)((i / All) * 100);
                if (prog < 0) prog = 0;
                if (prog > 100) prog = 100;
                if (Progress != null)
                    Progress(prog);
                wSDoc.Save();
                wSDoc.Close(dontSave, ref omissing, ref omissing);
            }//end of foreach rows
            //clean up
            wSapp.Quit(ref dontSave);
            Marshal.FinalReleaseComObject(wSapp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return true;
        }


        private bool WriteToAllOneFileOpenXml()
        {
            float i = 0f;
            float All = (float)_DicData.Keys.Count;
            ArrayList WordFiles = new ArrayList();
            string FirstFile = "";
            foreach (string row in _DicData.Keys)
            {
                if (System.IO.File.Exists(row))
                {
                    System.IO.File.Delete(row);
                }
                System.IO.File.Copy(_SourceFileName, row);
                i++;
                if (_STOP)
                {
                    break;
                }
                foreach (string key in _DicData[row].Keys)
                {
                    if (_STOP)
                    {
                        break;
                    }
                    COWTranslation.TranslationToWord.UpadteTextAfterBookmarke(key, _DicData[row][key], row);
                }//End of keys foreach
                int prog = (int)((i / All) * 100);
                if (prog < 0) prog = 0;
                if (prog > 100) prog = 100;
                if (Progress != null)
                    Progress(prog);
                if (i > 1)
                    WordFiles.Add(row);
                else
                    FirstFile = row;
            }//end of foreach rows
            COWTranslation.DocMerger merger = new COWTranslation.DocMerger();
            if (System.IO.File.Exists(_DestinationFileName))
                System.IO.File.Delete(_DestinationFileName);
            System.IO.File.Copy(FirstFile, _DestinationFileName);
            System.IO.File.Delete(FirstFile);
            merger.DestinationFile = _DestinationFileName;
            merger.WordFiles = (string[])WordFiles.ToArray(typeof(string));
            merger.Merge();
            foreach (string fil in merger.WordFiles)
            {
                if (System.IO.File.Exists(fil))
                    System.IO.File.Delete(fil);
            }
            //clean up
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return true;
        }

        private bool WriteToFileOpenXml()
        {
            float i = 0f;
            float All = (float)_DicData.Keys.Count;
            foreach (string row in _DicData.Keys)
            {
                if (System.IO.File.Exists(row))
                {
                    System.IO.File.Delete(row);
                }
                System.IO.File.Copy(_SourceFileName, row);            
                i++;
                if (_STOP)
                {
                    break;
                }
                foreach (string key in _DicData[row].Keys)
                {
                    if (_STOP)
                    {
                        break;
                    }
                    COWTranslation.TranslationToWord.UpadteTextAfterBookmarke(key, _DicData[row][key], row);
                }//End of keys foreach
                int prog = (int)((i / All) * 100);
                if (prog < 0) prog = 0;
                if (prog > 100) prog = 100;
                if (Progress != null)
                    Progress(prog);
            }//end of foreach rows
            //clean up
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return true;
        }

    }
}
