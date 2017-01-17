using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.SmartTag;
using System.Runtime.InteropServices;
namespace Excel_Transformer
{
    class WordPage
    {
        public delegate void delProgress (int value);
        public event delProgress Progress;

        #region "Properites"

        private Dictionary<int, Dictionary<char, string>> _DicData;
        public Dictionary <int,Dictionary<char, string>> TransformedData
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
        private int _StartRow;
        public int StartRow
        {
            get { return _StartRow; }
            set { _StartRow = value; }
        }
        private int _EndRow;
        public int EndRow
        {
            get { return _EndRow; }
            set { _EndRow = value; }
        }
        private char _StartColumn;
        public char StartColumn
        {
            get { return _StartColumn; }
            set { _StartColumn = value; }
        }
        private char _EndColumn;
        public char EndColumn
        {
            get { return _EndColumn; }
            set { _EndColumn = value; }
        }
        #endregion

        private bool _STOP = false;
        public WordPage(string SourceFile, string DestFile, char startCol, char EndCol, Dictionary<int, Dictionary<char, string>> data)
        {
            this._DestinationFileName = DestFile;
            this._SourceFileName = SourceFile;
            this._StartRow = 1;
            this._EndRow = 100;
            this._StartColumn = startCol;
            this._EndColumn = EndCol;
            this._DicData = data;
        }
        public void Stop()
        {
            _STOP = true;
        }
        public void Start()
        {
            _STOP = false;
            if (this.OneFile)
                WriteToOneFile();
        }
        private bool WriteToOneFile()
        {
            Application wSapp = new Application();
            Application wDapp = new Application();
            System.IO.File.Copy(_SourceFileName, _DestinationFileName);
            object omissing = Type.Missing;
            object ofilename = _SourceFileName;
            object oDfilename = _DestinationFileName;
            bool FirstPage = true;
            Document wSDoc = wSapp.Documents.Open(ref ofilename ,ref omissing,ref omissing,ref omissing,ref omissing,ref omissing,
                ref omissing, ref omissing, ref omissing,ref omissing,ref omissing);
            Document wDdoc = wDapp.Documents.Open(ref oDfilename, ref omissing, ref omissing, ref omissing, ref omissing, ref omissing,
                ref omissing, ref omissing, ref omissing, ref omissing, ref omissing);
            float  i = 0f;
            float All = (float)_DicData.Keys.Count;
            foreach (int row in _DicData.Keys)
            {
                i++;
                if (_STOP)
                {
                    break;
                }
                foreach (char key in _DicData[row].Keys)
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
                object what = WdGoToItem.wdGoToPage;
                object which = WdGoToDirection.wdGoToFirst;
                object count = 1;
                Range startRange = wSapp.Selection.GoTo(ref what, ref which, ref count, ref omissing);
                object count2 = (int)count + 3;
                Range endRange = wSapp.Selection.GoTo(ref what, ref which, ref count2, ref omissing);
                endRange.SetRange(startRange.Start, endRange.End - 1);
                endRange.Select();
                wSapp.Selection.Copy();
                if (!FirstPage)
                {
                    wDdoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);                    
                }
                else
                {
                    startRange = wDapp.Selection.GoTo(ref what, ref which, ref count, ref omissing);
                    count2 = (int)count + 3;
                    endRange = wDapp.Selection.GoTo(ref what, ref which, ref count2, ref omissing);
                    endRange.SetRange(startRange.Start, endRange.End - 1);
                    endRange.Select();
                    wDapp.Selection.Delete();
                    FirstPage = false;
                }
                wDapp.Selection.Paste();
                int prog = (int)((i / All) * 100);
                if (prog < 0) prog = 0;
                if (prog > 100) prog = 100;
                Progress(prog);
            }//end of foreach rows
            //clean up
            object dontSave = WdSaveOptions.wdDoNotSaveChanges;
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
    }
}
