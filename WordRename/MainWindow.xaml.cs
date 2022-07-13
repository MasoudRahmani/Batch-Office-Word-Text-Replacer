using System.Collections.ObjectModel;
using System.Windows;
using System.IO;
using _Word = Microsoft.Office.Interop.Word;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace WordRename
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public ObservableCollection<MyDocument> Mydocuments { get; set; }
        private string _logPath;
        private StreamWriter _logger;
        private string _outputfoldername = "WrdRn_Output";

        public MainWindow()
        {
            Mydocuments = new ObservableCollection<MyDocument>();
            InitializeComponent();
            DataContext = this;
            Closed += MainWindow_Closed;
            Closing += Main_Closing;
        }

        private void Main_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_logPath))
            {
                LogMe("End of Activity");
            }
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            if (_logger != null) _logger.Close();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".doc | .docx";
            dlg.Filter = "Word Document (.doc;.docx)|*.doc;*.docx";

            // Display OpenFileDialog by calling ShowDialog method
            var result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                srcbox.Text = dlg.FileName;
            }
        }

        private string _path;
        private void Start_Click(object sender, RoutedEventArgs e)
        {
            _path = srcbox.Text;
            string find = findbox.Text;
            string replace = replaceBox.Text;
            bool matchhw = (bool)matchHW.IsChecked;

            if (string.IsNullOrEmpty(findbox.Text) | string.IsNullOrWhiteSpace(_path))
            {
                ShowStatus("Find and Path cannot be empty");
                return;
            }
            ShowStatus("Working...");
            Working(true);

            PathType ptype = GetPathType(_path);
            if (ptype != PathType.invalid)
            {
                var newdir = Path.Combine(Path.GetDirectoryName(_path), _outputfoldername);
                Directory.CreateDirectory(newdir);
                SetLogPath(_path);
            }

            switch (ptype)
            {
                case PathType.File:
                    LogMe($"File Job: '{_path}'");
                    LogMe($"^^^^^^^^^^^^^^^^^ find: '{find}' & replace: '{replace}' ^^^^^^^^^^^^^^^^^");
                    Task.Factory.StartNew(new Action(() =>
                    {
                        Find_Replace(_path, find, replace, matchhw, _outputfoldername);
                    })).ContinueWith((s) => { DispatchMe(() => { Working(false); }); ShowStatus("Done."); }); ;
                    break;
                case PathType.Directory:

                    Task.Factory.StartNew(new Action<object>((object p) =>
                   {
                       var path = (string)p;
                       var files = Directory.EnumerateFiles(path).Where(f => f.ToLower().EndsWith("docx") || f.ToLower().EndsWith("doc")).ToList();
                       files.RemoveAll(x => { return Path.GetFileName(x).StartsWith("~$"); });

                       LogMe($"Directory Job for *{files.Count}* files in: '{path}'");
                       LogMe($"^^^^^^^^^^^^^^^^^ find: '{find}' & replace: '{replace}' ^^^^^^^^^^^^^^^^^");

                       files.ForEach(x =>
                       {
                           Find_Replace(x, find, replace, matchhw, _outputfoldername);
                       });
                   }), _path).ContinueWith((s) => { DispatchMe(() => { Working(false); ShowStatus("Done."); }); });
                    break;
                case PathType.invalid:
                    ShowStatus("Inavlid Adress");
                    Working(false);
                    break;
                default:
                    ShowStatus("Exception, Path Type is wrong");
                    Working(false);
                    break;
            }
        }

        private void Find_Replace(string document, string findText, string replaceText, bool matchhw, string outputFolder)
        {
            var newDoc = Path.Combine(Path.GetDirectoryName(document), outputFolder, Path.GetFileName(document));

            var wordApp = new _Word.Application();
            try
            {
                LogMe($"*** Starting -> '{document}' ***");
                var doc = wordApp.Documents.Open(document, ReadOnly: false);
                var finder = wordApp.Selection.Find;

                finder.ClearFormatting();
                finder.Text = findText;
                finder.Replacement.Text = replaceText;
                finder.Replacement.ClearFormatting();
                finder.MatchCase = false;
                finder.Wrap = _Word.WdFindWrap.wdFindContinue;
                finder.MatchWholeWord = matchhw;
                finder.Format = false;
                finder.Forward = true;

                bool result = finder.Execute(Replace: _Word.WdReplace.wdReplaceAll);

                doc.SaveAs2(newDoc);
                doc.Close();

                if (result)
                    LogMe($"\t@@@ Found and Replaced. Saved at {newDoc} @@@");
                else LogMe("\t------ Nothing was found! ------");


                DispatchMe(() => { Mydocuments.Add(new MyDocument() { Document = Path.GetFileName(document), Result = result, NewDoc = Path.GetFileName(newDoc) }); });
            }
            catch (Exception err)
            {
                LogMe(err.Message);
                DispatchMe(() => { ShowStatus(" ?????????? error: Check log near document file! ???????????????"); });
            }
            finally
            {
                //SUPER IMPORTANT!
                //If you don't do this, each time you run the code 
                //the winword.exe process will keep running on background (for ever!),
                //at 10MB a piece, you may end up with a huge memory leak.
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp);
            }
        }
        private void DispatchMe(Action a)
        {
            Application.Current.Dispatcher.BeginInvoke(new Action(a));
        }
        private string GetUniqueID()
        {
            return (DateTime.Now - new DateTime(2021, 03, 19)).TotalSeconds.ToString().Substring(0, 7);
        }
        private PathType GetPathType(string src)
        {
            if (File.Exists(src))
                return PathType.File;
            else if (Directory.Exists(src))
            {
                var d = src.LastOrDefault();
                if (d != '\\') _path = src + @"\";
                return PathType.Directory;
            }
            return PathType.invalid;
        }
        private string CreateBackup(string doc)
        {
            var backup_name = Path.Combine(Path.GetDirectoryName(doc), $"{Path.GetFileNameWithoutExtension(doc)}_WorkDoc_{GetUniqueID()}{Path.GetExtension(doc)}");
            File.Copy(doc, backup_name);
            LogMe($"BackUp Created: {backup_name}");
            return backup_name;
        }
        private void SetLogPath(string path)
        {
            var reqDirectoryPath = Path.GetDirectoryName(path);

            //if not created or folder is changed
            if (string.IsNullOrWhiteSpace(_logPath) ||
                Path.GetDirectoryName(_logPath) != reqDirectoryPath)
            {
                _logPath = Path.Combine(reqDirectoryPath, $"WordRename_Log_{GetUniqueID()}.txt");
            }

        }
        private void LogMe(string msg)
        {
            //if (Application.Current.Resources["wrdrnLogger"] == null ||
            //    Path.GetDirectoryName((((StreamWriter)Application.Current.Resources["wrdrnLogger"]).BaseStream as FileStream).Name) != Path.GetDirectoryName(_logPath))
            //    Application.Current.Resources.Add("wrdrnLogger", new StreamWriter(_logPath, true, System.Text.Encoding.UTF8));

            //await ((StreamWriter)Application.Current.Resources["wrdrnLogger"]).WriteLineAsync($"{DateTime.Now} -- {msg}");
            ////            await _logger.
            //await ((StreamWriter)Application.Current.Resources["wrdrnLogger"]).FlushAsync();

            if (_logger == null ||
                Path.GetDirectoryName((_logger.BaseStream as FileStream).Name) != Path.GetDirectoryName(_logPath))
            {
                _logger = new StreamWriter(_logPath, true, System.Text.Encoding.UTF8);
                _logger.AutoFlush = true;
            }
            _logger.WriteLine($"{DateTime.Now} -- {msg}");
        }
        private void ShowStatus(string msg)
        {
            snackStatus.MessageQueue.Enqueue(msg, "OK", () => { snackStatus.IsActive = false; });

        }
        private void Working(bool arewe)
        {
            browsebtn.IsEnabled = donebtn.IsEnabled = !arewe;
        }

    }

    public class MyDocument
    {
        public string Document { get; set; }
        public bool Result { get; set; }
        public string NewDoc { get; set; }
    }

    internal enum PathType
    {
        File,
        Directory,
        invalid
    }
}
