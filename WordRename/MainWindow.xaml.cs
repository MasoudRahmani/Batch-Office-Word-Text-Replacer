using System.Collections.ObjectModel;
using System.Windows;
using System.IO;
using _Word = Microsoft.Office.Interop.Word;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Diagnostics;

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
        private readonly string _appendString = "Renamed";
        private bool _dragged = false;
        private string _paths;
        private string _defaultDestination = string.Empty;
        public MainWindow()
        {
            Mydocuments = new ObservableCollection<MyDocument>();
            InitializeComponent();
            DataContext = this;
            LogMe("Start of Activity");
            Closed += MainWindow_Closed;
            Closing += Main_Closing;

            _defaultDestination = destBox.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "BatchWordRename");

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
            LogMe("Browsed...");

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".doc | .docx";
            dlg.Filter = "Word Document (.doc;.docx)|*.doc;*.docx";
            dlg.Multiselect = true;

            // Display OpenFileDialog by calling ShowDialog method
            var result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                srcbox.Text = string.Join(";\n", dlg.FileNames);
            }
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {
            _paths = srcbox.Text;
            string find = findbox.Text;
            string replace = replaceBox.Text;
            bool matchhw = (bool)matchHW.IsChecked;
            bool IsDestinationChecked = (bool)VarDestination.IsChecked;
            var destination = destBox.Text;

            if (IsDestinationChecked == false && string.IsNullOrEmpty(destBox.Text))
            {
                ShowStatus("Destination cannot be empty!"); return;
            }
            if (Directory.Exists(destBox.Text) == false && IsDestinationChecked == false)
            {
                if (destination == _defaultDestination)
                {
                    Directory.CreateDirectory(_defaultDestination);
                }
                else
                {
                    ShowStatus("Destination is wrong!"); return;
                }
            }

            if (string.IsNullOrEmpty(findbox.Text) | string.IsNullOrWhiteSpace(_paths))
            {
                ShowStatus("Find and Path cannot be empty!");
                return;
            }
            ShowStatus("Working...");
            Working(true);

            switch (GetPathType(_paths))
            {
                case PathType.File:
                    LogMe($"File Job: '{_paths}'");
                    LogMe($"^^^^^^^^^^^^^^^^^ find: '{find}' & replace: '{replace}' ^^^^^^^^^^^^^^^^^");
                    var FACT = new Action<string, bool, string>((string p, bool IsDC, string dest) =>
                    {
                        Find_Replace(p, find, replace, matchhw, _appendString, IsDC, dest);
                    });
                    Task.Factory.StartNew(() => { FACT(_paths, IsDestinationChecked, destination); })
                        .ContinueWith((s) => { DispatchMe(() => { Working(false); }); ShowStatus("Done."); }); ;
                    break;
                case PathType.Directory:
                    var DAct = new Action<string, bool, string>((string p, bool IsDC, string dest) =>
                    {
                        var path = p;
                        var files = Directory.EnumerateFiles(path).Where(f => f.ToLower().EndsWith("docx") || f.ToLower().EndsWith("doc")).ToList();
                        files.RemoveAll(x => { return Path.GetFileName(x).StartsWith("~$"); });

                        LogMe($"Directory Job for *{files.Count}* files: \r\n\r\n{string.Join("\r\n", files)}");
                        LogMe($"^^^^^^^^^^^^^^^^^ find: '{find}' & replace: '{replace}' ^^^^^^^^^^^^^^^^^");

                        files.ForEach(x =>
                        {
                            Find_Replace(x, find, replace, matchhw, _appendString, IsDC, dest);
                        });
                    });
                    Task.Factory.StartNew(() => DAct(_paths, IsDestinationChecked, destination))
                        .ContinueWith((s) => { DispatchMe(() => { Working(false); ShowStatus("Done."); }); });
                    break;

                case PathType.MultipleFile:
                    var MFAct = new Action<string, bool, string>((string p, bool IsDC, string dest) =>
                    {
                        var files = p.Split(";\n").ToList();

                        files.RemoveAll(x => { return Path.GetFileName(x).StartsWith("~$"); });

                        LogMe($"MultipleFile Job for *{files.Count}* files: \r\n\r\n{string.Join("\r\n", files)}");
                        LogMe($"^^^^^^^^^^^^^^^^^ find: '{find}' & replace: '{replace}' ^^^^^^^^^^^^^^^^^");

                        files.ForEach(x =>
                        {
                            Find_Replace(x, find, replace, matchhw, _appendString, IsDC, dest);
                        });
                    });
                    Task.Factory.StartNew(() => MFAct(_paths, IsDestinationChecked, destination))
                        .ContinueWith((s) => { DispatchMe(() => { Working(false); ShowStatus("Done."); }); });

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

        private void Find_Replace(string document, string findText, string replaceText, bool matchhw, string appendString, bool varDest, string dest)
        {
            bool SameFolder = varDest;
            string newDoc = string.Empty;
            var fileDirectory = Path.GetDirectoryName(document);
            var fileName_NoExt = Path.GetFileNameWithoutExtension(document);
            var ext = Path.GetExtension(document);

            if (SameFolder)
            {
                newDoc = Path.Combine(fileDirectory, string.Concat(fileName_NoExt, appendString, ext));
                if (File.Exists(newDoc))
                {
                    int index = 1;
                    newDoc = Path.Combine(fileDirectory, string.Concat(fileName_NoExt, appendString, index, ext));
                    while (File.Exists(newDoc))
                        newDoc = Path.Combine(fileDirectory, string.Concat(fileName_NoExt, appendString, ++index, ext));
                }
            }
            else
                newDoc = Path.Combine(dest, Path.GetFileName(document));


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

                if (result)
                {
                    doc.SaveAs2(newDoc);
                    LogMe($"\t@@@ Found and Replaced. Saved at {newDoc} @@@");
                }
                else
                {
                    newDoc = "**** XXXX Not Found XXX ****";
                    LogMe("\t------ Nothing was found! ------");
                }

                doc.Close();

                DispatchMe(() =>
                {
                    Mydocuments.Add(new MyDocument()
                    {
                        Document = Path.GetFileName(document),
                        Result = result,
                        NewDoc = Path.GetFileName(newDoc)
                    });
                });
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
        private static string GetUniqueID()
        {
            return (DateTime.Now - new DateTime(2021, 03, 19)).TotalSeconds.ToString().Substring(0, 7);
        }

        private PathType GetPathType(string src)
        {
            var filelist = src.Split(";\n");

            if (filelist.Length > 1)
                return PathType.MultipleFile;

            if (File.Exists(src))
                return PathType.File;
            else if (Directory.Exists(src))
            {
                var d = src.LastOrDefault();
                if (d != '\\') _paths = src + @"\";
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

            _logPath ??= Path.Combine(Environment.CurrentDirectory, $"wordrename-{DateTime.Now:yy-MM-dd}.txt");

            if (_logger == null ||
                Path.GetDirectoryName((_logger.BaseStream as FileStream).Name) != Path.GetDirectoryName(_logPath))
            {
                _logger = new StreamWriter(_logPath, true, System.Text.Encoding.UTF8) { AutoFlush = true };
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

        private void Srcbox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void Srcbox_Drop(object sender, DragEventArgs e)
        {
            string txt = string.Empty;
            int ok = 0, not = 0;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if ((file.EndsWith(".doc") || file.EndsWith(".docx")) && File.Exists(file))
                    {
                        txt += string.Concat(file, ";\n");
                        ok++;
                    }
                    else not++;
                }
            }
            srcbox.Text = txt.Substring(0, txt.Length - 2);
            LogMe($"Drag and Drop added: {ok} word file was added and {not} were not accepted. ");
            ShowStatus($"> Only {ok} document got accepted, other {not} file were not.");
            _dragged = true;
        }

        private void Srcbox_PreviewDragOVer(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void LogBtn_Clicked(object sender, RoutedEventArgs e)
        {
            var a = new Action<string>((string p) =>
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = p,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal,
                    UseShellExecute = true
                });
            });
            Task.Factory.StartNew(() => { a(_logPath); });
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
        invalid,
        MultipleFile
    }
}
