using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
namespace Outlook
{
    public class AttachmentViewModel : INotifyPropertyChanged
    {
        #region private
        Microsoft.Office.Interop.Outlook.Application application = null;
        NameSpace @namespace = null;
        private Thread t;
        #endregion
        public AttachmentViewModel()
        {
            try
            {
                IsBusy = Visibility.Hidden;
                OutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Attachment");
                // Check if there is an Outlook process running. 
                if (Process.GetProcessesByName("OUTLOOK").Length > 0)
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                    application = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                }
                else
                {
                    // If not, create a new instance of Outlook and log on to the default profile. 
                    application = new Microsoft.Office.Interop.Outlook.Application();
                    //nameSpace.Logon("", "", Missing.Value, Missing.Value);
                    //nameSpace = null;
                }
                @namespace = application.GetNamespace("MAPI");
                mapiFolder = @namespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            }
            catch (System.Exception ex)
            {
                Common.ShowMessage("发生异常", ex.Message);
            }
        }
        #region Property
        private MAPIFolder _mapiFolder;
        public MAPIFolder mapiFolder
        {
            get { return _mapiFolder; }
            set
            {
                if (null == value)
                    return;
                _mapiFolder = value;
                this.Total = value.Items.Count;
                if (Total == 0)
                    Common.ShowMessage("提示", "当前文件夹没有邮件，无需导出");
                OnPropertyChanged(() => this.mapiFolder);
                OnPropertyChanged(() => this.Folder);
                OnPropertyChanged(() => Message);
            }
        }
        [DefaultValue(true)]
        private bool _word=true;
        public bool Word
        {
            get { return _word; }
            set
            {
                _word = value;
                OnPropertyChanged(() => Word);
            }
        }
        private bool _excel=true;
        [DefaultValue(true)]
        public bool Excel
        {
            get
            { return _excel; }
            set
            {
                _excel = value;
                OnPropertyChanged(() => Excel);
            }
        }
        private bool _pdf=true;
        [DefaultValue(true)]
        public bool PDF
        {
            get { return _pdf; }
            set
            {
                _pdf = value;
                OnPropertyChanged(() => PDF);
            }
        }
        private bool _others=true;
        [DefaultValue(true)]
        public bool Others
        {
            get { return _others; }
            set
            {
                _others = value;
                OnPropertyChanged(() => Others);
            }
        }
        private string _outputPath;
        public string OutputPath
        {
            get { return _outputPath; }
            set
            {
                _outputPath = value;
                OnPropertyChanged(() => this.OutputPath);
            }
        }
        private Visibility _isBusy;
        public Visibility IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged(() => IsBusy);
            }
        }
        public string Folder
        {
            get
            {
                if (mapiFolder != null)
                    return mapiFolder.Name;
                else
                    return "收件箱";
            }
            set
            {
                OnPropertyChanged(() => this.Folder);
            }
        }
        public string Message
        {
            get
            {
                return string.Format("当前文件夹中邮件总数：{0}  正在处理第{1}条邮件  已经导出附件个数:{2}  ", Total, Current, Attachments);
            }
        }
        [DefaultValue(0)]
        public int Total { get; set; }
        [DefaultValue(0)]
        public int Current { get; set; }
        [DefaultValue(0)]
        public int Attachments { get; set; }
        #endregion
        #region Commands
        private ICommand _outputCommand;
        public ICommand OutputCommand
        {
            get
            {
                if (_outputCommand == null)
                    _outputCommand = new DelegateCommand(Output);
                return _outputCommand;
            }
        }
        private void Output()
        {
            if (!Word && !Excel && !PDF && !Others)
            {
                Common.ShowMessage("提示", "请选择导出附件类型");
                return;
            }
            try
            {
                t = new Thread(new ThreadStart(OutputAttachments));
                t.Start();
            }
            catch (System.Exception ex)
            {
                IsBusy = Visibility.Hidden;
                Common.ShowMessage("发生异常", ex.Message);
            }
        }
        private void OutputAttachments()
        {
            Current = 0;
            Attachments = 0;
            IsBusy = Visibility.Visible; 
            try
            {
                if (!Directory.Exists(OutputPath))
                    Directory.CreateDirectory(OutputPath);
                for (int i = 0; i < mapiFolder.Items.Count; i++)
                {               
                    ++Current;
                    MailItem mailItem = (MailItem)mapiFolder.Items[i+1];
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        if (Filter(attachment.FileName))
                        {
                            Console.WriteLine("已忽略 " + attachment.FileName);
                            continue;
                        }
                        else
                        {
                            Attachments++;
                            Console.WriteLine(Path.Combine(OutputPath, attachment.FileName));
                            attachment.SaveAsFile(Path.Combine(OutputPath,attachment.FileName));     
                            OnPropertyChanged(() => Message);
                        }
                    }
                OnPropertyChanged(() => Message);
                } 
                Common.ShowMessage("提示", "导出完成！");
                        
            }        
            catch (ThreadAbortException)
            { }
            catch (System.Exception ex)
            {
                Common.ShowMessage("发生异常", ex.Message);
            } 
        }
        private bool Filter(string filename)
        {
            string ext = Path.GetExtension(filename);
            if(!Word && (ext ==".doc" || ext == ".docx"))
                return true;
            else if(!Excel && (ext ==".xls" || ext == ".xlsx"))
                return true;
            else if(!PDF && ext ==".pdf")
                return true;
            else if(!Others)
                return true;
            return false;
        }
        private ICommand _selectPathCommand;
        public ICommand SelectPathCommand
        {
            get
            {
                if (_selectPathCommand == null)
                    _selectPathCommand = new DelegateCommand(SelectPath);
                return _selectPathCommand;
            }
        }
        private FolderBrowserDialog fbd;
        private void SelectPath()
        {
            if (fbd == null)
                fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                OutputPath = fbd.SelectedPath;
        }

        private ICommand _selectFolderCommand;
        public ICommand SelectFolderCommand
        {
            get
            {
                if (_selectFolderCommand == null)
                    _selectFolderCommand = new DelegateCommand(SelectFolder);
                return _selectFolderCommand;
            }
        }
        private void SelectFolder()
        {
            if (@namespace != null)
                mapiFolder = @namespace.PickFolder();
            else
            {
                Common.ShowMessage("异常", "获取Outlook文件夹失败");
            }
        }

        private ICommand _stopCommand;
        public ICommand StopCommand
        {
            get
            {
                if (_stopCommand == null)
                    _stopCommand = new DelegateCommand(Stop);
                return _stopCommand;
            }
        }
        private void Stop()
        {
            IsBusy = Visibility.Hidden;
            if (this.t != null && this.t.ThreadState != System.Threading.ThreadState.Aborted)
            {
                this.t.Abort();
            }
        }

        private ICommand _explorerCommand;
        public ICommand ExplorerCommand
        {
            get {
                if (_explorerCommand == null)
                    _explorerCommand = new DelegateCommand(Explorer);
                return _explorerCommand;
            }
        }
        private void Explorer()
        {
            if (!Directory.Exists(OutputPath))
                Directory.CreateDirectory(OutputPath);
            Process.Start(OutputPath);
        }
        #endregion

        #region INotifyPropertyChanged implementation

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged<TProperty>(Expression<Func<TProperty>> propertyExpresion)
        {
            var property = propertyExpresion.Body as MemberExpression;
            if (property == null || !(property.Member is PropertyInfo) ||
                !IsPropertyOfThis(property))
            {
                throw new ArgumentException(string.Format(
                    CultureInfo.CurrentCulture,
                    "Expression must be of the form 'this.PropertyName'. Invalid expression '{0}'.",
                    propertyExpresion), "propertyExpression");
            }

            this.OnPropertyChanged(property.Member.Name);
        }

        private bool IsPropertyOfThis(MemberExpression property)
        {
            var constant = RemoveCast(property.Expression) as ConstantExpression;
            return constant != null && constant.Value == this;
        }

        private System.Linq.Expressions.Expression RemoveCast(System.Linq.Expressions.Expression expression)
        {
            if (expression.NodeType == ExpressionType.Convert ||
                expression.NodeType == ExpressionType.ConvertChecked)
                return ((UnaryExpression)expression).Operand;

            return expression;
        }

        protected void OnPropertyChanged(
            string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName))
            {
                throw new ArgumentException("argument propertyName cannot by null oraz empty", "propertyName");
            }

            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler == null)
            {
                return;
            }

            handler(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
    public class Common
    {
        public static void ShowMessage(string title, string message)
        {
            System.Windows.MessageBox.Show(message, title);
        }
    }
    public class DelegateCommand : ICommand
    {
        System.Action action = null;
        public DelegateCommand(System.Action executeMethod)
        { action = executeMethod; }

        public bool CanExecute(object parameter)
        {
            return true;
            //throw new NotImplementedException();
        }
        public void Execute(object parameter)
        {
            if (action != null)
                action();
            //throw new NotImplementedException();
        }
        public event EventHandler CanExecuteChanged;
    }
}
