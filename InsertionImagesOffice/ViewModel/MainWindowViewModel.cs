
using InsertionImagesOffice.Helpers;
using InsertionImagesOffice.Office;
using Microsoft.Win32;

namespace InsertionImagesOffice.ViewModel
{
    public class MainWindowViewModel : ObservableObject
    {
        #region Members
        private bool _buttonIsEnabled;
        private string _filePath;
        private RelayCommand<object> _openFileDialogCommand;
        private RelayCommand<object> _openWithWordCommand;
        private RelayCommand<object> _openWithOutlookCommand;
        #endregion

        #region Properties 
        public bool ButtonIsEnabled
        {
            get { return _buttonIsEnabled; }
            set
            {
                if (_buttonIsEnabled != value)
                {
                    _buttonIsEnabled = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string FilePath
        {
            get { return _filePath; }
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public RelayCommand<object> OpenFileDialogCommand
        {
            get { return _openFileDialogCommand; }
            set
            {
                if (_openFileDialogCommand != value)
                {
                    _openFileDialogCommand = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public RelayCommand<object> OpenWithWordCommand
        {
            get { return _openWithWordCommand; }
            set
            {
                if (_openWithWordCommand != value)
                {
                    _openWithWordCommand = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public RelayCommand<object> OpenWithOutlookCommand
        {
            get { return _openWithOutlookCommand; }
            set
            {
                if (_openWithOutlookCommand != value)
                {
                    _openWithOutlookCommand = value;
                    NotifyPropertyChanged();
                }
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Execute outlook with file selected
        /// </summary>
        /// <param name="args"></param>
        public void OpenWithOutlook(object args)
        {
            Outlook.Start(FilePath);
        }

        /// <summary>
        /// Execute word with file selected
        /// </summary>
        /// <param name="args"></param>
        public void OpenWithWord(object args)
        {
            Word.Start(FilePath);
        }

        /// <summary>
        /// Open a file dialog with a html file filter
        /// </summary>
        /// <param name="args"></param>
        public void OpenFileDialog(object args)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "HTML files|*.html;*.htm" };
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                ButtonIsEnabled = true;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Ctor
        /// </summary>
        public MainWindowViewModel()
        {
            ButtonIsEnabled = false;
            OpenFileDialogCommand = new RelayCommand<object>(OpenFileDialog);
            OpenWithOutlookCommand = new RelayCommand<object>(OpenWithOutlook);
            OpenWithWordCommand = new RelayCommand<object>(OpenWithWord);
        }
        #endregion

    }
}
