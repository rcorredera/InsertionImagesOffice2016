
using InsertionImagesOffice.Helpers;
using InsertionImagesOffice.Office;
using Microsoft.Win32;

namespace InsertionImagesOffice.ViewModel
{
    public class MainWindowViewModel : ObservableObject
    {
        private bool _buttonIsEnabled;
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

        private string _filePath;
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

        private RelayCommand<object> _openFileDialogCommand;

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

        private RelayCommand<object> _openWithWordCommand;

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

        private RelayCommand<object> _openWithOutlookCommand;

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



        public void OpenWithOutlook(object args)
        {
            Outlook.Start(FilePath);
        }

        public void OpenWithWord(object args)
        {
            Word.Start(FilePath);
        }

        public void OpenFileDialog(object args)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "HTML files|*.html;*.htm";
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                ButtonIsEnabled = true;
            }
        }

        public MainWindowViewModel()
        {
            ButtonIsEnabled = false;
            OpenFileDialogCommand = new RelayCommand<object>(OpenFileDialog);
            OpenWithOutlookCommand = new RelayCommand<object>(OpenWithOutlook);
            OpenWithWordCommand = new RelayCommand<object>(OpenWithWord);
        }
    }
}
