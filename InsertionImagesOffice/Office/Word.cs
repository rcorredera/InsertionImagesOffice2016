using System;
using System.Threading;
using VstoWord = Microsoft.Office.Interop.Word;

namespace InsertionImagesOffice.Office
{
    public sealed class Word
    {
        private static object _missing = Type.Missing;
        private static object _confirmConversion = false;
        private static object _link = false;
        private static object _attachment = false;
        public static void Start(string filePath)
        {
            VstoWord.Application wordApp = new VstoWord.Application
            {
                Visible = true
            };
            VstoWord.Document doc = wordApp.Documents.Add(ref _missing, ref _missing, ref _missing, true);

            Thread.Sleep(50);
            doc.Range(doc.Range().End - 1)
                .InsertFile(filePath, ref _missing, ref _confirmConversion, ref _link, ref _attachment);
        }
    }
}
