using System;
using System.Threading;
using VstoWord = Microsoft.Office.Interop.Word;
using VstoOutlook = Microsoft.Office.Interop.Outlook;

namespace InsertionImagesOffice.Office
{
    public sealed class Outlook
    {
        private static object _missing = Type.Missing;
        private static object _confirmConversion = false;
        private static object _link = false;
        private static object _attachment = false;
        public static void Start(string filePath)
        {
            VstoOutlook.Application outlookApp = new VstoOutlook.Application();

            outlookApp.Inspectors.NewInspector += inspector =>
            {
                if (inspector == null) return;
                dynamic currentItem = inspector.CurrentItem;

                if (currentItem == null) return;

                Thread t2 = new Thread(() =>
                {
                    VstoOutlook.MailItem item = currentItem as VstoOutlook.MailItem;
                    if (!currentItem.Sent)
                    {
                        Thread.Sleep(50);
                        try
                        {
                            VstoOutlook.Inspector currentInspector = item?.GetInspector;
                            VstoWord.Document wordEditor = currentInspector?.WordEditor as VstoWord.Document;
                            if (wordEditor == null) return;
                            wordEditor.Range(wordEditor.Range().End - 1)
                                .InsertFile(filePath, ref _missing, ref _confirmConversion, ref _link, ref _attachment);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                    }
                });
                t2.Start();
            };

            VstoOutlook.MailItem mailItem = (VstoOutlook.MailItem)
                outlookApp.CreateItem(VstoOutlook.OlItemType.olMailItem);

            mailItem.Display();
        }



    }
}
