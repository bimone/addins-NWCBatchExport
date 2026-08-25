using System.Threading;
using System.Globalization;
using System.Windows;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

using NWCBatchExporter.Views;
using NWCBatchExporter.Resources;


namespace NWCBatchExporter {
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            //If the host document is not saved
            if (doc == null || string.IsNullOrEmpty(doc.PathName)) {
                System.Windows.MessageBox.Show(Resource.MsgBoxInfo_ProjectMustBeSaved, Resource.MsgBoxTitle_ProjectNotSaved, MessageBoxButton.OK,MessageBoxImage.Warning);
                return Result.Cancelled;
            }
            else
            {
                if (OptionalFunctionalityUtils.IsNavisworksExporterAvailable() == false) {
                    var revitVersion = commandData.Application.Application.VersionNumber;
                    if (System.Windows.MessageBox.Show(string.Format(Resource.MsgBoxInfo_NeedInstallAutodesk, revitVersion),Resource.MsgBoxTitle_MissingUtility, MessageBoxButton.YesNo,MessageBoxImage.Asterisk) == MessageBoxResult.Yes)
                    {
                        // UseShellExecute is required on .NET to open a URL.
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                            GetExporterDownloadUrl(revitVersion)) { UseShellExecute = true });
                    }
                    return Result.Failed;
                }
                else
                {
                    NWCBatchExporterWindow dlg = new NWCBatchExporterWindow(doc);
                    dlg.ShowDialog();
                    return Result.Succeeded;
                }
            }
        }

        // Add a per-year direct link below when a stable one is known, otherwise the general
        // exporters page is used and the dialog tells the user which year to pick.
        private static string GetExporterDownloadUrl(string revitVersion)
        {
            switch (revitVersion)
            {
                default:
                    return "https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/How-to-update-Navisworks-Exporters-to-the-latest-current-version.html";
            }
        }
    }

}
