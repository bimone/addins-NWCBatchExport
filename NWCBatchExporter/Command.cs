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
                    if (System.Windows.MessageBox.Show(Resource.MsgBoxInfo_NeedInstallAutodesk,Resource.MsgBoxTitle_MissingUtility, MessageBoxButton.YesNo,MessageBoxImage.Asterisk) == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start("http://www.autodesk.com/products/navisworks/autodesk-navisworks-nwc-export-utility");
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
    }

}
