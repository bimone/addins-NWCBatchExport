using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace NWCBatchExporter {
    [Transaction(TransactionMode.Manual)]
    public class Button : IExternalApplication {
        private const string TabLabel = "BIM One";
        public Result OnStartup(UIControlledApplication application) {
            var toolsPanel = GetOrCreateRibbonPanel(application);

            string assemblieFolder = Path.GetDirectoryName(Assembly.GetAssembly(GetType()).Location);
            string commandPath = Path.Combine(assemblieFolder, "NWCBatchExporter.dll");

            PushButton pushButton = toolsPanel.AddItem(new PushButtonData(
                                    "Batch Exporter",
                                    "Batch\nExporter",
                                    commandPath,
                                    "NWCBatchExporter.Command")) as PushButton;

            var buttonImage = Path.Combine(assemblieFolder, @"Resources\button.png");
            if (!File.Exists(buttonImage))
                buttonImage = Path.Combine(Directory.GetParent(assemblieFolder).FullName, @"Resources\button.png");

            pushButton.LargeImage = new BitmapImage(new Uri(buttonImage));
            pushButton.ToolTip = "Export your Revit 3D views in batch to Navisworks";
            pushButton.ToolTipImage = new BitmapImage(new Uri(buttonImage));

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application) {
            return Result.Succeeded;
        }

        private RibbonPanel GetOrCreateRibbonPanel(UIControlledApplication application)
        {
            var ribbonPanel = application.GetRibbonPanels(Tab.AddIns).Find(x => x.Name == TabLabel);
            if (ribbonPanel == null)
                ribbonPanel = application.CreateRibbonPanel(Tab.AddIns, TabLabel);

            return ribbonPanel;
        }
    }
}
