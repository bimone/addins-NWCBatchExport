using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using System.Globalization;
using System.Threading;
using NWCBatchExporter.Resources;
using NWCBatchExporter.ViewModels;

namespace NWCBatchExporter.Views {
    /// <summary>
    /// Interaction logic for NWCBatchExporterWindow.xaml
    /// </summary>
    public partial class NWCBatchExporterWindow : Window {
        Document _doc = null;

        public class View3DData {
            public string Name { get; set; }
            public ElementId Id  { get; set; }
            public bool Selected { get; set; }
        }

        public class ViewsData {
            public List<View3DData> DataList { get; set; }
            public string Path { get; set; }
            public NavisworksExportOptions NEO { get; set; }
            public Document doc { get; set; }
        }
        private readonly BackgroundWorker worker = new BackgroundWorker();
        private NavisworksExportOptions default_neo = new NavisworksExportOptions();
        public NWCBatchExporterWindow(Document doc) {
            //====================================We use the same language as Revit=========================================
            Autodesk.Revit.ApplicationServices.Application application = doc.Application;
            LanguageType lang = application.Language;
            if (lang.ToString().Contains("French"))
            {
                var cultureInfo = new CultureInfo("fr-FR");
                Thread.CurrentThread.CurrentCulture = Thread.CurrentThread.CurrentUICulture = cultureInfo;
            }
            //==============================================================================================================
            _doc = doc;
            InitializeComponent();
            PopulateViews3DList(_doc);
            var mainViewModel = new MainViewModel();
            DataContext = mainViewModel;

            default_neo.ExportParts = false;
            default_neo.ExportElementIds = true;
            default_neo.Parameters = NavisworksParameters.All;
            default_neo.ConvertElementProperties = false;
            default_neo.ExportLinks = true;
            default_neo.ExportRoomAsAttribute = true;
            default_neo.ExportUrls = true;
            default_neo.Coordinates = NavisworksCoordinates.Internal;
            default_neo.DivideFileIntoLevels = false;
            default_neo.ExportScope = NavisworksExportScope.View;
            default_neo.ExportRoomGeometry = true;
            default_neo.FindMissingMaterials = true;

            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;

            Title = "Batch Exporter " + Assembly.GetAssembly(GetType()).GetName().Version.ToString();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e) {
            using (System.Windows.Forms.FolderBrowserDialog fd = new System.Windows.Forms.FolderBrowserDialog()) {
                System.Windows.Forms.DialogResult result = fd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK) 
                    tbFolderPath.Text = fd.SelectedPath;
            }
        }

        private void PopulateViews3DList(Document doc) {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(View3D));
            List<View3DData> dt = new List<View3DData>();
            //Code de départ - 9sept2019 
            /*foreach (View3D v3d in collector)
                if ((!dt.Any(x => x.Name == v3d.ViewName)) & (!v3d.IsTemplate) & (v3d.IsValidObject))
                    dt.Add(new View3DData { Name = v3d.ViewName, Id = v3d.Id, Selected = false });*/
            //Code ajouter par MP le 9 sept2019
            foreach (View3D v3d in collector)
            {
             
                    if ((!dt.Any(x => x.Name == v3d.Name)) & (!v3d.IsTemplate) & (v3d.IsValidObject))
                        dt.Add(new View3DData { Name = v3d.Name, Id = v3d.Id, Selected = false });
            }
            dt = dt.OrderBy(x=>x.Name).ToList();
            lbViews3D.ItemsSource = dt;
        }

        private void btnCheckAll_Click(object sender, RoutedEventArgs e) {
            List<View3DData> dt = lbViews3D.ItemsSource as List<View3DData>;
            for (int i = 0; i < dt.Count; ++i) 
                dt[i].Selected = true;
            lbViews3D.ItemsSource = null;
            lbViews3D.ItemsSource = dt;
        }

        private void btnCheckNone_Click(object sender, RoutedEventArgs e) {
            List<View3DData> dt = lbViews3D.ItemsSource as List<View3DData>;
            for (int i = 0; i < dt.Count; ++i)
                dt[i].Selected = false;
            lbViews3D.ItemsSource = null;
            lbViews3D.ItemsSource = dt;
        }

        private NavisworksExportOptions neo = null;
        private void btnExportOptions_Click(object sender, RoutedEventArgs e) {
            if (neo == null)
                neo = default_neo;
            NWCBatchExporter.Views.ExportOptionsWindow oew = new ExportOptionsWindow(neo);
            oew.Owner = this;
            oew.ShowDialog();
            neo = oew._neo;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            List<View3DData> dt = lbViews3D.ItemsSource as List<View3DData>;
            if (neo == null)
                neo = default_neo;
            int s = 0, f = 0;
            string fViewNames = "";
            var canceled = false;
            if ((tbFolderPath.Text.Length != 0) & (System.IO.Directory.Exists(tbFolderPath.Text)))
            {
                foreach (View3DData v3d in dt)
                {
                    if (v3d.Selected)
                    {
                        neo.ViewId = v3d.Id;
                        try
                        {
                            var fileName = NormalizeFileName(v3d.Name);
                            _doc.Export(tbFolderPath.Text, fileName + ".nwc", neo);
                            if (!File.Exists(tbFolderPath.Text + "\\" + fileName + ".nwc"))
                            {
                                canceled = true;
                                MessageBox.Show(Resource.MsgBoxInfo_ExportProcessCanceledByTheUser, Resource.MsgBoxTitle_ExportResults, MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                            }
                            s++;
                        }
                        catch (SystemException)
                        {
                            f++;
                            fViewNames += v3d.Name + ",";
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(Resource.MsgBoxInfo_SelectPathFirst, Resource.MsgBoxTitle_Error, MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!canceled)
            {
                if (fViewNames.Length != 0)
                {
                    //"Sucessful : " + s.ToString() + " view(s), Failed:" + f.ToString() + "view(s) - " + fViewNames.Remove(fViewNames.Length - 1)
                    MessageBox.Show(string.Format(Resource.MsgBoxInfo_SucessfulFailed, s.ToString(), f.ToString(), fViewNames.Remove(fViewNames.Length - 1)), Resource.MsgBoxTitle_ExportResults, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    //"Sucessful : " + s.ToString() + " view(s), Failed: 0 views",
                    MessageBox.Show(string.Format(Resource.MsgBoxInfo_SucessfulXFailed0, s.ToString()), Resource.MsgBoxTitle_ExportResults, MessageBoxButton.OK, MessageBoxImage.Information);
                }

            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e) {
           
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

        }

        private void lbViews3D_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            List<View3DData> dt = lbViews3D.ItemsSource as List<View3DData>;
            if(dt == null)
            {
                return;
            }

            int count = 0;
            foreach (View3DData v3d in dt) {
                if (v3d.Selected)
                    count++;
            }
            if (count != 0)
                lbCount.Content = count.ToString() + " " + Resource.FormLbl_ViewSelected;
            else
                lbCount.Content = "";
        }

        private void btnClose_Click(object sender, RoutedEventArgs e) {
            this.Close();
        }

        private string NormalizeFileName(string fileName)
        {
            return fileName.Replace("{", "").Replace("}", "");
        }
      
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
