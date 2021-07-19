using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using NWCBatchExporter.Resources;

namespace NWCBatchExporter.Views {
    /// <summary>
    /// Interaction logic for ExportOptionsWindow.xaml
    /// </summary>
    /// 
    public partial class ExportOptionsWindow : Window {
        public NavisworksExportOptions _neo = null, dneo;
        public ExportOptionsWindow(NavisworksExportOptions default_neo) {
            InitializeComponent();
            _neo = default_neo;
            /*cbConvertElParameters.Items.Add(NavisworksParameters.All);
            cbConvertElParameters.Items.Add(NavisworksParameters.Elements);
            cbConvertElParameters.Items.Add(NavisworksParameters.None);*/
            cbConvertElParameters.Items.Add(Resource.FormLst_All);
            cbConvertElParameters.Items.Add(Resource.FormLst_Elements);
            cbConvertElParameters.Items.Add(Resource.FormLst_None);

            /*cbCoordinates.Items.Add(NavisworksCoordinates.Internal);
            cbCoordinates.Items.Add(NavisworksCoordinates.Shared);*/
            cbCoordinates.Items.Add(Resource.FormLst_Internal);
            cbCoordinates.Items.Add(Resource.FormLst_Shared);

            /*cbExport.Items.Add(NavisworksExportScope.Model);
            cbExport.Items.Add(NavisworksExportScope.SelectedElements);
            cbExport.Items.Add(NavisworksExportScope.View);*/
            cbExport.Items.Add(Resource.FormLst_Model);
            cbExport.Items.Add(Resource.FormLst_SelectedElements);
            cbExport.Items.Add(Resource.FormLst_View);

            SetDefaultsValues(_neo);
        }

        private void SetDefaultsValues(NavisworksExportOptions default_neo) {
            default_neo.FindMissingMaterials = true;
            chbxConstructionParts.IsChecked = default_neo.ExportParts;
            chbxElementIds.IsChecked = default_neo.ExportElementIds;
            chbxElementProperies.IsChecked = default_neo.ConvertElementProperties;
            chbxLinkedFiles.IsChecked = default_neo.ExportLinks;
            chbxRoomAttr.IsChecked = default_neo.ExportRoomAsAttribute;
            chbxConvertURL.IsChecked = default_neo.ExportUrls;
            chbxDivideFiles.IsChecked = default_neo.DivideFileIntoLevels;
            chbxExportGeometry.IsChecked = default_neo.ExportRoomGeometry;
            chbxMissingMaterials.IsChecked = default_neo.FindMissingMaterials;
            switch (default_neo.Parameters) {
                case NavisworksParameters.All:
                    cbConvertElParameters.SelectedIndex = 0;
                    break;

                case NavisworksParameters.Elements:
                    cbConvertElParameters.SelectedIndex = 1;
                    break;

                case NavisworksParameters.None:
                    cbConvertElParameters.SelectedIndex = 2;
                    break;
            }

            switch (default_neo.Coordinates) {
                case NavisworksCoordinates.Internal:
                    cbCoordinates.SelectedIndex = 0;
                    break;

                case NavisworksCoordinates.Shared:
                    cbCoordinates.SelectedIndex = 1;
                    break;
            }
            switch (default_neo.ExportScope) {
                case NavisworksExportScope.Model:
                    cbExport.SelectedIndex = 0;
                    break;

                case NavisworksExportScope.SelectedElements:
                    cbExport.SelectedIndex = 1;
                    break;

                case NavisworksExportScope.View:
                    cbExport.SelectedIndex = 2;
                    break;
            }
        }
        private void chbxConstructionParts_Click(object sender, RoutedEventArgs e) {
            _neo.ExportParts = (bool)chbxConstructionParts.IsChecked;
        }
        private void chbxElementIds_Click(object sender, RoutedEventArgs e) {
            _neo.ExportElementIds = (bool)chbxElementIds.IsChecked;
        }
        private void chbxElementProperies_Click(object sender, RoutedEventArgs e) {
            _neo.ConvertElementProperties = (bool)chbxElementProperies.IsChecked;
        }
        private void chbxLinkedFiles_Click(object sender, RoutedEventArgs e) {
            _neo.ExportLinks = (bool)chbxLinkedFiles.IsChecked;
        }
        private void chbxRoomAttr_Click(object sender, RoutedEventArgs e) {
            _neo.ExportRoomAsAttribute = (bool)chbxRoomAttr.IsChecked;
        }
        private void chbxConvertURL_Click(object sender, RoutedEventArgs e) {
            _neo.ExportUrls = (bool)chbxConvertURL.IsChecked;
        }
        private void chbxDivideFiles_Click(object sender, RoutedEventArgs e) {
            _neo.DivideFileIntoLevels = (bool)chbxDivideFiles.IsChecked;
        }
        private void chbxExportGeometry_Click(object sender, RoutedEventArgs e) {
            _neo.ExportRoomGeometry = (bool)chbxExportGeometry.IsChecked;
        }
        private void chbxMissingMaterials_Click(object sender, RoutedEventArgs e) {
            _neo.FindMissingMaterials = (bool)chbxMissingMaterials.IsChecked;
        }
        private void btnDefault_Click(object sender, RoutedEventArgs e) {
            NavisworksExportOptions default_neo = new NavisworksExportOptions();
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
            SetDefaultsValues(default_neo);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e) {
            this.Close();
        }

        private void cbConvertElParameters_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            switch (cbConvertElParameters.SelectedIndex) { 
                case 0: _neo.Parameters = NavisworksParameters.All;
                    break;
                case 1: _neo.Parameters = NavisworksParameters.Elements;
                    break;
                case 2: _neo.Parameters = NavisworksParameters.None;
                    break;
            }
        }

        private void cbCoordinates_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            switch (cbCoordinates.SelectedIndex) {
                case 0: _neo.Coordinates = NavisworksCoordinates.Internal;
                    break;
                case 1: _neo.Coordinates = NavisworksCoordinates.Shared;
                    break;
            }
        }

        private void cbExport_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            switch (cbExport.SelectedIndex) {
                case 0: _neo.ExportScope = NavisworksExportScope.Model;
                    break;
                case 1: _neo.ExportScope = NavisworksExportScope.SelectedElements;
                    break;
                case 2: _neo.ExportScope = NavisworksExportScope.View;
                    break;
            }
        }


    }
}
