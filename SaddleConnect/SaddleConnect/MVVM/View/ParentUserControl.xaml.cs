using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Input;
using System.Diagnostics;
using System.Threading.Tasks;
using SaddleConnect;
using System.Data;
using System.Collections.ObjectModel;
using TIGUtility;

namespace SaddleConnect
{
    /// <summary>
    /// UI Events
    /// </summary>
    public partial class ParentUserControl : UserControl
    {
        public static ParentUserControl Instance;
        public System.Windows.Window _window = new System.Windows.Window();
        Document _doc = null;
        UIDocument _uidoc = null;
        string _offsetVariable = string.Empty;
        List<ExternalEvent> _externalEvents = new List<ExternalEvent>();
        List<string> _angleList = new List<string>() { "5.00", "11.25", "15.00", "22.50", "30.00", "45.00", "60.00", "90.00" };
        public ParentUserControl(List<ExternalEvent> externalEvents, CustomUIApplication application, Window window)
        {
            _uidoc = application.UIApplication.ActiveUIDocument;
            _doc = _uidoc.Document;
            _offsetVariable = application.OffsetVariable;
            _externalEvents = externalEvents;
            InitializeComponent();
            Instance = this;

            try
            {
                _window = window;
                LoadTab();
                ddlAngle.Attributes = new MultiSelectAttributes()
                {
                    Label = "Angle",
                    Width = 285
                };
                txtheight.UIApplication = application.UIApplication;
                List<MultiSelect> angleList = new List<MultiSelect>();
                foreach (string item in _angleList)
                    angleList.Add(new MultiSelect() { Name = item });
                ddlAngle.ItemsSource = angleList;
                ddlAngle.SelectedItem = angleList[4];
                string json = Utility.GetGlobalParametersManager(application.UIApplication, "ConPak-VConnect-gp");
                if (!string.IsNullOrEmpty(json))
                {
                    try
                    {
                        VerticalOffsetConnectGP globalParam = JsonConvert.DeserializeObject<VerticalOffsetConnectGP>(json);
                        ddlAngle.SelectedItem = angleList[angleList.FindIndex(x => x.Name == globalParam.AngleValue)];
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                _externalEvents[0].Raise();


            }
            catch (Exception exception)
            {

                System.Windows.MessageBox.Show("Some error has occured. \n" + exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void LoadTab()
        {
            List<CustomTab> MainTabsList = new List<CustomTab>();
            CustomTab a = new CustomTab();
            a.Id = 1;
            a.Name = "3 Pt Saddle";
            // a.Icon = PackIconKind.Pipe;
            MainTabsList.Add(a);
            a = new CustomTab();
            a.Id = 2;
            a.Name = "4 Pt Saddle";
            //a.Icon = PackIconKind.Four;
            MainTabsList.Add(a);
            shaddleControl.ItemsSource = MainTabsList;
            shaddleControl.SelectedIndex = 0;

            List<CustomTab> customTabsList = new List<CustomTab>();
            CustomTab b = new CustomTab();
            b.Id = 1;
            b.Icon = PackIconKind.ArrowTop;
            customTabsList.Add(b);
            b = new CustomTab();
            b.Id = 2;
            b.Icon = PackIconKind.ArrowDown;
            customTabsList.Add(b);
            b = new CustomTab();
            b.Id = 3;
            b.Icon = PackIconKind.ArrowExpand;
            customTabsList.Add(b);
            tagControl.ItemsSource = customTabsList;
            tagControl.SelectedIndex = 0;
        }


        private void Vconnect_BtnClick(object sender)
        {

        }

        private void tagControl_SelectionChanged(object sender)
        {
            View view = _uidoc.ActiveView;
            if (view != null && view.ViewType == ViewType.ThreeD)
            {
                if (tagControl.SelectedIndex == 2)
                {
                    Utility.AlertMessage("Tool requires PlanView to Connect", false, MainWindow.Instance.SnackbarSeven);
                    tagControl.SelectedIndex = 0;
                }

            }
            if (shaddleControl.SelectedIndex == 1)
            {
                if (tagControl.SelectedIndex == 2)
                {

                    txtheight.Visibility = System.Windows.Visibility.Collapsed;
                }
                else if (tagControl.SelectedIndex == 0 || tagControl.SelectedIndex == 1)
                {
                    txtheight.Visibility = System.Windows.Visibility.Visible;
                }
            }
            else
            {
                txtheight.Visibility = System.Windows.Visibility.Collapsed;
            }

        }

        private void shaddleControl_SelectionChanged(object sender)
        {
            if (shaddleControl.SelectedIndex == 1)
            {
                if (tagControl.SelectedIndex == 2)
                    txtheight.Visibility = System.Windows.Visibility.Collapsed;
                else
                    txtheight.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {

                txtheight.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            txtheight.Click_load(txtheight);
        }
    }
}

