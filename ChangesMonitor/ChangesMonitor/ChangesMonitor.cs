//
// (C) Copyright 2003-2019 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using System;
using System.IO;
using System.Data;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB.Electrical;
using TIGUtility;
using Autodesk.Revit.DB.Events;
using System.Windows.Controls;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using ChangesMonitor;
using System.Security.Cryptography;
using System.Windows.Media.Imaging;
using Autodesk.Windows;
using System.Runtime.Remoting.Contexts;

namespace Revit.SDK.Samples.ChangesMonitor.CS
{

    /// <summary>
    /// A class inherits IExternalApplication interface and provide an entry of the sample.
    /// It create a modeless dialog to track the changes.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class ExternalApplication : IExternalApplication
    {

        List<Element> collection = null;

        #region  Class Member Variables
        /// <summary>
        /// A controlled application used to register the DocumentChanged event. Because all trigger points
        /// in this sample come from UI, the event must be registered to ControlledApplication. 
        /// If the trigger point is from API, user can register it to application 
        /// which can retrieve from ExternalCommand.
        /// </summary>
        private static ControlledApplication m_CtrlApp;

        /// <summary>
        /// data table for information windows.
        /// </summary>
        private static DataTable m_ChangesInfoTable;

        /// <summary>
        /// The window is used to show changes' information.
        /// </summary>
        private static ChangesInformationForm m_InfoForm;
        #endregion

        #region Class Static Property
        /// <summary>
        /// Property to get and set private member variables of changes log information.
        /// </summary>
        public static DataTable ChangesInfoTable
        {
            get { return m_ChangesInfoTable; }
            set { m_ChangesInfoTable = value; }
        }
        public static PushButton AutoConnectButton { get; set; }
        public static PushButton ToggleConPakToolsButton { get; set; }
        /// <summary>
        /// Property to get and set private member variables of info form.
        /// </summary>
        public static ChangesInformationForm InfoForm
        {
            get { return ExternalApplication.m_InfoForm; }
            set { ExternalApplication.m_InfoForm = value; }
        }
        #endregion

        #region IExternalApplication Members
        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit starts before a file or default template is actually loaded.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param> 
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully started. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If false is returned then Revit should inform the user that the external application 
        /// failed to load and the release the internal reference.</returns>
        public Result OnStartup(UIControlledApplication application)
        {
            /* UIDocument _uidoc = null;
             Document _doc = null;*/
            // initialize member variables.
            OnButtonCreate(application);

            m_CtrlApp = application.ControlledApplication;
            m_ChangesInfoTable = CreateChangeInfoTable();
            m_InfoForm = new ChangesInformationForm(ChangesInfoTable);
            // register the DocumentChanged event
            // m_CtrlApp.DocumentOpened += new EventHandler<Autodesk.Revit.DB.Events.DocumentOpenedEventArgs>(application_DocumentOpened);
            m_CtrlApp.DocumentChanged += new EventHandler<Autodesk.Revit.DB.Events.DocumentChangedEventArgs>(CtrlApp_DocumentChanged);

            // show dialog

            //m_InfoForm.Width = 300;
            // m_InfoForm.Show();
            m_InfoForm.Hide();
            // TaskDialog.Show("ChangesInfoTable", m_ChangesInfoTable.ToString());
            // Debug.Print(m_ChangesInfoTable.ToString());

            return Result.Succeeded;
        }

        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit is about to exit,Any documents must have been closed before this method is called.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param>
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully shutdown. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If false is returned then the Revit user should be warned of the failure of the external 
        /// application to shut down correctly.</returns>
        public Result OnShutdown(UIControlledApplication application)
        {
            m_CtrlApp.DocumentChanged += CtrlApp_DocumentChanged;
            m_InfoForm = null;
            m_ChangesInfoTable = null;
            return Result.Succeeded;
        }
        #endregion

        #region Event handler
        /// <summary>
        /// This method is the event handler, which will dump the change information to tracking dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        public static void Toggle()
        {
            string s = ToggleConPakToolsButton.ItemText;


            BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-off 32x32.png"));

            BitmapImage OnImage = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-on 16x16.png"));

            BitmapImage OnLargeImage = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-on 32x32.png"));

            BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-off 16x16.png"));

            if (s == "OFF")
            {
                ProjectParameterHandler projectParameterHandler = new ProjectParameterHandler();
                ExternalEvent Event = ExternalEvent.Create(projectParameterHandler);
                Event.Raise();
                ToggleConPakToolsButton.LargeImage = OnLargeImage;
                ToggleConPakToolsButton.Image = OnImage;

                Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
                foreach (Autodesk.Windows.RibbonTab tab in ribbon.Tabs)
                {
                    if (tab.Title.Equals("Sanveo Tools"))
                    {
                        foreach (Autodesk.Windows.RibbonPanel panel in tab.Panels)
                        {

                            if (panel.Source.AutomationName == "AutoUpdator")
                            {
                                RibbonItemCollection collctn = panel.Source.Items;


                                foreach (Autodesk.Windows.RibbonItem ri in collctn)
                                {
                                    if (ri is RibbonRowPanel)
                                    {
                                        foreach (var item in (ri as RibbonRowPanel).Items)
                                        {
                                            if (item is Autodesk.Windows.RibbonButton)
                                            {

                                                ri.IsVisible = false;
                                                ri.ShowText = false;
                                                ri.ShowImage = false;

                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                }

            }
            else
            {
                Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
                foreach (Autodesk.Windows.RibbonTab tab in ribbon.Tabs)
                {
                    if (tab.Title.Equals("Sanveo Tools"))
                    {
                        foreach (Autodesk.Windows.RibbonPanel panel in tab.Panels)
                        {

                            if (panel.Source.AutomationName == "AutoUpdator")
                            {
                                RibbonItemCollection collctn = panel.Source.Items;
                                foreach (Autodesk.Windows.RibbonItem ri in collctn)
                                {

                                    if (ri is RibbonRowPanel)
                                    {
                                        foreach (var item in (ri as RibbonRowPanel).Items)
                                        {
                                            if (item is Autodesk.Windows.RibbonButton)
                                            {

                                                ri.IsVisible = true;
                                                ri.ShowText = true;
                                                ri.ShowImage = true;

                                            }
                                        }
                                    }


                                }
                            }
                        }
                    }
                }
                ToggleConPakToolsButton.LargeImage = OffLargeImage;
                ToggleConPakToolsButton.Image = OffImage;
            }


            ToggleConPakToolsButton.ItemText = s.Equals("OFF") ? "ON" : "OFF";


        }
        private void OnButtonCreate(UIControlledApplication application)
        {
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string dllLocation = Path.Combine(executableLocation, "ChangesMonitor.dll");

            PushButtonData buttondata = new PushButtonData("ModifierBtn", "OFF", dllLocation, "Revit.SDK.Samples.ChangesMonitor.CS.Command");
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-off 32x32.png"));
            buttondata.LargeImage = pb1Image;

            BitmapImage pb1Image2 = new BitmapImage(new Uri("pack://application:,,,/ChangesMonitor;component/Resources/switch-off 16x16.png"));
            buttondata.Image = pb1Image2;
            buttondata.AvailabilityClassName = "Revit.SDK.Samples.ChangesMonitor.CS.Availability";
            var ribbonPanel = RibbonPanel(application);
            if (ribbonPanel != null)
                ToggleConPakToolsButton = ribbonPanel.AddItem(buttondata) as PushButton;


        }

        public Autodesk.Revit.UI.RibbonPanel RibbonPanel(UIControlledApplication a)
        {
            string tab = "Sanveo Tools"; // Archcorp
            string ribbonPanelText = "AutoUpdator"; // Architecture

            // Empty ribbon panel 
            Autodesk.Revit.UI.RibbonPanel ribbonPanel = null;
            // Try to create ribbon tab. 
            try
            {
                a.CreateRibbonTab(tab);
            }
            catch { }
            // Try to create ribbon panel.
            try
            {
                Autodesk.Revit.UI.RibbonPanel panel = a.CreateRibbonPanel(tab, ribbonPanelText);
            }
            catch { }
            // Search existing tab for your panel.
            List<Autodesk.Revit.UI.RibbonPanel> panels = a.GetRibbonPanels(tab);
            foreach (Autodesk.Revit.UI.RibbonPanel p in panels)
            {
                if (p.Name == ribbonPanelText)
                {
                    ribbonPanel = p;
                }
            }
            //return panel 
            return ribbonPanel;
        }
        void CtrlApp_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            if (ToggleConPakToolsButton.ItemText == "ON")
            {

                // get the current document.
                Document doc = e.GetDocument();
                ICollection<ElementId> modifiedElem = e.GetModifiedElementIds();
                try
                {
                    List<Element> elementlist = new List<Element>();
                    List<ElementId> rvConduitlist = new List<ElementId>();
                    string value = string.Empty;
                    foreach (ElementId id in modifiedElem)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem.Category != null && elem.Category.Name == "Conduits")
                        {
                            Parameter parameter = elem.LookupParameter("Bend Angle");
                            value = parameter.AsString();
                            elementlist.Add(elem);
                        }

                    }
                    ChangesInformationForm.instance.MidSaddlePt = elementlist.Distinct().ToList();
                    ChangesInformationForm.instance._elemIdone.Clear();
                    ChangesInformationForm.instance._elemIdtwo.Clear();
                    List<ElementId> FittingElem = new List<ElementId>();
                    for (int i = 0; i < elementlist.Count; i++)
                    {
                        ConnectorSet connector = GetConnectorSet(elementlist[i]);
                        List<ElementId> Icollect = new List<ElementId>();
                        foreach (Connector connect in connector)
                        {
                            ConnectorSet cs1 = connect.AllRefs;
                            foreach (Connector c in cs1)
                            {
                                Icollect.Add(c.Owner.Id);
                            }
                            foreach (ElementId eid in Icollect)
                            {
                                if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null &&
                                    doc.GetElement(eid).Category.Name == "Conduit Fittings"))
                                {
                                    FittingElem.Add(eid);
                                }
                            }
                        }
                    }

                    List<ElementId> FittingElements = new List<ElementId>();

                    FittingElements = FittingElem.Distinct().ToList();
                    List<Element> BendElements = new List<Element>();
                    foreach (ElementId id in FittingElements)
                    {
                        BendElements.Add(doc.GetElement(id));


                    }
                    List<ElementId> Icollector = new List<ElementId>();


                    for (int i = 0; i < BendElements.Count; i++)
                    {
                        ConnectorSet connector = GetConnectorSet(BendElements[i]);
                        foreach (Connector connect in connector)
                        {
                            ConnectorSet cs1 = connect.AllRefs;
                            foreach (Connector c in cs1)
                            {
                                Icollector.Add(c.Owner.Id);
                            }
                        }
                    }
          
                    foreach (ElementId eid in Icollector)
                    {
                        if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null &&
                            doc.GetElement(eid).Category.Name == "Conduits"))
                        {
                            ChangesInformationForm.instance._selectedElements.Add(eid);
                        }
                    }
                    List<Element> elementtwo = new List<Element>();
                    List<ElementId> RefID = new List<ElementId>();
                    for (int i = 0; i < BendElements.Count; i++)
                    {
                        for (int j = i + 1; j < BendElements.Count; j++)
                        {
                            Element elemOne = BendElements[i];
                            Element elemTwo = BendElements[j];
                            Parameter parameter = elemOne.LookupParameter("Angle");
                            if (parameter.AsValueString() == "90.00°")
                            {
                                ConnectorSet firstconnector = GetConnectorSet(elemOne);
                                foreach (Connector connector in firstconnector)
                                {
                                    ConnectorSet cs1 = connector.AllRefs;
                                    foreach (Connector c in cs1)
                                    {
                                        RefID.Add(c.Owner.Id);
                                    }

                                }
                            }
                            ChangesInformationForm.instance._refConduitKick.AddRange(RefID);
                            ChangesInformationForm.instance._Value = value;
                            if (elemOne != null)
                            {
                                ConnectorSet firstconnector = GetConnectorSet(elemOne);
                                ConnectorSet secondconnector = GetConnectorSet(elemTwo);
                                try
                                {
                                    List<ElementId> IDone = new List<ElementId>();
                                    foreach (Connector connector in firstconnector)
                                    {
                                        ConnectorSet cs1 = connector.AllRefs;
                                        foreach (Connector c in cs1)
                                        {
                                            IDone.Add(c.Owner.Id);
                                        }
                                        foreach (ElementId eid in IDone)
                                        {
                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduits"))
                                            {
                                                ChangesInformationForm.instance._elemIdone.Add(eid);
                                            }
                                        }
                                    }


                                    List<ElementId> IDtwo = new List<ElementId>();
                                    foreach (Connector connector in secondconnector)
                                    {
                                        ConnectorSet cs1 = connector.AllRefs;
                                        foreach (Connector c in cs1)
                                        {
                                            IDtwo.Add(c.Owner.Id);
                                        }
                                        foreach (ElementId eid in IDtwo)
                                        {
                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduits"))
                                            {
                                                ChangesInformationForm.instance._elemIdtwo.Add(eid);
                                                if (ChangesInformationForm.instance._elemIdone.Any(r => r == eid))
                                                {
                                                    ChangesInformationForm.instance._deletedIds.Add(eid);
                                                    rvConduitlist.Add(eid);

                                                }
                                            }
                                        }
                                    }
                                    ChangesInformationForm.instance._deletedIds.Add(elemOne.Id);
                                    ChangesInformationForm.instance._deletedIds.Add(elemTwo.Id);
                                }
                                catch
                                {

                                }

                            }
                        }
                    }
                    try
                    {
                        var l = rvConduitlist.Distinct();
                        ChangesInformationForm.instance._selectedElements = ChangesInformationForm.instance._selectedElements.Except(l).ToList();
                        AngleDrawHandler handler = new AngleDrawHandler();
                        ExternalEvent DrawEvent = ExternalEvent.Create(handler);
                        DrawEvent.Raise();

                    }
                    catch
                    {
                        MessageBox.Show("Error");
                    }
                }
                catch { }
            }

        }
        #endregion

        #region Class Methods
        /// <summary>
        /// This method is used to retrieve the changed element and add row to data table.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="doc"></param>
        /// <param name="changeType"></param>
        public void AddChangeInfoRow(ElementId id, Document doc, string changeType)
        {
            // retrieve the changed element
            Element elem = doc.GetElement(id);

            MessageBox.Show(elem.Id.ToString());
            DataRow newRow = m_ChangesInfoTable.NewRow();
            if (elem != null)
            {

                Element primaryelement = null;
                ConnectorSet firstconnector = Utility.GetConnectorSet(elem);
                //ConnectorSet secondconnector = Utility.GetConnectorSet(elem);
                try
                {
                    foreach (Connector connector in firstconnector)
                    {
                        primaryelement = connector.Owner;
                        MessageBox.Show(primaryelement.Id.ToString());
                    }
                }
                catch
                {
                    MessageBox.Show("error");
                }
                Parameter parameter = elem.LookupParameter("Bend Angle");
                string value = parameter.AsString();
                // Parameter p = elem.get_Parameter(BuiltInParameter.CONDUIT_STANDARD_TYPE_PARAM);
                // Pname = p.Definition.Name;
                // string changes = string.Empty;
                List<string> list = new List<string>();
                /* Element elemtcollec = new FilteredElementCollector(doc).OfClass(typeof(Conduit)).FirstOrDefault();
                 try
                 {
                     foreach (var p in elem.GetOrderedParameters().Where(x => !x.IsReadOnly).ToList())
                     {
                         Parameter parameter1 = elem.LookupParameter(p.Definition.Name);
                         if (parameter1.AsString() != value)
                         {
                             //list.Add(p.Definition.Name);
                             //changes = parameter1.AsString();
                         }
                         else
                         {

                         }
                     }
                 }
                 catch
                 {
                     return;
                 }*/
                if (elem.Category.Name == "Center line")
                {
                    //MessageBox.Show("Category : " + elem.Category.Name + " \n" + "ID :" + id.ToString() + " in " + string.Join(",", list));

                }
                else
                {
                    if (value != null)
                        MessageBox.Show("Category : " + elem.Category.Name + " \n" + "ID : " + id.ToString() + "\n" + " Bend Angle : " + value);
                    //list.Clear();
                }
                // return value;
            }

        }

        public static ConnectorSet GetConnectorSet(Autodesk.Revit.DB.Element Ele)
        {
            ConnectorSet result = null;
            if (Ele is Autodesk.Revit.DB.FamilyInstance)
            {
                MEPModel mEPModel = ((Autodesk.Revit.DB.FamilyInstance)Ele).MEPModel;
                if (mEPModel != null && mEPModel.ConnectorManager != null)
                {
                    result = mEPModel.ConnectorManager.Connectors;
                }
            }
            else if (Ele is MEPCurve)
            {
                result = ((MEPCurve)Ele).ConnectorManager.Connectors;
            }

            return result;
        }

        /// <summary>
        /// Generate a data table with five columns for display in window
        /// </summary>
        /// <returns>The DataTable to be displayed in window</returns>
        private DataTable CreateChangeInfoTable()
        {
            // create a new dataTable
            DataTable changesInfoTable = new DataTable("ChangesInfoTable");

            // Create a "ChangeType" column. It will be "Added", "Deleted" and "Modified".
            DataColumn styleColumn = new DataColumn("ChangeType", typeof(System.String));
            styleColumn.Caption = "ChangeType";
            changesInfoTable.Columns.Add(styleColumn);

            // Create a "Id" column. It will be the Element ID
            DataColumn idColum = new DataColumn("Id", typeof(System.String));
            idColum.Caption = "Id";
            changesInfoTable.Columns.Add(idColum);

            // Create a "Name" column. It will be the Element Name
            DataColumn nameColum = new DataColumn("Name", typeof(System.String));
            nameColum.Caption = "Name";
            changesInfoTable.Columns.Add(nameColum);

            // Create a "Category" column. It will be the Category Name of the element.
            DataColumn categoryColum = new DataColumn("Category", typeof(System.String));
            categoryColum.Caption = "Category";
            changesInfoTable.Columns.Add(categoryColum);

            // Create a "Document" column. It will be the document which own the changed element.
            DataColumn docColum = new DataColumn("Document", typeof(System.String));
            docColum.Caption = "Document";
            changesInfoTable.Columns.Add(docColum);

            // return this data table 
            return changesInfoTable;
        }
        #endregion
    }

    /// <summary>
    /// This class inherits IExternalCommand interface and used to retrieve the dialog again.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {

        public List<ElementId> _deletedIds = new List<ElementId>();
        #region IExternalCommand Members
        /// <summary>
        /// Implement this method as an external command for Revit.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application
        /// which contains data related to the command,
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application
        /// which will be displayed if a failure or cancellation is returned by
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command.
        /// A result of Succeeded means that the API external method functioned as expected.
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with
        /// the operation.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            /*if (ExternalApplication.InfoForm == null)
            {
                ExternalApplication.InfoForm = new ChangesInformationForm(ExternalApplication.ChangesInfoTable);
            }
            ExternalApplication.InfoForm.Show();*/
            ExternalApplication.Toggle();

            UIDocument uIDocument = commandData.Application.ActiveUIDocument;
            Document doc = uIDocument.Document;
            if (doc.IsReadOnly)
            {
                MessageBox.Show("doc is read Only");
            }

            return Result.Succeeded;
        }
        #endregion
    }

}
