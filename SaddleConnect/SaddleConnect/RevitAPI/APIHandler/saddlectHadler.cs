using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Newtonsoft.Json;
using Application = Autodesk.Revit.ApplicationServices.Application;
using TIGUtility;
using Autodesk.Revit.DB.Analysis;
using static System.Windows.Forms.AxHost;
using System.Windows.Media.Media3D;
using MaterialDesignThemes.Wpf;

namespace SaddleConnect
{
    [Transaction(TransactionMode.Manual)]
    public class saddlectHadler : IExternalEventHandler
    {

        public List<Element> _deleteElements = new List<Element>();
        List<Element> SelectedElements = new List<Element>();
        List<Element> DistanceElements = new List<Element>();
        bool _isfirst;
        DateTime startDate = DateTime.UtcNow;
        UIApplication _uiapp = null;
        XYZ pickpoint;
        public void Execute(UIApplication uiapp)
        {
            _uiapp = uiapp;
            UIDocument uidoc = _uiapp.ActiveUIDocument;
            Application app = _uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(_uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";

            try
            {


                for (int ii = 0; ii < 100; ii++)
                {
                    _isfirst = false;
                    if (SelectedElements.Count <= 0)
                    {
                        SelectedElements = Utility.GetPickedElements(uidoc, "Please select the conduits", typeof(Conduit), true);
                        if (SelectedElements == null)
                        {
                            return;
                        }
                        if (SelectedElements.Count() == 0)
                        {
                            System.Windows.MessageBox.Show("Please select the conduits alone to perform action", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                            uidoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                            continue;
                        }
                    }
                    if (SelectedElements.Count() != 0)
                    {
                        List<Element> conduitCollection = new List<Element>();

                        if (ParentUserControl.Instance.ddlAngle.SelectedItem == null || string.IsNullOrEmpty(ParentUserControl.Instance.ddlAngle.SelectedItem.Name))
                        {
                            return;
                        }

                        VerticalOffsetConnectGP globalParam = new VerticalOffsetConnectGP
                        {
                            AngleValue = ParentUserControl.Instance.ddlAngle.SelectedItem == null ? "30.00" : ParentUserControl.Instance.ddlAngle.SelectedItem.Name
                        };

                        using (Transaction tr = new Transaction(doc, "ConPak-VConnect"))
                        {
                            tr.Start();

                            var CongridDictionary1 = Utility.GroupByElements(SelectedElements);

                            Dictionary<double, List<Element>> group = new Dictionary<double, List<Element>>();
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                pickpoint = Utility.PickPoint(uidoc);

                            if (CongridDictionary1.Count == 2)
                            {
                                Dictionary<double, List<Element>> groupPrimary = Utility.GroupByElementsWithElevation(CongridDictionary1.First().Value.Select(x => x.Conduit).ToList(), offsetVariable);
                                Dictionary<double, List<Element>> groupSecondary = Utility.GroupByElementsWithElevation(CongridDictionary1.Last().Value.Select(x => x.Conduit).ToList(), offsetVariable);
                                foreach (var elem in groupPrimary)
                                {
                                    foreach (var elem2 in elem.Value)
                                    {
                                        DistanceElements.Add(elem2);
                                    }
                                }
                                if (groupPrimary.Count == groupSecondary.Count)
                                {

                                    for (int i = 0; i < groupPrimary.Count; i++)
                                    {

                                        List<Element> primarySortedElementspre = SortbyPlane(doc, groupPrimary.ElementAt(i).Value);

                                        List<Element> secondarySortedElementspre = SortbyPlane(doc, groupSecondary.ElementAt(i).Value);


                                        bool isNotStaright = ReverseingConduits(doc, ref primarySortedElementspre, ref secondarySortedElementspre);

                                        //defind the primary and secondary sets 
                                        double conduitlengthone = primarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                        double conduitlengthsecond = secondarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                        List<Element> primarySortedElements = new List<Element>();
                                        List<Element> secondarySortedElements = new List<Element>();
                                        if (conduitlengthone < conduitlengthsecond)
                                        {
                                            primarySortedElements = primarySortedElementspre;
                                            secondarySortedElements = secondarySortedElementspre;
                                        }
                                        else
                                        {
                                            primarySortedElements = secondarySortedElementspre;
                                            secondarySortedElements = primarySortedElementspre;
                                        }

                                        if (primarySortedElements.Count == secondarySortedElements.Count)
                                        {
                                            Element primaryFirst = primarySortedElements.First();
                                            Element secondaryFirst = secondarySortedElements.First();
                                            Element primaryLast = primarySortedElements.Last();
                                            Element secondaryLast = secondarySortedElements.Last();

                                            XYZ priFirstDir = ((primaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ priLastDir = ((primaryLast.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ secFirstDir = ((secondaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ secLastDir = ((secondaryLast.Location as LocationCurve).Curve as Line).Direction;

                                            bool isSamDireFirst = Utility.IsSameDirectionWithRoundOff(priFirstDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priFirstDir, secLastDir, 3);
                                            bool isSamDireLast = Utility.IsSameDirectionWithRoundOff(priLastDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priLastDir, secLastDir, 3);
                                            //Same Elevations 
                                            bool isSamDir = !isNotStaright || isSamDireFirst && isSamDireLast;
                                            if (!isSamDir)
                                            {
                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);

                                                XYZ firstInte = MultiConnectFindIntersectionPoint(priFirst, secFirst);
                                                if (firstInte != null)
                                                {
                                                    firstInte = MultiConnectFindIntersectionPoint(priFirst, secLast);

                                                    if (firstInte != null)
                                                    {
                                                        isSamDir = false;
                                                    }
                                                }
                                            }
                                            if (!isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                            {
                                                //Multi connect
                                                System.Windows.MessageBox.Show("Warning. \n" + "Please use Multi Connect tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                return;
                                            }
                                            else if (isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                            {

                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                ConnectorSet firstConnectors = null;
                                                ConnectorSet secondConnectors = null;
                                                firstConnectors = Utility.GetConnectors(primaryFirst);
                                                secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));
                                                if (isSamDirecheckline)
                                                {
                                                    //Extend

                                                    if (ParentUserControl.Instance.shaddleControl.SelectedIndex == 0)
                                                    {

                                                        ThreePtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }
                                                    else
                                                    {
                                                        FourPtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }


                                                }
                                                else
                                                {
                                                    //Hoffset
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Horizontal Offset tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                    return;
                                                }
                                            }
                                            else
                                            {
                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                ConnectorSet firstConnectors = null;
                                                ConnectorSet secondConnectors = null;
                                                firstConnectors = Utility.GetConnectors(primaryFirst);
                                                secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));

                                                double priSlope = -Math.Round(priFirst.Direction.X, 6) / Math.Round(priFirst.Direction.Y, 6);
                                                double SecSlope = -Math.Round(secFirst.Direction.X, 6) / Math.Round(secFirst.Direction.Y, 6);

                                                if ((priSlope == -1 && SecSlope == 0) || Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4) == -1 || Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4).ToString() == double.NaN.ToString())
                                                {
                                                    //kick
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Kick tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                    return;
                                                }
                                                else if (isSamDirecheckline)
                                                {
                                                    if (ParentUserControl.Instance.shaddleControl.SelectedIndex == 0)
                                                    {

                                                        ThreePtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }
                                                    else
                                                    {
                                                        FourPtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }
                                                    //Voffset
                                                    /* VoffsetExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                     Utility.SetGlobalParametersManager(uiapp, "ConPak-VConnect-gp", JsonConvert.SerializeObject(globalParam));*/
                                                }
                                                else
                                                {
                                                    //Roffset
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Rolling Offset tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);

                                                    return;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (CongridDictionary1.Count == 1)
                            {
                                Utility.GroupByElevation(SelectedElements, offsetVariable, ref group);
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                    pickpoint = Utility.PickPoint(uidoc);

                                Dictionary<double, List<Element>> groupPrimary = new Dictionary<double, List<Element>>();
                                Dictionary<double, List<Element>> groupSecondary = new Dictionary<double, List<Element>>();
                                foreach (var elem in groupPrimary)
                                {
                                    foreach (var elem2 in elem.Value)
                                    {
                                        DistanceElements.Add(elem2);
                                    }
                                }
                                int k = group.Count / 2;
                                for (int i = 0; i < group.Count(); i++)
                                {
                                    if (i >= k)
                                    {
                                        groupSecondary.Add(group.ElementAt(i).Key, group.ElementAt(i).Value);
                                    }
                                    else
                                    {
                                        groupPrimary.Add(group.ElementAt(i).Key, group.ElementAt(i).Value);
                                    }


                                }

                                if (groupPrimary.Count == groupSecondary.Count)
                                {

                                    for (int i = 0; i < groupPrimary.Count; i++)
                                    {

                                        List<Element> primarySortedElementspre = SortbyPlane(doc, groupPrimary.ElementAt(i).Value);

                                        List<Element> secondarySortedElementspre = SortbyPlane(doc, groupSecondary.ElementAt(i).Value);


                                        bool isNotStaright = ReverseingConduits(doc, ref primarySortedElementspre, ref secondarySortedElementspre);

                                        //defind the primary and secondary sets 
                                        double conduitlengthone = primarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                        double conduitlengthsecond = secondarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                        List<Element> primarySortedElements = new List<Element>();
                                        List<Element> secondarySortedElements = new List<Element>();
                                        if (conduitlengthone < conduitlengthsecond)
                                        {
                                            primarySortedElements = primarySortedElementspre;
                                            secondarySortedElements = secondarySortedElementspre;
                                        }
                                        else
                                        {
                                            primarySortedElements = secondarySortedElementspre;
                                            secondarySortedElements = primarySortedElementspre;
                                        }

                                        if (primarySortedElements.Count == secondarySortedElements.Count)
                                        {
                                            Element primaryFirst = primarySortedElements.First();
                                            Element secondaryFirst = secondarySortedElements.First();
                                            Element primaryLast = primarySortedElements.Last();
                                            Element secondaryLast = secondarySortedElements.Last();

                                            XYZ priFirstDir = ((primaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ priLastDir = ((primaryLast.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ secFirstDir = ((secondaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                            XYZ secLastDir = ((secondaryLast.Location as LocationCurve).Curve as Line).Direction;

                                            bool isSamDireFirst = Utility.IsSameDirectionWithRoundOff(priFirstDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priFirstDir, secLastDir, 3);
                                            bool isSamDireLast = Utility.IsSameDirectionWithRoundOff(priLastDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priLastDir, secLastDir, 3);
                                            //Same Elevations 
                                            bool isSamDir = !isNotStaright || isSamDireFirst && isSamDireLast;
                                            if (!isSamDir)
                                            {
                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);

                                                XYZ firstInte = MultiConnectFindIntersectionPoint(priFirst, secFirst);
                                                if (firstInte != null)
                                                {
                                                    firstInte = MultiConnectFindIntersectionPoint(priFirst, secLast);

                                                    if (firstInte != null)
                                                    {
                                                        isSamDir = false;
                                                    }
                                                }
                                            }
                                            if (!isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                            {
                                                //Multi connect
                                                System.Windows.MessageBox.Show("Warning. \n" + "Please use Multi Connect tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                return;
                                            }
                                            else if (isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                            {

                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                ConnectorSet firstConnectors = null;
                                                ConnectorSet secondConnectors = null;
                                                firstConnectors = Utility.GetConnectors(primaryFirst);
                                                secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));
                                                if (isSamDirecheckline)
                                                {
                                                    //Extend

                                                    if (ParentUserControl.Instance.shaddleControl.SelectedIndex == 0)
                                                    {
                                                        ThreePtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }
                                                    else
                                                    {
                                                        FourPtSaddleExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                    }
                                                }
                                                else
                                                {
                                                    //Hoffset
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Horizontal Offset tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                    return;
                                                }
                                            }
                                            else
                                            {
                                                Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                ConnectorSet firstConnectors = null;
                                                ConnectorSet secondConnectors = null;
                                                firstConnectors = Utility.GetConnectors(primaryFirst);
                                                secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));

                                                double priSlope = -Math.Round(priFirst.Direction.X, 6) / Math.Round(priFirst.Direction.Y, 6);
                                                double SecSlope = -Math.Round(secFirst.Direction.X, 6) / Math.Round(secFirst.Direction.Y, 6);

                                                if ((priSlope == -1 && SecSlope == 0) || Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4) == -1 || Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4).ToString() == double.NaN.ToString())
                                                {
                                                    //kick
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Kick tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                    return;
                                                }
                                                else if (isSamDirecheckline)
                                                {
                                                    //Voffset
                                                    /* VoffsetExecute(_uiapp, ref primarySortedElements, ref secondarySortedElements);
                                                     Utility.SetGlobalParametersManager(uiapp, "ConPak-VConnect-gp", JsonConvert.SerializeObject(globalParam));*/
                                                }
                                                else
                                                {
                                                    //Roffset
                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Rolling Offset tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);

                                                    return;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            tr.Commit();
                        }
                    }
                    SelectedElements.Clear();
                    DistanceElements.Clear();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ParentUserControl.Instance._window.Close();
            }
        }

        public static List<Element> GetElementsByOder(List<Element> a_PrimaryElements)
        {
            List<Element> PrimaryElements = new List<Element>();
            XYZ PrimaryDirection = ((a_PrimaryElements.LastOrDefault().Location as LocationCurve).Curve as Line).Direction;
            if (Math.Abs(PrimaryDirection.Z) != 1)
            {
                PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).Y).ToList();
                if (PrimaryDirection.Y == 1 || PrimaryDirection.Y == -1)
                {
                    PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
                }
            }
            else
            {
                PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
            }
            return PrimaryElements;
        }

        #region Saddle Offset
        public void ThreePtSaddleExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            ElementsFilter filter = new ElementsFilter("Conduit Tags");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                startDate = DateTime.UtcNow;
                //SecondaryElements = Elements.GetElementsByReference(SecondaryReference, doc);
                LocationCurve findDirec = PrimaryElements[0].Location as LocationCurve;
                Line n = findDirec.Curve as Line;
                XYZ dire = n.Direction;
                Connector ConnectOne = null;
                Connector ConnectTwo = null;
                Utility.GetClosestConnectors(PrimaryElements[0], SecondaryElements[0], out ConnectOne, out ConnectTwo);
                XYZ ax = ConnectOne.Origin;
                Line pickline = null;
                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                {
                    pickline = Line.CreateBound(pickpoint, pickpoint + new XYZ(dire.X + 10, dire.Y, dire.Z));
                    if (dire.X == 1)
                    {
                        if (pickline.Origin.Y < ax.Y)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                    else if (dire.Y == -1)
                    {
                        PrimaryElements = GetElementsByOder(PrimaryElements);
                        SecondaryElements = GetElementsByOder(SecondaryElements);
                        /*if (pickline.Origin.X < ax.X)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }*/
                    }
                    else if (dire.X == -1)
                    {
                        if (pickline.Origin.Y < ax.Y)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                    else
                    {
                        if (pickline.Origin.X < ax.X)
                        {
                            if (dire.X == -1)
                            {
                                PrimaryElements = GetElementsByOder(PrimaryElements);
                                SecondaryElements = GetElementsByOder(SecondaryElements);
                            }
                            else
                            {
                                PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                                SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                            }
                        }
                        else
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                    }
                }
                else
                {
                    PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                    SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                }
                List<Element> thirdElements = new List<Element>();
                List<Element> forthElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;

                        tx.Start();
                        /* Line picked = null;
                         LocationCurve newcurve = PrimaryElements.LastOrDefault().Location as LocationCurve;
                         Line ncl = newcurve.Curve as Line;
                         XYZ direc = ncl.Direction;
                         picked = Line.CreateBound(new XYZ(pickpoint.X, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X , direc.Y, direc.Z));
                         Curve cur = picked as Curve;
                         // Create the DetailCurve element
                        // DetailCurve dtCurve = doc.Create.NewDetailCurve(doc.ActiveView, cur);*/
                        /* var s = uidoc.Selection.PickObject(ObjectType.Element, filter, "Select Tag");
                         IndependentTag refTag = doc.GetElement(s) as IndependentTag;
                         var re = refTag.get*/
                        double l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                        double offSet = 0;
                        double basedistance = 0;
                        if (!string.IsNullOrEmpty(ParentUserControl.Instance.txtheight.Text))
                        {
                            offSet = ParentUserControl.Instance.txtheight.AsDouble;
                            basedistance = (l_angle * (180 / Math.PI)) == 90 ? 1 : offSet / Math.Tan(l_angle);
                        }

                        double givendist = 0;
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            Element el = PrimaryElements[0];
                            LocationCurve findDirect = el.Location as LocationCurve;
                            Line ncDer = findDirect.Curve as Line;
                            XYZ dir = ncDer.Direction;
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            XYZ newenpt2 = new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, ConnectorTwo.Origin.Z);

                            Conduit newConCopy = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            Conduit newCon2Copy = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, newenpt2);
                            Parameter parameter = newConCopy.LookupParameter("Middle Elevation");
                            var middle = parameter.AsDouble();
                            XYZ Pri_mid = Utility.GetMidPoint(newConCopy);
                            XYZ Sec_mid = Utility.GetMidPoint(newCon2Copy);
                            /* Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                             Conduit newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);*/
                            double distance = 0;
                            LocationCurve newcurve = DistanceElements[0].Location as LocationCurve;
                            if (DistanceElements.Count() >= 2)
                            {
                                LocationCurve newcurve2 = DistanceElements[1].Location as LocationCurve;
                                XYZ start1 = newcurve.Curve.GetEndPoint(0);
                                XYZ start2 = newcurve2.Curve.GetEndPoint(0);
                                distance = start1.DistanceTo(start2);
                            }
                            Line ncl = newcurve.Curve as Line;
                            XYZ direc = ncl.Direction;
                            Conduit newCon = null;
                            Conduit newCon2 = null;
                            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {
                                if (dir.X == 1 )
                                {
                                    if (pickline.Origin.Y < ax.Y)
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X + 3, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X - 3, direc.Y, direc.Z));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 3, direc.Y - givendist, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 3, direc.Y - givendist, direc.Z));
                                        }
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X + 3, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X - 3, direc.Y, direc.Z));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 3, direc.Y + givendist, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 3, direc.Y + givendist, direc.Z));
                                        }
                                    }
                                }
                                else if (dir.X == -1)
                                {
                                    if (pickline.Origin.Y < ax.Y)
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint +direc.Multiply(-2));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 2, direc.Y - givendist, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 2, direc.Y - givendist, direc.Z));
                                        }
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(-2));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 2, direc.Y + givendist, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 2, direc.Y + givendist, direc.Z));
                                        }
                                    }
                                }
                                else if (dir.Y == -1) //vertical
                                {
                                    if (pickline.Origin.X < ax.X)
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint +  direc.Multiply(.5));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(-.5));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y - .5, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y - .5, direc.Z));
                                        }
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(.5));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(-.5));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + .5, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y - .5, direc.Z));
                                        }
                                    }
                                }
                                else //angled
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            XYZ end = pickpoint + new XYZ(direc.X, direc.Y, direc.Z);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, end);
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                                newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);
                                /* Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                 XYZ midlinept = Utility.GetMidPoint(centline);
                                 XYZ ConduitlineDir = centline.Direction;
                                 XYZ refStartPoint = ConnectorOne.Origin + ConduitlineDir.Multiply(Math.Abs(0));
                                 XYZ refEndPoint = refStartPoint + ConduitlineDir.Multiply(5);
                                 XYZ refStartPoint2 = ConnectorTwo.Origin + ConduitlineDir.Multiply(Math.Abs(0));
                                 XYZ refEndPoint2 = refStartPoint2 + ConduitlineDir.Multiply(-5);
                                 newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, refStartPoint, refEndPoint);

                                 newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, refStartPoint2, refEndPoint2);*/

                            }
                            Parameter param = newCon.LookupParameter("Middle Elevation");
                            Parameter param2 = newCon2.LookupParameter("Middle Elevation");
                            if (ParentUserControl.Instance.tagControl.SelectedIndex != 2)
                            {
                                /* double elevation = newCon.LookupParameter(offsetVariable).AsDouble();
                                 Parameter newElevation = newCon.LookupParameter(offsetVariable);
                                 newElevation.Set(elevation + offSet);
                                 Parameter newElevation2 = newCon2.LookupParameter(offsetVariable);
                                 newElevation2.Set(elevation + offSet);*/
                            }
                            else
                            {
                                param.Set(middle);
                                param2.Set(middle);
                            }
                            Utility.DeleteElement(doc, newConCopy.Id);
                            Utility.DeleteElement(doc, newCon2Copy.Id);

                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            Element e2 = doc.GetElement(newCon2.Id);
                            thirdElements.Add(e);
                            forthElements.Add(e2);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve findDirection = ElementOne.Location as LocationCurve;
                        Line nc = findDirection.Curve as Line;
                        XYZ direct = nc.Direction;
                        LocationCurve findDirection2 = ElementTwo.Location as LocationCurve;
                        Line nc2 = findDirection2.Curve as Line;
                        XYZ directDown = nc2.Direction;
                        //primary
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = null;
                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            axisStart = pickpoint;
                        else
                            axisStart = ConnectorOne.Origin;
                        XYZ axisSt = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        //secondary
                        LocationCurve newconcurve2 = forthElements[0].Location as LocationCurve;
                        Line ncl2 = newconcurve2.Curve as Line;
                        XYZ direction2 = ncl2.Direction;
                        XYZ axisStart2 = null;
                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            axisStart2 = pickpoint;
                        else
                            axisStart2 = ConnectorTwo.Origin;
                        XYZ axisEnd2 = axisStart2.Add(XYZ.BasisZ.CrossProduct(direction2));
                        Line axisLine2 = Line.CreateBound(axisStart2, axisEnd2);
                        Line pickedline = null;

                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                        {
                            pickedline = Line.CreateBound(pickpoint, pickpoint + new XYZ(direction.X + 10, direction.Y, direction.Z));
                            Curve cu = pickedline as Curve;
                            // Create the DetailCurve element
                            // DetailCurve detailCurve = doc.Create.NewDetailCurve(doc.ActiveView, cu);
                        }



                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }

                        try
                        {
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {

                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {
                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == 1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else
                                    {

                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);/////////////////////
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else //angle conduit
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                    else //right
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                {
                                    if (direct.Y != 0 )
                                    {
                                        if (pickedline.Origin.X < axisSt.X)//up
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                           else if (direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else if(direct.Y == -1 && directDown.Y == -1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else if (direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                    }
                                    else
                                    {
                                        if (pickedline.Origin.Y > axisSt.Y)
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                    }
                                }
                                else
                                {
                                    Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                }
                               

                            }
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {

                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2, -l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {
                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == 1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }

                                    else
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, l_angle);
                                    }

                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);/////////////
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else //right
                                    {

                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < SecondaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                {
                                    if (direct.Y != 0)
                                    {
                                        if (pickedline.Origin.X < axisSt.X)
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else if (direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else if (direct.Y == -1 && directDown.Y == -1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else if (direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                    }
                                    else
                                    {
                                        if (pickedline.Origin.Y > axisSt.Y)
                                        {
                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                        {

                                            if (direct.X == -1 && directDown.X == 1 || direct.Y == -1 && directDown.Y == 1)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                    }
                                }
                                else
                                {
                                    Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                }
                               
                            }
                            for (int i = 0; i < thirdElements.Count; i++)
                            {
                                try
                                {
                                    Utility.CreateElbowFittings(thirdElements[i], forthElements[i], doc, uiapp);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        Utility.CreateElbowFittings(forthElements[i], thirdElements[i], doc, uiapp);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                        return;
                                    }
                                }
                            }


                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {

                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }

                        tx.Commit();
                        doc.Regenerate();
                       // _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Vertical Offset", Util.ProductVersion, "Connect");

                    }
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        tx.Start();
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    //_ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");

                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
               // _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
            }
        }
        public void FourPtSaddleExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            ElementsFilter filter = new ElementsFilter("Conduit Tags");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                startDate = DateTime.UtcNow;
                //SecondaryElements = Elements.GetElementsByReference(SecondaryReference, doc);
                LocationCurve findDirec = PrimaryElements[0].Location as LocationCurve;
                Line n = findDirec.Curve as Line;
                XYZ dire = n.Direction;
                Connector ConnectOne = null;
                Connector ConnectTwo = null;
                Utility.GetClosestConnectors(PrimaryElements[0], SecondaryElements[0], out ConnectOne, out ConnectTwo);
                XYZ ax = ConnectOne.Origin;
                Line pickline = null;
                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                {
                    pickline = Line.CreateBound(pickpoint, pickpoint + new XYZ(dire.X + 10, dire.Y, dire.Z));
                    if (dire.X == 1||dire.X == -1)
                    {
                        if (pickline.Origin.Y < ax.Y)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                    else if (dire.Y == -1)
                    {
                        PrimaryElements = GetElementsByOder(PrimaryElements);
                        SecondaryElements = GetElementsByOder(SecondaryElements);

                        /*if (pickline.Origin.X < ax.X)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }*/
                    }
                    else
                    {
                        if (pickline.Origin.X < ax.X)//left
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else//right
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                }
                else
                {
                    PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                    SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                }
                List<Element> thirdElements = new List<Element>();
                List<Element> forthElements = new List<Element>();
                List<Element> FifthElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;
                        XYZ stpt = null;
                        XYZ edpt = null;

                        tx.Start();

                        double l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                        double offSet = 1;
                        double basedistance = 0;
                        if (!string.IsNullOrEmpty(ParentUserControl.Instance.txtheight.Text))
                        {
                            offSet = ParentUserControl.Instance.txtheight.AsDouble;
                            basedistance = (l_angle * (180 / Math.PI)) == 90 ? 1 : offSet / Math.Tan(l_angle);
                        }
                        double givendist = 0;
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            Element el = PrimaryElements[0];
                            LocationCurve findDirect = el.Location as LocationCurve;
                            Line ncDer = findDirect.Curve as Line;
                            XYZ dir = ncDer.Direction;
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            XYZ newenpt2 = new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, ConnectorTwo.Origin.Z);

                            Conduit newConCopy = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            Conduit newCon2Copy = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, newenpt2);
                            Parameter parameter = newConCopy.LookupParameter("Middle Elevation");
                            var middle = parameter.AsDouble();
                            XYZ Pri_mid = Utility.GetMidPoint(newConCopy);
                            XYZ Sec_mid = Utility.GetMidPoint(newCon2Copy);
                            // Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                            //Conduit newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);
                            double distance = 0;
                            LocationCurve newcurve = DistanceElements[0].Location as LocationCurve;
                            if (DistanceElements.Count() >= 2)
                            {
                                LocationCurve newcurve2 = DistanceElements[1].Location as LocationCurve;
                                XYZ start1 = newcurve.Curve.GetEndPoint(0);
                                XYZ start2 = newcurve2.Curve.GetEndPoint(0);
                                distance = start1.DistanceTo(start2);
                            }
                            Line ncl = newcurve.Curve as Line;
                            XYZ direc = ncl.Direction;
                            Conduit newCon = null;
                            Conduit newCon2 = null;
                            Conduit newCon3 = null;
                            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (dir.X == 1)
                                {
                                    if (pickline.Origin.Y < ax.Y)
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                            XYZ ConduitlineDir = centline.Direction;
                                            XYZ midlinept = Utility.GetMidPoint(centline);
                                            XYZ mid=Utility.GetMidPoint(newCon3);
                                            //stpt = mid- ConduitlineDir.Multiply(1);
                                           // edpt = mid + ConduitlineDir.Multiply(1);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + direc.Multiply(-2));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y - givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                            XYZ ConduitlineDir = centline.Direction;
                                            XYZ midlinept = Utility.GetMidPoint(centline);
                                            XYZ mid = Utility.GetMidPoint(newCon3);
                                           // stpt = mid - ConduitlineDir.Multiply(1);
                                            //edpt = mid + ConduitlineDir.Multiply(1);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X + 2, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X - 2, direc.Y, direc.Z));
                                        }
                                    }
                                    else //up
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X + 2, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X - 2, direc.Y, direc.Z));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y + givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X + 2, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X - 2, direc.Y, direc.Z));
                                        }
                                    }
                                }
                               else if ( dir.X == -1)
                                {
                                    if (pickline.Origin.Y < ax.Y)
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1.5, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + 1.5, direc.Y, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                            XYZ ConduitlineDir = centline.Direction;
                                            XYZ midlinept = Utility.GetMidPoint(centline);
                                            XYZ mid = Utility.GetMidPoint(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + direc.Multiply(-2));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1.5, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1.5, direc.Y - givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                            XYZ ConduitlineDir = centline.Direction;
                                            XYZ midlinept = Utility.GetMidPoint(centline);
                                            XYZ mid = Utility.GetMidPoint(newCon3);
                                          
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + direc.Multiply(-2));
                                        }
                                    }
                                    else //up
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + direc.Multiply(-2));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y + givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + direc.Multiply(-2));
                                        }
                                    }
                                }
                                else if (dir.Y == 1) //vertical
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                }
                                else if (dir.Y == -1) //vertical
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                }
                                else //angled
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y, pickpoint.Z), pickpoint + (new XYZ(direc.X, direc.Y, direc.Z)) * 2);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt - new XYZ(direc.X, direc.Y, direc.Z));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            XYZ refpic = new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z) + (new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z), refpic);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            XYZ refpic = new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z) + (new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), refpic);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                                {
                                    Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                    XYZ ConduitlineDir = centline.Direction;
                                    XYZ midlinept = Utility.GetMidPoint(centline);
                                    XYZ thiredconpt = new XYZ(midlinept.X, midlinept.Y, midlinept.Z);
                                    newCon3 = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(.5), thiredconpt + ConduitlineDir.Multiply(.5));
                                    Line con3 = Utility.GetLineFromConduit(newCon3);
                                    stpt = thiredconpt - ConduitlineDir.Multiply(2);
                                    edpt = thiredconpt + ConduitlineDir.Multiply(2);

                                    XYZ refEndPoint = (thiredconpt - ConduitlineDir.Multiply(2)) + ConduitlineDir.Multiply(-2);

                                    XYZ refEndPoint2 = (thiredconpt + ConduitlineDir.Multiply(2)) + ConduitlineDir.Multiply(2);
                                    /*newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), refEndPoint);

                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, thiredconpt + ConduitlineDir.Multiply(2), refEndPoint2);*/
                                    newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);
                                }
                                else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                                {
                                    Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                    XYZ ConduitlineDir = centline.Direction;
                                    XYZ midlinept = Utility.GetMidPoint(centline);
                                    XYZ thiredconpt = new XYZ(midlinept.X, midlinept.Y, midlinept.Z);
                                    newCon3 = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(.5), thiredconpt + ConduitlineDir.Multiply(.5));
                                    Line con3 = Utility.GetLineFromConduit(newCon3);
                                    stpt = con3.GetEndPoint(0);
                                    edpt = con3.GetEndPoint(1);

                                    XYZ refEndPoint = thiredconpt - ConduitlineDir.Multiply(2) + ConduitlineDir.Multiply(-2);

                                    XYZ refEndPoint2 = thiredconpt + ConduitlineDir.Multiply(2) + ConduitlineDir.Multiply(2);
                                   /* newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), refEndPoint);

                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, thiredconpt + ConduitlineDir.Multiply(2), refEndPoint2);*/
                                    newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);
                                }
                            }
                            Parameter param = newCon.LookupParameter("Middle Elevation");
                            Parameter param2 = newCon2.LookupParameter("Middle Elevation");
                            Parameter param3 = newCon3.LookupParameter("Middle Elevation");
                            if (ParentUserControl.Instance.tagControl.SelectedIndex != 2)
                            {
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                                {
                                    double elevation = newCon.LookupParameter(offsetVariable).AsDouble();
                                    Parameter newElevation = newCon.LookupParameter(offsetVariable);
                                    newElevation.Set(elevation + offSet);
                                    Parameter newElevation2 = newCon2.LookupParameter(offsetVariable);
                                    newElevation2.Set(elevation + offSet);
                                    Parameter newElevation3 = newCon3.LookupParameter(offsetVariable);
                                    newElevation3.Set(elevation + offSet);
                                }
                                else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                                {
                                    double elevation = newCon.LookupParameter(offsetVariable).AsDouble();
                                    Parameter newElevation = newCon.LookupParameter(offsetVariable);
                                    newElevation.Set(elevation - offSet);
                                    Parameter newElevation2 = newCon2.LookupParameter(offsetVariable);
                                    newElevation2.Set(elevation - offSet);
                                    Parameter newElevation3 = newCon3.LookupParameter(offsetVariable);
                                    newElevation3.Set(elevation - offSet);
                                }
                            }
                            else
                            {
                                param.Set(middle);
                                param2.Set(middle);
                                param3.Set(middle);
                            }
                            Utility.DeleteElement(doc, newConCopy.Id);
                            Utility.DeleteElement(doc, newCon2Copy.Id);

                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            Element e2 = doc.GetElement(newCon2.Id);
                            Element e3 = doc.GetElement(newCon3.Id);
                            thirdElements.Add(e);
                            forthElements.Add(e2);
                            FifthElements.Add(e3);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve findDirection = ElementOne.Location as LocationCurve;
                        Line nc = findDirection.Curve as Line;
                        XYZ direct = nc.Direction;
                        LocationCurve findDirection2 = ElementTwo.Location as LocationCurve;
                        Line nc2 = findDirection2.Curve as Line;
                        XYZ directDown = nc2.Direction;
                        //primary
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = null;
                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            axisStart = stpt;
                        else
                            axisStart = ConnectorOne.Origin; ;
                        XYZ axisSt = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        //secondary
                        LocationCurve newconcurve2 = forthElements[0].Location as LocationCurve;
                        Line ncl2 = newconcurve2.Curve as Line;
                        XYZ direction2 = ncl2.Direction;
                        XYZ axisStart2 = null;
                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            axisStart2 = edpt;
                        else
                            axisStart2 = ConnectorTwo.Origin; ;
                        XYZ axisEnd2 = axisStart2.Add(XYZ.BasisZ.CrossProduct(direction2));
                        Line axisLine2 = Line.CreateBound(axisStart2, axisEnd2);
                        Line pickedline = null;

                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                        {
                            pickedline = Line.CreateBound(pickpoint, pickpoint + new XYZ(direction.X + 10, direction.Y, direction.Z));
                            Curve cu = pickedline as Curve;
                            // Create the DetailCurve element
                            // DetailCurve detailCurve = doc.Create.NewDetailCurve(doc.ActiveView, cu);
                        }



                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                      /*  if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }*/

                        try
                        {
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {

                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {
                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -l_angle);
                            }

                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == 1 )
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else
                                    {

                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                               else if (direct.X == -1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)//down
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1,- l_angle);
                                    }
                                    else
                                    {

                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);/////////////////////
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else //angle conduit
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else //right
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                try
                                {
                                    if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                    {
                                        if ((pickedline.Origin.Y > axisSt.Y))
                                        {
                                            if (direct.Y == -1 && directDown.Y == 1 || direct.X == -1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else if(direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else if (direct.Y == -1 && directDown.Y == -1)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                        {
                                            if (direct.Y == -1 && directDown.Y == 1 || direct.X == -1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);

                                        }
                                    }
                                    else
                                    {
                                        Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                    }
                                }
                                catch
                                {
                                    Utility.CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp);
                                }

                            }

                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {

                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {
                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2,- l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == 1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }

                                    else
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                               else if (direct.X == -1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y) //down
                                    {
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, l_angle);
                                    }

                                    else
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }

                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);/////////////
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l1 = Line.CreateBound(end, axisStart2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else //right
                                    {

                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < SecondaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                Line tr = Utility.GetLineFromConduit(secondElement);
                                try
                                {
                                    if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                                    {
                                        if (pickedline.Origin.Y > axisSt.Y)
                                        {
                                            if (direct.Y == -1 && directDown.Y == 1 || direct.X == -1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else if (direct.X == 1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else if (direct.Y == -1 && directDown.Y == -1)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                        {
                                            if (direct.Y == -1 && directDown.Y == 1 || direct.X == -1 && directDown.X == 1)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                    }
                                    else
                                    {
                                        Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                    }
                                }
                                catch
                                {
                                    Utility.CreateElbowFittings(forthElement, SecondaryElements[i], doc, uiapp);
                                }

                            }
                            for (int i = 0; i < thirdElements.Count; i++)
                            {
                                try
                                {
                                    Utility.CreateElbowFittings(thirdElements[i], FifthElements[i], doc, uiapp);
                                    Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        Utility.CreateElbowFittings(FifthElements[i], thirdElements[i], doc, uiapp);
                                        Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                    }
                                    catch (Exception)
                                    {
                                        Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                        Utility.CreateElbowFittings(FifthElements[i], thirdElements[i], doc, uiapp);
                                    }
                                }
                            }


                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {
                                                               // tx.RollBack();
                                                               // DeleteElements(doc, thirdElements);
                                                                //DeleteElements(doc, forthElements);
                                                                //DeleteElements(doc, FifthElements);
                                                              
                                                                //FourPtSaddleReverse(uiapp,ref PrimaryElements,ref SecondaryElements);
                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }

                        tx.Commit();
                        doc.Regenerate();
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Vertical Offset", Util.ProductVersion, "Connect");

                    }
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        tx.Start();
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");

                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
            }
        }
        public void FourPtSaddleReverse(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            ElementsFilter filter = new ElementsFilter("Conduit Tags");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                startDate = DateTime.UtcNow;
                //SecondaryElements = Elements.GetElementsByReference(SecondaryReference, doc);
                LocationCurve findDirec = PrimaryElements[0].Location as LocationCurve;
                Line n = findDirec.Curve as Line;
                XYZ dire = n.Direction;
                Connector ConnectOne = null;
                Connector ConnectTwo = null;
                Utility.GetClosestConnectors(PrimaryElements[0], SecondaryElements[0], out ConnectOne, out ConnectTwo);
                XYZ ax = ConnectOne.Origin;
                Line pickline = null;
                if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                {
                    pickline = Line.CreateBound(pickpoint, pickpoint + new XYZ(dire.X + 10, dire.Y, dire.Z));
                    if (dire.X == 1)
                    {
                        if (pickline.Origin.Y < ax.Y)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                    else if (dire.Y == -1)
                    {
                        PrimaryElements = GetElementsByOder(PrimaryElements);
                        SecondaryElements = GetElementsByOder(SecondaryElements);

                        /*if (pickline.Origin.X < ax.X)
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }*/
                    }
                    else
                    {
                        if (pickline.Origin.X < ax.X)//left
                        {
                            PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                            SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                        }
                        else//right
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                }
                else
                {
                    PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                    SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                }
                List<Element> thirdElements = new List<Element>();
                List<Element> forthElements = new List<Element>();
                List<Element> FifthElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;
                        XYZ stpt = null;
                        XYZ edpt = null;

                        tx.Start();

                        double l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                        double offSet = 1;
                        double basedistance = 0;
                        if (!string.IsNullOrEmpty(ParentUserControl.Instance.txtheight.Text))
                        {
                            offSet = ParentUserControl.Instance.txtheight.AsDouble;
                            basedistance = (l_angle * (180 / Math.PI)) == 90 ? 1 : offSet / Math.Tan(l_angle);
                        }
                        double givendist = 0;
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            Element el = PrimaryElements[0];
                            LocationCurve findDirect = el.Location as LocationCurve;
                            Line ncDer = findDirect.Curve as Line;
                            XYZ dir = ncDer.Direction;
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            XYZ newenpt2 = new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, ConnectorTwo.Origin.Z);

                            Conduit newConCopy = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            Conduit newCon2Copy = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, newenpt2);
                            Parameter parameter = newConCopy.LookupParameter("Middle Elevation");
                            var middle = parameter.AsDouble();
                            XYZ Pri_mid = Utility.GetMidPoint(newConCopy);
                            XYZ Sec_mid = Utility.GetMidPoint(newCon2Copy);
                            // Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, Pri_mid);
                            //Conduit newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, Sec_mid);
                            double distance = 0;
                            LocationCurve newcurve = DistanceElements[0].Location as LocationCurve;
                            if (DistanceElements.Count() >= 2)
                            {
                                LocationCurve newcurve2 = DistanceElements[1].Location as LocationCurve;
                                XYZ start1 = newcurve.Curve.GetEndPoint(0);
                                XYZ start2 = newcurve2.Curve.GetEndPoint(0);
                                distance = start1.DistanceTo(start2);
                            }
                            Line ncl = newcurve.Curve as Line;
                            XYZ direc = ncl.Direction;
                            Conduit newCon = null;
                            Conduit newCon2 = null;
                            Conduit newCon3 = null;
                            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (dir.X == 1)
                                {
                                    if (pickline.Origin.Y < ax.Y)
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + 1, pickpoint.Y, pickpoint.Z), pickpoint + direc.Multiply(-1));
                                          Line con3 = Utility.GetLineFromConduit(newCon3);
                                            /*   Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                               XYZ ConduitlineDir = centline.Direction;
                                               XYZ midlinept = Utility.GetMidPoint(centline);
                                               XYZ mid = Utility.GetMidPoint(newCon3);
                                               stpt = ConnectOne.Origin;
                                               edpt = ConnectorTwo.Origin;*/
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + direc.Multiply(2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + direc.Multiply(-2));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y - givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                         /*   Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                            XYZ ConduitlineDir = centline.Direction;
                                            XYZ midlinept = Utility.GetMidPoint(centline);
                                            XYZ mid = Utility.GetMidPoint(newCon3);
                                            stpt = mid + ConduitlineDir.Multiply(1);
                                            edpt = mid - ConduitlineDir.Multiply(1);*/
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + direc.Multiply(-2));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + direc.Multiply(2));
                                        }
                                    }
                                    else //up
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X + 2, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X - 2, direc.Y, direc.Z));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - 1, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 1, direc.Y + givendist, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X + 2, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X - 2, direc.Y, direc.Z));
                                        }
                                    }
                                }
                                else if (dir.Y == -1) //vertical
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {
                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y - 1.5, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + 1.5, direc.Z));
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(0);
                                            edpt = con3.GetEndPoint(1);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt + new XYZ(direc.X, direc.Y + 1.5, direc.Z));
                                        }
                                    }
                                }
                                else //angled
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y, pickpoint.Z), pickpoint + (new XYZ(direc.X, direc.Y, direc.Z)) * 2);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt - new XYZ(direc.X, direc.Y, direc.Z));

                                        }
                                        else
                                        {
                                            givendist += distance;
                                            XYZ refpic = new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z) + (new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z), refpic);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, stpt, stpt + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, edpt, edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            XYZ refpic = new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z) + (new XYZ(direc.X, direc.Y, direc.Z) * 2);
                                            newCon3 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), refpic);
                                            Line con3 = Utility.GetLineFromConduit(newCon3);
                                            stpt = con3.GetEndPoint(1);
                                            edpt = con3.GetEndPoint(0);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(stpt.X, stpt.Y, stpt.Z), stpt + new XYZ(direc.X, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(edpt.X, edpt.Y, edpt.Z), edpt - new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                                {
                                    Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                    XYZ ConduitlineDir = centline.Direction;
                                    XYZ midlinept = Utility.GetMidPoint(centline);
                                    XYZ thiredconpt = new XYZ(midlinept.X, midlinept.Y, midlinept.Z);
                                    newCon3 = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), thiredconpt + ConduitlineDir.Multiply(2));
                                    Line con3 = Utility.GetLineFromConduit(newCon3);
                                    stpt = thiredconpt - ConduitlineDir.Multiply(2);
                                    edpt = thiredconpt + ConduitlineDir.Multiply(2);

                                    XYZ refEndPoint = (thiredconpt - ConduitlineDir.Multiply(2)) + ConduitlineDir.Multiply(-2);

                                    XYZ refEndPoint2 = (thiredconpt + ConduitlineDir.Multiply(2)) + ConduitlineDir.Multiply(2);
                                    newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), refEndPoint);

                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, thiredconpt + ConduitlineDir.Multiply(2), refEndPoint2);
                                }
                                else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                                {
                                    Line centline = Line.CreateBound(ConnectorOne.Origin, ConnectorTwo.Origin);
                                    XYZ ConduitlineDir = centline.Direction;
                                    XYZ midlinept = Utility.GetMidPoint(centline);
                                    XYZ thiredconpt = new XYZ(midlinept.X, midlinept.Y, midlinept.Z);
                                    newCon3 = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), thiredconpt + ConduitlineDir.Multiply(2));
                                    Line con3 = Utility.GetLineFromConduit(newCon3);
                                    stpt = con3.GetEndPoint(0);
                                    edpt = con3.GetEndPoint(1);

                                    XYZ refEndPoint = thiredconpt - ConduitlineDir.Multiply(2) + ConduitlineDir.Multiply(-2);

                                    XYZ refEndPoint2 = thiredconpt + ConduitlineDir.Multiply(2) + ConduitlineDir.Multiply(2);
                                    newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, thiredconpt - ConduitlineDir.Multiply(2), refEndPoint);

                                    newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, thiredconpt + ConduitlineDir.Multiply(2), refEndPoint2);
                                }
                            }
                            Parameter param = newCon.LookupParameter("Middle Elevation");
                            Parameter param2 = newCon2.LookupParameter("Middle Elevation");
                            Parameter param3 = newCon3.LookupParameter("Middle Elevation");
                            if (ParentUserControl.Instance.tagControl.SelectedIndex != 2)
                            {
                                if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                                {
                                    double elevation = newCon.LookupParameter(offsetVariable).AsDouble();
                                    Parameter newElevation = newCon.LookupParameter(offsetVariable);
                                    newElevation.Set(elevation + offSet);
                                    Parameter newElevation2 = newCon2.LookupParameter(offsetVariable);
                                    newElevation2.Set(elevation + offSet);
                                    Parameter newElevation3 = newCon3.LookupParameter(offsetVariable);
                                    newElevation3.Set(elevation + offSet);
                                }
                                else if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                                {
                                    double elevation = newCon.LookupParameter(offsetVariable).AsDouble();
                                    Parameter newElevation = newCon.LookupParameter(offsetVariable);
                                    newElevation.Set(elevation - offSet);
                                    Parameter newElevation2 = newCon2.LookupParameter(offsetVariable);
                                    newElevation2.Set(elevation - offSet);
                                    Parameter newElevation3 = newCon3.LookupParameter(offsetVariable);
                                    newElevation3.Set(elevation - offSet);
                                }
                            }
                            else
                            {
                                param.Set(middle);
                                param2.Set(middle);
                                param3.Set(middle);
                            }
                            Utility.DeleteElement(doc, newConCopy.Id);
                            Utility.DeleteElement(doc, newCon2Copy.Id);

                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            Element e2 = doc.GetElement(newCon2.Id);
                            Element e3 = doc.GetElement(newCon3.Id);
                            thirdElements.Add(e);
                            forthElements.Add(e2);
                            FifthElements.Add(e3);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve findDirection = ElementOne.Location as LocationCurve;
                        Line nc = findDirection.Curve as Line;
                        XYZ direct = nc.Direction;
                        //primary
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = stpt;
                        XYZ axisSt = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        //secondary
                        LocationCurve newconcurve2 = forthElements[0].Location as LocationCurve;
                        Line ncl2 = newconcurve2.Curve as Line;
                        XYZ direction2 = ncl2.Direction;
                        XYZ axisStart2 = edpt;
                        XYZ axisEnd2 = axisStart2.Add(XYZ.BasisZ.CrossProduct(direction2));
                        Line axisLine2 = Line.CreateBound(axisStart2, axisEnd2);
                        Line pickedline = null;

                        if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                        {
                            pickedline = Line.CreateBound(pickpoint, pickpoint + new XYZ(direction.X + 10, direction.Y, direction.Z));
                            Curve cu = pickedline as Curve;
                            // Create the DetailCurve element
                            // DetailCurve detailCurve = doc.Create.NewDetailCurve(doc.ActiveView, cu);
                        }



                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }

                        try
                        {
                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {

                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {
                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle);
                            }

                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == -1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1,-l_angle);
                                    }
                                    else
                                    {

                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);/////////////////////
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else //angle conduit
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else //right
                                    {
                                        XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                        Line l1 = Line.CreateBound(axisStart, end);
                                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                try
                                {

                                     Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                }
                                catch
                                {
                                    Utility.CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp);
                                }

                            }

                            if (ParentUserControl.Instance.tagControl.SelectedIndex == 1)
                            {

                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 0)
                            {
                                ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), axisLine2, l_angle);
                            }
                            else if (ParentUserControl.Instance.tagControl.SelectedIndex == 2)
                            {

                                if (direct.X == -1)
                                {
                                    if (pickedline.Origin.Y < axisSt.Y)
                                    {
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2,l_angle);
                                    }

                                    else
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                }
                                else if (direct.Y == -1)
                                {
                                    if (pickedline.Origin.X < axisSt.X)
                                    {
                                        //left in vertical
                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }

                                    else
                                    {
                                        //right in vertical
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);/////////////
                                        Line l1 = Line.CreateBound(axisStart2, end);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                    }
                                }
                                else
                                {
                                    if (pickedline.Origin.X < axisSt.X) //left
                                    {
                                        XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                        Line l1 = Line.CreateBound(end, axisStart2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                    }
                                    else //right
                                    {

                                        XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                        Line l2 = Line.CreateBound(axisStart2, end2);
                                        ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                    }
                                }

                            }

                            for (int i = 0; i < SecondaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                try
                                {
                                   Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                }
                                catch
                                {
                                    Utility.CreateElbowFittings(forthElement, SecondaryElements[i], doc, uiapp);
                                }

                            }
                            for (int i = 0; i < thirdElements.Count; i++)
                            {
                                try
                                {
                                   Utility.CreateElbowFittings(thirdElements[i], FifthElements[i], doc, uiapp);
                                     Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        Utility.CreateElbowFittings(FifthElements[i], thirdElements[i], doc, uiapp);
                                        Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                    }
                                    catch (Exception)
                                    {
                                        Utility.CreateElbowFittings(FifthElements[i], forthElements[i], doc, uiapp);
                                        Utility.CreateElbowFittings(FifthElements[i], thirdElements[i], doc, uiapp);
                                    }
                                }
                            }


                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {

                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }

                        tx.Commit();
                        doc.Regenerate();
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Vertical Offset", Util.ProductVersion, "Connect");

                    }
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        tx.Start();
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");

                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
            }
        }
        #endregion
        private static Line AlignElement(Element pickedElement, XYZ refPoint, Document doc)
        {
            Line NewLine = null;
            using (SubTransaction subTx = new SubTransaction(doc))
            {
                subTx.Start();
                Line firstLine = (pickedElement.Location as LocationCurve).Curve as Line;
                XYZ startPoint = firstLine.GetEndPoint(0);
                XYZ endPoint = firstLine.GetEndPoint(1);
                LocationCurve curve = pickedElement.Location as LocationCurve;
                XYZ normal = firstLine.Direction;
                XYZ cross = normal.CrossProduct(XYZ.BasisZ);
                XYZ newEndPoint = refPoint + cross.Multiply(5);
                Line boundLine = Line.CreateBound(refPoint, newEndPoint);
                XYZ interSecPoint = Utility.FindIntersectionPoint(firstLine.GetEndPoint(0), firstLine.GetEndPoint(1), boundLine.GetEndPoint(0), boundLine.GetEndPoint(1));
                interSecPoint = new XYZ(interSecPoint.X, interSecPoint.Y, startPoint.Z);
                ConnectorSet connectorSet = Utility.GetUnusedConnectors(pickedElement);
                if (connectorSet.Size == 2)
                {
                    if (startPoint.DistanceTo(interSecPoint) > endPoint.DistanceTo(interSecPoint))
                    {
                        NewLine = Line.CreateBound(startPoint, interSecPoint);
                    }
                    else
                    {
                        NewLine = Line.CreateBound(interSecPoint, endPoint);
                    }
                }
                else
                {
                    connectorSet = Utility.GetConnectors(pickedElement);
                    foreach (Connector con in connectorSet)
                    {
                        if (con.IsConnected)
                        {
                            if (Utility.IsXYZTrue(con.Origin, startPoint))
                            {
                                NewLine = Line.CreateBound(con.Origin, interSecPoint);
                                break;
                            }
                            if (Utility.IsXYZTrue(con.Origin, endPoint))
                            {
                                NewLine = Line.CreateBound(interSecPoint, con.Origin);
                                break;
                            }
                        }
                    }
                }
                subTx.Commit();
            }
            return NewLine;
        }
        public void DeleteElements(Document doc, List<Element> elementsToDelete)
        {
            // Start a transaction
            using (Transaction transaction = new Transaction(doc, "Delete Elements"))
            {
                transaction.Start();

                // Iterate through the list of elements and delete each one
                foreach (Element element in elementsToDelete)
                {
                    // Use the Delete method to delete the element
                    ElementId elementId = element.Id;
                    doc.Delete(elementId);
                }

                // Commit the transaction
                transaction.Commit();
            }
        }
        private static List<Element> SortbyPlane(Document doc, List<Element> arrelements)
        {
            List<Element> conduitCollection = new List<Element>();

            //ascending conduits based on the intersection
            Dictionary<double, Element> dictcond = new Dictionary<double, Element>();
            View view = doc.ActiveView;
            XYZ vieworgin = view.Origin;
            XYZ viewdirection = view.ViewDirection;

            Line CondutitLine1 = (arrelements.First().Location as LocationCurve).Curve as Line;
            XYZ vieworgin1 = CondutitLine1.Origin;

            foreach (Element c in arrelements)
            {
                conduitCollection.Clear();
                Line CondutitLine = (c.Location as LocationCurve).Curve as Line;

                SketchPlane sp = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(CondutitLine1.Direction, vieworgin1));

                //SketchPlane sp = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(viewdirection, vieworgin));

                double denominator = CondutitLine.Direction.Normalize().DotProduct(sp.GetPlane().Normal);
                double numerator = (sp.GetPlane().Origin - CondutitLine.GetEndPoint(0)).DotProduct(sp.GetPlane().Normal);

                double parameter = numerator / denominator;
                XYZ intersectionPoint = CondutitLine.GetEndPoint(0) + parameter * CondutitLine.Direction;

                double xdirection = CondutitLine.Direction.X;
                double ydirection = CondutitLine.Direction.Y;
                double zdirection = CondutitLine.Direction.Z;

                if (ydirection == -1 || ydirection == 1)
                {
                    dictcond.Add(intersectionPoint.X, c);
                }
                //else if(zdirection == -1 || zdirection == 1)
                //{

                //}
                else
                {

                    dictcond.Add(intersectionPoint.Y, c);


                }


            }
            conduitCollection = dictcond.OrderBy(x => x.Key).Select(x => x.Value as Element).ToList();

            return conduitCollection;
        }

        public static XYZ MultiConnectFindIntersectionPoint(Line lineOne, Line lineTwo)
        {
            return MultiConnectFindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), lineTwo.GetEndPoint(0), lineTwo.GetEndPoint(1));
        }

        public static XYZ MultiConnectFindIntersectionPoint(XYZ s1, XYZ e1, XYZ s2, XYZ e2)
        {
            s1 = Utility.XYZroundOf(s1, 5);
            e1 = Utility.XYZroundOf(e1, 5);
            s2 = Utility.XYZroundOf(s2, 5);
            e2 = Utility.XYZroundOf(e2, 5);

            double a1 = e1.Y - s1.Y;
            double b1 = s1.X - e1.X;
            double c1 = a1 * s1.X + b1 * s1.Y;

            double a2 = e2.Y - s2.Y;
            double b2 = s2.X - e2.X;
            double c2 = a2 * s2.X + b2 * s2.Y;

            double delta = a1 * b2 - a2 * b1;
            //If lines are parallel, the result will be (NaN, NaN).
            return delta == 0 || Convert.ToString(delta).Contains("E") == true ? null
                : new XYZ((b2 * c1 - b1 * c2) / delta, (a1 * c2 - a2 * c1) / delta, 0);
        }
        private bool ReverseingConduits(Document doc, ref List<Element> primaryElements, ref List<Element> secondaryElements)
        {
            Line priFirst = ((primaryElements.First().Location as LocationCurve).Curve as Line);
            Line prilast = ((primaryElements.Last().Location as LocationCurve).Curve as Line);
            Line secFirst = ((secondaryElements.First().Location as LocationCurve).Curve as Line);
            Line seclast = ((secondaryElements.Last().Location as LocationCurve).Curve as Line);

            XYZ firstinter = MultiConnectFindIntersectionPoint(priFirst, secFirst);
            XYZ lastinter = MultiConnectFindIntersectionPoint(prilast, seclast);
            if (firstinter == null || lastinter == null)
            {
                return false;
            }
            priFirst = AlignElement(primaryElements.First(), firstinter, doc);
            secFirst = AlignElement(secondaryElements.First(), firstinter, doc);
            prilast = AlignElement(primaryElements.Last(), lastinter, doc);
            seclast = AlignElement(secondaryElements.Last(), lastinter, doc);

            Line primFirstextentionline = Line.CreateBound(new XYZ(priFirst.GetEndPoint(0).X, priFirst.GetEndPoint(0).Y, 0), new XYZ(priFirst.GetEndPoint(1).X, priFirst.GetEndPoint(1).Y, 0));
            Line secoFirstnextentionline = Line.CreateBound(new XYZ(secFirst.GetEndPoint(0).X, secFirst.GetEndPoint(0).Y, 0), new XYZ(secFirst.GetEndPoint(1).X, secFirst.GetEndPoint(1).Y, 0));
            Line primLastextentionline = Line.CreateBound(new XYZ(prilast.GetEndPoint(0).X, prilast.GetEndPoint(0).Y, 0), new XYZ(prilast.GetEndPoint(1).X, prilast.GetEndPoint(1).Y, 0));
            Line secoLastnextentionline = Line.CreateBound(new XYZ(seclast.GetEndPoint(0).X, seclast.GetEndPoint(0).Y, 0), new XYZ(seclast.GetEndPoint(1).X, seclast.GetEndPoint(1).Y, 0));

            XYZ interpointset1 = Utility.GetIntersection(primFirstextentionline, secoLastnextentionline);
            XYZ interpointset2 = Utility.GetIntersection(secoFirstnextentionline, primLastextentionline);
            if (interpointset1 == null || interpointset2 == null)
            {
                secondaryElements.Reverse();
            }
            if (interpointset1 == null && interpointset2 == null)
            {
                primaryElements.Reverse();
            }
            return true;
        }

        #region Vertical Offset
        public void VoffsetExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            ElementsFilter filter = new ElementsFilter("Conduits");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                startDate = DateTime.UtcNow;
                //SecondaryElements = Elements.GetElementsByReference(SecondaryReference, doc);
                PrimaryElements = APICommon.GetElementsByOder(PrimaryElements);
                SecondaryElements = APICommon.GetElementsByOder(SecondaryElements);
                List<Element> thirdElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;
                        tx.Start();
                        double l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            thirdElements.Add(e);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }
                        try
                        {
                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -l_angle);
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                            }
                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {

                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }

                        tx.Commit();
                        doc.Regenerate();
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Vertical Offset", Util.ProductVersion, "Connect");

                    }
                    using (SubTransaction tx = new SubTransaction(doc))
                    {
                        tx.Start();
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");

                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Vertical Offset", Util.ProductVersion, "Connect");
            }
        }
        #endregion

        #region Horizontal Offset
        public void HoffsetExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
            DateTime startDate = DateTime.UtcNow;
            ElementsFilter filter = new ElementsFilter("Conduits");
            //IList<Reference> reference = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select conduits to add the horizontal offset");
            try
            {

                List<Element> thirdElements = new List<Element>();
                double angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);

                //string direction = H_Offset.Instance.HDirection.SelectedValue.ToString();
                //double conduitLength = Utility.FeetnInchesToDouble(H_Offset.Instance.txtLength.Text);
                bool isVerticalConduits = false;
                //XYZ Pickpoint = uidoc.Selection.PickPoint();
                using (SubTransaction tx = new SubTransaction(doc))
                {
                    tx.Start();

                    for (int i = 0; i < PrimaryElements.Count; i++)
                    {
                        Element firstElement = PrimaryElements[i];
                        Element secondElement = SecondaryElements[i];
                        Line firstLine = (firstElement.Location as LocationCurve).Curve as Line;
                        Line secondLine = (secondElement.Location as LocationCurve).Curve as Line;
                        Line newLine = Utility.GetParallelLine(firstElement, secondElement, ref isVerticalConduits);
                        double elevation = firstElement.LookupParameter(offsetVariable).AsDouble();
                        Conduit thirdConduit = Utility.CreateConduit(doc, firstElement as Conduit, newLine.GetEndPoint(0), newLine.GetEndPoint(1));
                        Element thirdElement = doc.GetElement(thirdConduit.Id);
                        thirdElements.Add(thirdElement);
                        //RetainParameters(firstElement, secondElement, doc);
                        //RetainParameters(firstElement, thirdElement, doc);
                    }
                    //Rotate Elements at Once
                    Element ElementOne = PrimaryElements[0];
                    Element ElementTwo = SecondaryElements[0];
                    Utility.GetClosestConnectors(ElementOne, ElementTwo, out Connector ConnectorOne, out Connector ConnectorTwo);
                    XYZ axisStart = ConnectorOne.Origin;
                    XYZ axisEnd = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                    Line axisLine = Line.CreateBound(axisStart, axisEnd);
                    //if (( offSet > 0) || (offSet < 0) ||
                    //    (offSet < 0) ||(offSet > 0))
                    //{
                    //    angle = -angle;
                    //}

                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, angle);

                    //Conduit rotate angle indentification
                    Conduit SecondConduit = SecondaryElements[0] as Conduit;
                    Line SceondConduitLine = (SecondConduit.Location as LocationCurve).Curve as Line;
                    XYZ pt1 = SceondConduitLine.GetEndPoint(0);
                    XYZ pt2 = SceondConduitLine.GetEndPoint(1);
                    XYZ SecondLineDirection = SceondConduitLine.Direction;
                    pt1 -= SecondLineDirection.Multiply(10);
                    Line firstline = Line.CreateBound(pt1, pt2);

                    Conduit ThirdConduit = thirdElements[0] as Conduit;
                    Line ThirdConduitLine = (ThirdConduit.Location as LocationCurve).Curve as Line;
                    XYZ pt3 = ThirdConduitLine.GetEndPoint(0);
                    XYZ pt4 = ThirdConduitLine.GetEndPoint(1);
                    XYZ ThirdLineDirection = ThirdConduitLine.Direction;
                    pt4 += ThirdLineDirection.Multiply(10);
                    Line secondline = Line.CreateBound(pt3, pt4);

                    XYZ IntersectionforangleConduit = Utility.GetIntersection(firstline, secondline);
                    if (IntersectionforangleConduit == null)
                    {
                        angle = 2 * angle;
                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -angle);
                    }


                    for (int i = 0; i < PrimaryElements.Count; i++)
                    {
                        Element firstElement = PrimaryElements[i];
                        Element secondElement = SecondaryElements[i];
                        Element thirdElement = thirdElements[i];
                        Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                        Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                        Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                        Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                    }
                    tx.Commit();
                    _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Horizontal Offset", Util.ProductVersion, "Draw");
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Warning. \n" + ex.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Horizontal Offset", Util.ProductVersion, "Draw");
            }
        }
        #endregion

        #region Extend 
        public void ExtendExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            DateTime startDate = DateTime.UtcNow;
            try
            {
                //IList<Reference> PrimaryReference = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select 1st set of conduits to connect");
                List<ElementId> ConduitCollection1 = new List<ElementId>();
                foreach (Element e in PrimaryElements)
                {
                    ConduitCollection1.Add(e.Id);
                }
                //IList<Reference> SecondaryReference = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select 2nd set of conduits to connect");
                List<ElementId> ConduitCollection2 = new List<ElementId>();
                foreach (Element e in SecondaryElements)
                {
                    ConduitCollection2.Add(e.Id);
                }

                using (SubTransaction tx = new SubTransaction(doc))
                {
                    ConnectorSet PrimaryConnectors = null;
                    try
                    {
                        tx.Start();

                        List<ElementId> AlignedConduitsIds = new List<ElementId>();
                        foreach (ElementId cond1 in ConduitCollection1)
                        {
                            Conduit Conduit1 = doc.GetElement(cond1) as Conduit;
                            Line ConduitLine1 = (Conduit1.Location as LocationCurve).Curve as Line;
                            XYZ pt1 = ConduitLine1.GetEndPoint(0);
                            XYZ pt2 = ConduitLine1.GetEndPoint(1);
                            XYZ referenecdirection = ConduitLine1.Direction;

                            foreach (ElementId cond2 in ConduitCollection2)
                            {
                                Conduit conduit2 = doc.GetElement(cond2) as Conduit;
                                Line ConduitLine2 = (conduit2.Location as LocationCurve).Curve as Line;
                                XYZ pt3 = ConduitLine2.GetEndPoint(0);
                                XYZ pt4 = ConduitLine2.GetEndPoint(1);

                                Line referenceline = Line.CreateBound(pt1, pt3);
                                XYZ referencelinedirectionsub = referenceline.Direction;
                                List<double> distanecollection = new List<double>();

                                if (Math.Abs(Math.Round(referenecdirection.X, 5)) == Math.Abs(Math.Round(referencelinedirectionsub.X, 5)) && Math.Abs(Math.Round(referenecdirection.Y, 5)) == Math.Abs(Math.Round(referencelinedirectionsub.Y, 5)) && Math.Abs(Math.Round(referenecdirection.Z, 5)) == Math.Abs(Math.Round(referencelinedirectionsub.Z, 5)))
                                {
                                    Utility.AutoRetainParameters(Conduit1, conduit2, doc, uiapp);
                                    double firstpointdistence = Math.Sqrt(Math.Pow(pt1.X - pt3.X, 2) + Math.Pow(pt1.Y - pt3.Y, 2));
                                    distanecollection.Add(firstpointdistence);
                                    double secondpointdistence = Math.Sqrt(Math.Pow(pt1.X - pt4.X, 2) + Math.Pow(pt1.Y - pt4.Y, 2));
                                    distanecollection.Add(secondpointdistence);
                                    double thirdpointdistance = Math.Sqrt(Math.Pow(pt2.X - pt3.X, 2) + Math.Pow(pt2.Y - pt3.Y, 2));
                                    distanecollection.Add(thirdpointdistance);
                                    double fourthpointdistance = Math.Sqrt(Math.Pow(pt2.X - pt4.X, 2) + Math.Pow(pt2.Y - pt4.Y, 2));
                                    distanecollection.Add(fourthpointdistance);

                                    double maxiumdistance = distanecollection.Max<double>();
                                    if (maxiumdistance == firstpointdistence)
                                    {
                                        PrimaryConnectors = Utility.GetConnectors(conduit2 as Element);
                                        (Conduit1.Location as LocationCurve).Curve = Line.CreateBound(pt1, pt3);
                                        FamilyInstance conduitfittings = Utility.GetFittingByConduit(doc, conduit2, null, "elbow");
                                        FamilyInstance conduitcoupling = Utility.GetFamilyByConduit(doc, conduit2, "union");
                                        FamilyInstance conduitfittings2 = Utility.GetFittingByConduit(doc, Conduit1, null, "elbow");
                                        FamilyInstance conduitcoupling2 = Utility.GetFamilyByConduit(doc, Conduit1, "union");
                                        if (conduitfittings != null)
                                        {
                                            Utility.Connect(pt3, conduitfittings, Conduit1);
                                        }
                                        else if (conduitcoupling != null)
                                        {
                                            Utility.Connect(pt3, conduitcoupling, Conduit1);
                                        }
                                    }
                                    else if (maxiumdistance == secondpointdistence)
                                    {
                                        (Conduit1.Location as LocationCurve).Curve = Line.CreateBound(pt1, pt4);
                                        FamilyInstance conduitfittings = Utility.GetFittingByConduit(doc, conduit2, null, "elbow");
                                        FamilyInstance conduitcoupling = Utility.GetFamilyByConduit(doc, conduit2, "union");
                                        FamilyInstance conduitfittings2 = Utility.GetFittingByConduit(doc, Conduit1, null, "elbow");
                                        FamilyInstance conduitcoupling2 = Utility.GetFamilyByConduit(doc, Conduit1, "union");

                                        if (conduitfittings != null)
                                        {
                                            Utility.Connect(pt4, conduitfittings, Conduit1);
                                        }
                                        else if (conduitcoupling != null)
                                        {
                                            Utility.Connect(pt4, conduitcoupling, Conduit1);
                                        }

                                    }
                                    else if (maxiumdistance == thirdpointdistance)
                                    {
                                        (Conduit1.Location as LocationCurve).Curve = Line.CreateBound(pt3, pt2);
                                        FamilyInstance conduitfittings = Utility.GetFittingByConduit(doc, conduit2, null, "elbow");
                                        FamilyInstance conduitcoupling = Utility.GetFamilyByConduit(doc, conduit2, "union");
                                        FamilyInstance conduitfittings2 = Utility.GetFittingByConduit(doc, Conduit1, null, "elbow");
                                        FamilyInstance conduitcoupling2 = Utility.GetFamilyByConduit(doc, Conduit1, "union");

                                        if (conduitfittings != null)
                                        {
                                            Utility.Connect(pt3, conduitfittings, Conduit1);
                                        }
                                        else if (conduitcoupling != null)
                                        {
                                            Utility.Connect(pt3, conduitcoupling, Conduit1);
                                        }


                                    }
                                    else if (maxiumdistance == fourthpointdistance)
                                    {
                                        (Conduit1.Location as LocationCurve).Curve = Line.CreateBound(pt4, pt2);
                                        FamilyInstance conduitfittings = Utility.GetFittingByConduit(doc, conduit2, null, "elbow");
                                        FamilyInstance conduitcoupling = Utility.GetFamilyByConduit(doc, conduit2, "union");
                                        FamilyInstance conduitfittings2 = Utility.GetFittingByConduit(doc, Conduit1, null, "elbow");
                                        FamilyInstance conduitcoupling2 = Utility.GetFamilyByConduit(doc, Conduit1, "union");

                                        if (conduitfittings != null)
                                        {
                                            Utility.Connect(pt4, conduitfittings, Conduit1);
                                        }
                                        else if (conduitcoupling != null)
                                        {
                                            Utility.Connect(pt4, conduitcoupling, Conduit1);
                                        }
                                    }
                                    AlignedConduitsIds.Add(cond2);
                                }
                            }
                        }
                        AlignedConduitsIds = AlignedConduitsIds.Distinct().ToList();
                        foreach (ElementId eid in AlignedConduitsIds)
                        {
                            doc.Delete(eid);
                        }
                        if (ConduitCollection2.Count() != AlignedConduitsIds.Count())
                        {
                            TaskDialog.Show("Warning", "Couldn't connect all runs. Please check conduit alignment for failing elements.");
                        }
                        tx.Commit();
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Multi Extend", Util.ProductVersion);
                    }
                    catch
                    {
                        tx.RollBack();
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "RollBack", "Multi Extend", Util.ProductVersion);
                    }
                }
            }
            catch (Exception exception)
            {

                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Multi Extend", Util.ProductVersion);
            }
        }
        #endregion
        #region Rolling Offset
        public List<ElementId> RoffsetExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements, string l_direction)
        {
            startDate = DateTime.UtcNow;
            _uiapp = uiapp;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            double l_angle;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";


            double elevationOne = PrimaryElements[0].LookupParameter(offsetVariable).AsDouble();
            double elevationTwo = SecondaryElements[0].LookupParameter(offsetVariable).AsDouble();
            l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString());

            bool isRollUp = elevationOne < elevationTwo;

            List<ElementId> Unwantedids;
            if (isRollUp)
            {
                Unwantedids = RollUp(doc, uidoc, PrimaryElements, SecondaryElements, l_angle, l_direction, offsetVariable, uiapp);
            }
            else
            {
                Unwantedids = RollDown(doc, uidoc, PrimaryElements, SecondaryElements, l_angle, l_direction, offsetVariable, uiapp);
            }

            return Unwantedids;
        }

        private List<ElementId> RollUp(Document doc, UIDocument uidoc, List<Element> PrimaryElements, List<Element> SecondaryElements, double l_angle, string l_direction, string offsetVariable, UIApplication uiapp)
        {
            List<ElementId> unwantedIds = new List<ElementId>();
            Dictionary<double, List<Element>> groupedFirstElements = new Dictionary<double, List<Element>>();
            Dictionary<double, List<Element>> groupedSecondElements = new Dictionary<double, List<Element>>();
            Utility.GroupByElevation(PrimaryElements, offsetVariable, ref groupedFirstElements);
            Utility.GroupByElevation(SecondaryElements, offsetVariable, ref groupedSecondElements);

            int j = 0;
            foreach (KeyValuePair<double, List<Element>> valuePair in groupedFirstElements)
            {
                PrimaryElements = valuePair.Value.ToList();
                SecondaryElements = groupedSecondElements.Values.ElementAt(j).ToList();
                double zSpace = groupedFirstElements.FirstOrDefault().Key - valuePair.Key;
                Line refLine = (PrimaryElements[0].Location as LocationCurve).Curve as Line;
                XYZ refDirection = refLine.Direction;
                XYZ refCross = refDirection.CrossProduct(XYZ.BasisZ);
                Line perdicularLine = Line.CreateBound(refLine.Origin, refLine.Origin + refCross.Multiply(10));
                for (int i = 0; i < PrimaryElements.Count; i++)
                {
                    double elevationOne = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                    double elevationTwo = SecondaryElements[i].LookupParameter(offsetVariable).AsDouble();
                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                    Line lineOne = lc1.Curve as Line;
                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                    Line lineTwo = lc2.Curve as Line;
                    XYZ sectionPoint = Utility.FindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), perdicularLine.GetEndPoint(0), perdicularLine.GetEndPoint(1));
                    XYZ OriginWithOutZAxis = new XYZ(refLine.Origin.X, refLine.Origin.Y, 0);
                    double space = OriginWithOutZAxis.DistanceTo(sectionPoint);
                    double l_Angle = l_angle * Math.PI / 180;
                    space = Math.Tan(l_Angle / 2.5) * space;
                    zSpace = Math.Tan(l_Angle / 2) * zSpace;
                    ConnectorSet PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                    ConnectorSet SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                    XYZ primaryLineDirection = lineTwo.Direction;
                    XYZ cross = primaryLineDirection.CrossProduct(XYZ.BasisZ);
                    Line lineThree = Line.CreateBound(ConnectorTwo.Origin, ConnectorTwo.Origin + cross.Multiply(ConnectorOne.Origin.DistanceTo(ConnectorTwo.Origin)));
                    XYZ interSecPoint = Utility.FindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), lineThree.GetEndPoint(0), lineThree.GetEndPoint(1));

                    XYZ newenpt = new XYZ(interSecPoint.X, interSecPoint.Y, ConnectorOne.Origin.Z);
                    Line newLine = Line.CreateBound(ConnectorOne.Origin, newenpt);

                    XYZ newStartPoint = (l_direction.Contains("Left-Down") || l_direction.Contains("Right-Down") || l_direction.Contains("Top-Right") || l_direction.Contains("Bottom-Right")) ?
                        newLine.Origin - (newLine.Direction * (space + zSpace)) : newLine.Origin + (newLine.Direction * (space + zSpace));

                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, newStartPoint, newenpt);
                    newCon.LookupParameter(offsetVariable).Set(elevationOne);
                    XYZ direction = ((newCon.Location as LocationCurve).Curve as Line).Direction;
                    LocationCurve curve = newCon.Location as LocationCurve;
                    Curve line = curve.Curve;

                    //RetainParameters(PrimaryElements[i], SecondaryElements[i], doc);
                    //RetainParameters(PrimaryElements[i], newCon as Element, doc);

                    if (curve != null)
                    {
                        XYZ aa = newStartPoint;
                        XYZ cc = new XYZ(aa.X, aa.Y, aa.Z + 10);
                        Line axisLine = Line.CreateBound(aa, cc);
                        double l_offSet = elevationOne < elevationTwo ? (elevationTwo - elevationOne) : (elevationOne - elevationTwo);
                        XYZ EndPointwithoutZ = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, 0);
                        double l_rollOffset = EndPointwithoutZ.DistanceTo(interSecPoint);
                        double rollAngle = Math.Atan2(l_offSet, l_rollOffset);
                        if (l_direction.Contains("Left-Up") || l_direction.Contains("Right-Down")
                            || l_direction.Contains("Top-Right")
                            || l_direction.Contains("Bottom-Left"))
                        {
                            curve.Rotate(axisLine, -l_angle * (Math.PI / 180));
                            curve.Rotate(newLine, -rollAngle);
                        }
                        else
                        {
                            curve.Rotate(axisLine, l_angle * (Math.PI / 180));
                            curve.Rotate(newLine, rollAngle);
                        }
                    }
                    try
                    {
                        Element e = doc.GetElement(newCon.Id);
                        ConnectorSet ThirdConnectors = Utility.GetConnectors(e);
                        Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                        Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                        Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                        Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                    }
                    catch
                    {
                        unwantedIds.Add(newCon.Id);


                    }

                }
                j++;
            }

            return unwantedIds;
        }
        private List<ElementId> RollDown(Document doc, UIDocument uidoc, List<Element> PrimaryElements, List<Element> SecondaryElements, double l_angle, string l_direction, string offsetVariable, UIApplication uiapp)
        {
            List<ElementId> unwantedIds = new List<ElementId>();
            Dictionary<double, List<Element>> groupedFirstElements = new Dictionary<double, List<Element>>();
            Dictionary<double, List<Element>> groupedSecondElements = new Dictionary<double, List<Element>>();
            Utility.GroupByElevation(PrimaryElements, offsetVariable, ref groupedFirstElements);
            Utility.GroupByElevation(SecondaryElements, offsetVariable, ref groupedSecondElements);


            int j = 0;
            foreach (KeyValuePair<double, List<Element>> valuePair in groupedFirstElements)
            {
                PrimaryElements = valuePair.Value.ToList();
                SecondaryElements = groupedSecondElements.Values.ElementAt(j).ToList();
                double zSpace = groupedFirstElements.FirstOrDefault().Key - valuePair.Key;
                Line refLine = (PrimaryElements[0].Location as LocationCurve).Curve as Line;
                XYZ refDirection = refLine.Direction;
                XYZ refCross = refDirection.CrossProduct(XYZ.BasisZ);
                Line perdicularLine = Line.CreateBound(refLine.Origin, (refLine.Origin + refCross.Multiply(10)));
                for (int i = 0; i < PrimaryElements.Count; i++)
                {
                    double elevationOne = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                    double elevationTwo = SecondaryElements[i].LookupParameter(offsetVariable).AsDouble();
                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                    Line lineOne = lc1.Curve as Line;
                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                    Line lineTwo = lc2.Curve as Line;
                    XYZ sectionPoint = Utility.FindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), perdicularLine.GetEndPoint(0), perdicularLine.GetEndPoint(1));
                    XYZ OriginWithOutZAxis = new XYZ(refLine.Origin.X, refLine.Origin.Y, 0);
                    double space = OriginWithOutZAxis.DistanceTo(sectionPoint);
                    double l_Angle = l_angle * Math.PI / 180;
                    space = Math.Tan(l_Angle / 2.5) * space;
                    zSpace = Math.Tan(l_Angle / 2) * zSpace;
                    ConnectorSet PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                    ConnectorSet SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                    XYZ primaryLineDirection = lineTwo.Direction;
                    XYZ cross = primaryLineDirection.CrossProduct(XYZ.BasisZ);
                    Line lineThree = Line.CreateBound(ConnectorTwo.Origin, ConnectorTwo.Origin + cross.Multiply(ConnectorOne.Origin.DistanceTo(ConnectorTwo.Origin)));
                    XYZ interSecPoint = Utility.FindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), lineThree.GetEndPoint(0), lineThree.GetEndPoint(1));

                    XYZ newenpt = new XYZ(interSecPoint.X, interSecPoint.Y, ConnectorOne.Origin.Z);
                    Line newLine = Line.CreateBound(ConnectorOne.Origin, newenpt);

                    XYZ newStartPoint = (l_direction.Contains("Left-Down") || l_direction.Contains("Right-Down") || l_direction.Contains("Top-Right") || l_direction.Contains("Bottom-Right")) ?
                        newLine.Origin - (newLine.Direction * (space + zSpace)) : newLine.Origin + (newLine.Direction * (space + zSpace));

                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, newStartPoint, newenpt);
                    newCon.LookupParameter(offsetVariable).Set(elevationOne);
                    XYZ direction = ((newCon.Location as LocationCurve).Curve as Line).Direction;
                    LocationCurve curve = newCon.Location as LocationCurve;
                    Curve line = curve.Curve;

                    //RetainParameters(PrimaryElements[i], SecondaryElements[i], doc);
                    //RetainParameters(PrimaryElements[i], newCon as Element, doc);

                    if (curve != null)
                    {
                        XYZ aa = newStartPoint;
                        XYZ cc = new XYZ(aa.X, aa.Y, aa.Z + 10);
                        Line axisLine = Line.CreateBound(aa, cc);
                        double l_offSet = elevationOne < elevationTwo ? (elevationTwo - elevationOne) : (elevationOne - elevationTwo);
                        XYZ EndPointwithoutZ = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, 0);
                        double l_rollOffset = EndPointwithoutZ.DistanceTo(interSecPoint);
                        double rollAngle = Math.Atan2(l_offSet, l_rollOffset);
                        if (l_direction.Contains("Left-Up") || l_direction.Contains("Right-Down") || l_direction.Contains("Top-Right")
                            || l_direction.Contains("Bottom-Left"))
                        {
                            curve.Rotate(axisLine, -l_angle * (Math.PI / 180));
                            curve.Rotate(newLine, rollAngle);
                        }
                        else
                        {
                            curve.Rotate(axisLine, l_angle * (Math.PI / 180));
                            curve.Rotate(newLine, -rollAngle);
                        }
                    }
                    try
                    {
                        Element e = doc.GetElement(newCon.Id);
                        ConnectorSet ThirdConnectors = Utility.GetConnectors(e);
                        Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                        Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                        Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                        Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                    }
                    catch
                    {
                        unwantedIds.Add(newCon.Id);
                    }

                }
                j++;




            }

            return unwantedIds;
        }

        #endregion

        #region Kick

        public void KickExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements, int first)
        {

            DateTime startDate = DateTime.UtcNow;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
            try
            {

                if (first == 0)
                {
                    ElementsFilter filter = new ElementsFilter("Conduits");
                    Reference reference = uidoc.Selection.PickObject(ObjectType.Element, filter, "Please select the conduit in group to define 90 near and 90 far");
                    if (!PrimaryElements.Any(e => e.Id == doc.GetElement(reference.ElementId).Id))
                    {
                        var temp = PrimaryElements;
                        PrimaryElements = SecondaryElements;
                        SecondaryElements = temp;

                        _isfirst = true;
                    }
                }
                if (first > 0)
                {
                    if (_isfirst)
                    {
                        var temp = PrimaryElements;
                        PrimaryElements = SecondaryElements;
                        SecondaryElements = temp;
                    }
                }

                double l_angle;
                bool isUp = PrimaryElements.FirstOrDefault().LookupParameter(offsetVariable).AsDouble() <
                    SecondaryElements.FirstOrDefault().LookupParameter(offsetVariable).AsDouble();
                if (!isUp)
                {
                    l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                    try
                    {
                        using (SubTransaction tx = new SubTransaction(doc))
                        {
                            ConnectorSet PrimaryConnectors = null;
                            ConnectorSet SecondaryConnectors = null;
                            ConnectorSet ThirdConnectors = null;
                            tx.Start();
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                Line l1 = lc1.Curve as Line;
                                LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                Line l2 = lc2.Curve as Line;
                                XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                XYZ EndPoint = ConnectorTwo.Origin;
                                XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                newCon.LookupParameter(offsetVariable).Set(elevation);
                                Element e = doc.GetElement(newCon.Id);
                                LocationCurve newConcurve = newCon.Location as LocationCurve;
                                Line ncl1 = newConcurve.Curve as Line;
                                XYZ ncenpt = ncl1.GetEndPoint(1);
                                XYZ direction = ncl1.Direction;
                                XYZ midPoint = ncenpt;
                                XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                Line axisLine = Line.CreateBound(midPoint, midHigh);
                                newConcurve.Rotate(axisLine, -l_angle);
                                ThirdConnectors = Utility.GetConnectors(e);
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                            }
                            tx.Commit();
                            _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Kick With Bend Down", Util.ProductVersion, "Connect");
                        }
                    }
                    catch
                    {
                        try
                        {
                            using (SubTransaction tx = new SubTransaction(doc))
                            {
                                ConnectorSet PrimaryConnectors = null;
                                ConnectorSet SecondaryConnectors = null;
                                ConnectorSet ThirdConnectors = null;

                                tx.Start();
                                for (int i = 0; i < PrimaryElements.Count; i++)
                                {
                                    double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                    Line l1 = lc1.Curve as Line;
                                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                    Line l2 = lc2.Curve as Line;
                                    XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                    PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                    SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                    XYZ EndPoint = ConnectorTwo.Origin;
                                    XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                    newCon.LookupParameter(offsetVariable).Set(elevation);
                                    Element e = doc.GetElement(newCon.Id);
                                    LocationCurve newConcurve = newCon.Location as LocationCurve;
                                    Line ncl1 = newConcurve.Curve as Line;
                                    XYZ ncenpt = ncl1.GetEndPoint(1);
                                    XYZ direction = ncl1.Direction;
                                    XYZ midPoint = ncenpt;
                                    XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                    Line axisLine = Line.CreateBound(midPoint, midHigh);
                                    newConcurve.Rotate(axisLine, l_angle);
                                    ThirdConnectors = Utility.GetConnectors(e);
                                    Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                    Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                                }
                                tx.Commit();
                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Kick With Bend Down", Util.ProductVersion, "Connect");

                            }
                        }
                        catch (Exception exception)
                        {
                            System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                            _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Kick With Bend Down", Util.ProductVersion, "Connect");
                        }
                    }
                }
                if (isUp)
                {
                    l_angle = Convert.ToDouble(ParentUserControl.Instance.ddlAngle.SelectedItem.Name.ToString()) * (Math.PI / 180);
                    try
                    {

                        using (SubTransaction tx = new SubTransaction(doc))
                        {
                            ConnectorSet PrimaryConnectors = null;
                            ConnectorSet SecondaryConnectors = null;
                            ConnectorSet ThirdConnectors = null;

                            tx.Start();
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                Line l1 = lc1.Curve as Line;
                                LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                Line l2 = lc2.Curve as Line;
                                XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                XYZ EndPoint = ConnectorTwo.Origin;
                                XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                newCon.LookupParameter(offsetVariable).Set(elevation);
                                Element e = doc.GetElement(newCon.Id);
                                LocationCurve newConcurve = newCon.Location as LocationCurve;
                                Line ncl1 = newConcurve.Curve as Line;
                                XYZ ncenpt = ncl1.GetEndPoint(1);
                                XYZ direction = ncl1.Direction;
                                XYZ midPoint = ncenpt;
                                XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                Line axisLine = Line.CreateBound(midPoint, midHigh);
                                newConcurve.Rotate(axisLine, l_angle);
                                ThirdConnectors = Utility.GetConnectors(e);
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                            }
                            tx.Commit();
                            _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Kick With Bend Up", Util.ProductVersion, "Connect");


                        }
                    }
                    catch
                    {
                        try
                        {
                            using (SubTransaction tx = new SubTransaction(doc))
                            {
                                ConnectorSet PrimaryConnectors = null;
                                ConnectorSet SecondaryConnectors = null;
                                ConnectorSet ThirdConnectors = null;

                                tx.Start();
                                for (int i = 0; i < PrimaryElements.Count; i++)
                                {
                                    double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                    Line l1 = lc1.Curve as Line;
                                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                    Line l2 = lc2.Curve as Line;
                                    XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                    PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                    SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                    XYZ EndPoint = ConnectorTwo.Origin;
                                    XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                    newCon.LookupParameter(offsetVariable).Set(elevation);
                                    Element e = doc.GetElement(newCon.Id);
                                    LocationCurve newConcurve = newCon.Location as LocationCurve;
                                    Line ncl1 = newConcurve.Curve as Line;
                                    XYZ ncenpt = ncl1.GetEndPoint(1);
                                    XYZ direction = ncl1.Direction;
                                    XYZ midPoint = ncenpt;
                                    XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                    Line axisLine = Line.CreateBound(midPoint, midHigh);
                                    newConcurve.Rotate(axisLine, -l_angle);
                                    ThirdConnectors = Utility.GetConnectors(e);
                                    Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                    Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                                }
                                tx.Commit();
                                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Completed", "Kick With Bend Up", Util.ProductVersion, "Connect");
                            }
                        }
                        catch (Exception exception)
                        {
                            System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                            _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Kick With Bend Up", Util.ProductVersion, "Connect");
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, Util.ApplicationWindowTitle, startDate, "Failed", "Kick With Bend", Util.ProductVersion, "Connect");
            }
        }

        #endregion
        public string GetName()
        {
            return "AutoConnect";
        }

        public static Autodesk.Revit.DB.FamilyInstance CreateElbowFittings(ConnectorSet PrimaryConnectors, ConnectorSet SecondaryConnectors, Document Doc)
        {
            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out var ConnectorOne, out var ConnectorTwo);
            return Doc.Create.NewElbowFitting(ConnectorOne, ConnectorTwo);
        }

        public static Autodesk.Revit.DB.FamilyInstance CreateElbowFittings(Autodesk.Revit.DB.Element One, Autodesk.Revit.DB.Element Two, Document doc, Autodesk.Revit.UI.UIApplication uiApp)
        {
            ConnectorSet connectorSet = GetConnectorSet(One);
            ConnectorSet connectorSet2 = GetConnectorSet(Two);
            Utility.AutoRetainParameters(One, Two, doc, uiApp);
            return CreateElbowFittings(connectorSet, connectorSet2, doc);
        }
        public static ConnectorSet GetConnectorSet(Autodesk.Revit.DB.Element Ele)
        {
            ConnectorSet result = null;
            if (Ele is Autodesk.Revit.DB.FamilyInstance)
            {
                MEPModel mEPModel = ((Autodesk.Revit.DB.FamilyInstance)Ele).MEPModel;
                if (mEPModel != null && mEPModel.ConnectorManager != null)
                {
                    result = mEPModel.ConnectorManager.UnusedConnectors;
                }
            }
            else if (Ele is MEPCurve)
            {
                result = ((MEPCurve)Ele).ConnectorManager.UnusedConnectors;
            }

            return result;
        }


    }

}
