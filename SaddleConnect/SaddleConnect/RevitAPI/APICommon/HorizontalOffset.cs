using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TIGUtility;

namespace SaddleConnect
{
    public class HorizontalOffset
    {
        public static void GetSecondaryElements(Document doc, ref List<Element> primaryElements, double angle, double offSet, string offSetVar, out List<Element> secondaryElements,XYZ Pickpoint)
        {
            secondaryElements = new List<Element>();

            Line lineOne = (primaryElements[0].Location as LocationCurve).Curve as Line;
            XYZ pt1 = lineOne.GetEndPoint(0);
            XYZ pt2 = lineOne.GetEndPoint(1);
            XYZ midpoint = (pt1 + pt2) / 2;
            XYZ PrimaryConduitDirection = lineOne.Direction;
            XYZ CrossProduct = PrimaryConduitDirection.CrossProduct(XYZ.BasisZ);
            XYZ PickPointTwo = Pickpoint + CrossProduct.Multiply(1);

            XYZ Intersectionpoint = Utility.FindIntersectionPoint(pt1, pt2, Pickpoint, PickPointTwo);

            Line ConduitDirectionLine = Line.CreateBound(midpoint, Intersectionpoint);
            XYZ DirectionPfconduit = ConduitDirectionLine.Direction;
            DirectionPfconduit = new XYZ(DirectionPfconduit.X, DirectionPfconduit.Y, PrimaryConduitDirection.Z);

            //line for perpendicular dir
            Line lineforperpenduicular = Line.CreateBound(Intersectionpoint, Pickpoint);
            XYZ Perpendiculardir = lineforperpenduicular.Direction;
            Perpendiculardir = new XYZ(Perpendiculardir.X, Perpendiculardir.Y, PrimaryConduitDirection.Z);

            //Set Line direction
            XYZ pickpointst1 = Pickpoint + PrimaryConduitDirection.Multiply(1);
            XYZ midpoint2 = midpoint + CrossProduct.Multiply(1);
            XYZ intersectiompointTwo = Utility.FindIntersectionPoint(Pickpoint, pickpointst1, midpoint, midpoint2);
            Line Linefoeoffset = Line.CreateBound(midpoint,intersectiompointTwo);
            XYZ Directionforoffset = Linefoeoffset.Direction;

            secondaryElements = new List<Element>();
            double basedistance = (angle * (180 / Math.PI)) == 90 ? 1 : offSet / Math.Tan(angle);
            primaryElements = primaryElements.OrderByDescending(r => ((r.Location as LocationCurve).Curve as Line).Origin.Y).ToList();

            for (int i = 0; i < primaryElements.Count; i++)
            {
                LocationCurve curve = primaryElements[i].Location as LocationCurve;
                Line l_Line = curve.Curve as Line;
                XYZ StartPoint = l_Line.GetEndPoint(0);
                XYZ EndPoint = l_Line.GetEndPoint(1);
                double SubdistanceOne = Math.Sqrt(Math.Pow((StartPoint.X - Intersectionpoint.X), 2) + Math.Pow((StartPoint.Y - Intersectionpoint.Y), 2));
                double SubdistanceTwo = Math.Sqrt(Math.Pow((EndPoint.X - Intersectionpoint.X), 2) + Math.Pow((EndPoint.Y - Intersectionpoint.Y), 2));
                XYZ ConduitStartpt = null;
                XYZ ConduitEndpoint = null;
                if (SubdistanceOne < SubdistanceTwo)
                {
                    ConduitStartpt = StartPoint;
                    ConduitEndpoint = EndPoint;
                }
                else
                {
                    ConduitStartpt = EndPoint;
                    ConduitEndpoint = StartPoint;
                }
                Line ConduitLine = Line.CreateBound(ConduitEndpoint, ConduitStartpt);
                XYZ ConduitLinedir = ConduitLine.Direction;
                XYZ refStartPoint = ConduitStartpt + ConduitLinedir.Multiply(Math.Abs(basedistance));
                refStartPoint = refStartPoint + Perpendiculardir.Multiply(offSet);
                XYZ refEndPoint = refStartPoint + ConduitLinedir.Multiply(10);

                Conduit newCon = Utility.CreateConduit(doc, primaryElements[i] as Conduit, refStartPoint, refEndPoint);
                double elevation = primaryElements[i].LookupParameter(offSetVar).AsDouble();
                Parameter newElevation = newCon.LookupParameter(offSetVar);
                newElevation.Set(elevation);
                Element ele = doc.GetElement(newCon.Id);
                secondaryElements.Add(ele);
            }


            //switch (direction)
            //{
            //    case "Right":
            //        if (offSet > 0)
            //            RightUpOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        else
            //            RightDownOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        break;
            //    case "Left":
            //        if (offSet > 0)
            //            LeftUpOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        else
            //            LeftDownOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        break;
            //    case "Top":
            //        if (offSet > 0)
            //            TopRightOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        else
            //            TopLeftOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        break;
            //    case "Bottom":
            //        if (offSet > 0)
            //            BottomRightOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        else
            //            BottomLeftOffset(doc, ref primaryElements, conduitLength, angle, offSet, offSetVar, out secondaryElements);
            //        break;
            //    default:
            //        break;
            //}
        }

        
    }
}
