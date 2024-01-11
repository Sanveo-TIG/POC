using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using SaddleConnect;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Autodesk.Revit.UI.Selection;

namespace SaddleConnect
{
    public class APICommon
    {
        public static List<Element> GetElementsByOder(List<Element> a_PrimaryElements)
        {
            List<Element> PrimaryElements = new List<Element>();
            XYZ PrimaryDirection = ((a_PrimaryElements.FirstOrDefault().Location as LocationCurve).Curve as Line).Direction;
            if (Math.Abs(PrimaryDirection.Z) != 1)
            {
                PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).Y).ToList();
                if (PrimaryDirection.Y == 1 || PrimaryDirection.Y == -1)
                {
                    PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
                }
            }
            else
            {
                PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
            }
            return PrimaryElements;
        }

        public static List<Element> GetElementsByReference(IList<Reference> References, Document doc)
        {
            List<Element> Elements = new List<Element>();
            foreach (Reference r in References)
            {
                Element e = doc.GetElement(r);
                if (e.GetType() == typeof(Conduit))
                    Elements.Add(e);
            }
            return Elements;
        }

        public static List<Element> Getconduitsinselection(UIApplication uiapp)
        {
            List<Element> PrimaryElements = new List<Element>();
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Selection selection = uidoc.Selection;
            List<ElementId> eids = selection.GetElementIds().ToList();
            foreach (ElementId eid in eids)
            {
                Element e = doc.GetElement(eid);

                if (e.GetType() == typeof(Conduit))
                {
                    PrimaryElements.Add(e);
                }
            }
            return PrimaryElements;
        }

    }
}
