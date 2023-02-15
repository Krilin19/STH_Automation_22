#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Architecture;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using adWin = Autodesk.Windows;
using System.Linq;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml;
using System.Runtime.Serialization.Formatters;
using Autodesk.Revit.Creation;
using System.Windows;
using System.Windows.Media.Media3D;
using Document = Autodesk.Revit.Creation.Document;
using System.Data.Common;
using System.Net;
using static Autodesk.Internal.Windows.SwfMediaPlayer;
#endregion
namespace STH_Automation_22
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rot_Ele_Angle_To_Grid : IExternalCommand
    {
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-5609-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();
            List<Element> ele2 = new List<Element>();
            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("!", "Select a reference Grid to find orthogonal walls");

            Autodesk.Revit.DB.Grid Grid_ = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Grid;
            //Autodesk.Revit.DB.Wall el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Wall;
            Element el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));
            LocationCurve lc = el.Location as LocationCurve;
            Autodesk.Revit.DB.LocationPoint lp = null;
            XYZ LineWallDir;
            if (lc != null)
            {
                LineWallDir = (lc.Curve as Autodesk.Revit.DB.Line).Direction;
            }
            else
            {
                Location loc = el.Location;
                lp = loc as Autodesk.Revit.DB.LocationPoint;

                var fi = el as FamilyInstance;
                XYZ DIR = fi.FacingOrientation;
                XYZ DirEnd = DIR * (1100 / 304.8);
                LineWallDir = DirEnd;
            }

            XYZ direction = Grid_.Curve.GetEndPoint(0).Subtract(Grid_.Curve.GetEndPoint(1)).Normalize();

            double RadiansOfRotation = LineWallDir.AngleTo(direction/*poscross2*/);
            double AngleOfRotation = RadiansOfRotation * 180 / Math.PI;

            Autodesk.Revit.DB.Line axis = null;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Wall Section View");

                if (lc != null)
                {
                    axis = Autodesk.Revit.DB.Line.CreateUnbound(lc.Curve.Evaluate(0.5, true), XYZ.BasisZ);
                }
                else
                {
                    axis = Autodesk.Revit.DB.Line.CreateUnbound((lp.Point), XYZ.BasisZ);
                }

                XYZ ElementDir = XYZ.Zero;

                if (AngleOfRotation > 90)
                {
                    RadiansOfRotation = Math.PI - RadiansOfRotation;
                    AngleOfRotation = RadiansOfRotation * 180 / Math.PI;
                }

                if (LineWallDir.IsAlmostEqualTo(direction))
                {
                    goto end;
                }
                else
                {
                    el.Location.Rotate(axis, (RadiansOfRotation * -1));

                    if (lc != null)
                    {
                        ElementDir = ((el.Location as LocationCurve).Curve as Autodesk.Revit.DB.Line).Direction;
                        double RadiansOfRotationTest = ElementDir.AngleTo(direction);
                        double AngleOfRotationTest = RadiansOfRotationTest * 180 / Math.PI;
                    }
                    else
                    {
                        Location loc = el.Location;
                        lp = loc as Autodesk.Revit.DB.LocationPoint;

                        var fi = el as FamilyInstance;
                        XYZ DIR = fi.FacingOrientation;

                        ElementDir = DIR;
                        double RadiansOfRotationTest = ElementDir.AngleTo(direction);
                        double AngleOfRotationTest = RadiansOfRotationTest * 180 / Math.PI;
                    }



                    if (ElementDir.IsAlmostEqualTo(direction))
                    {
                        goto end;
                    }

                    XYZ Inverted = InvCoord(direction);
                    if (ElementDir.IsAlmostEqualTo(Inverted))
                    {
                        goto end;
                    }
                    else
                    {
                        el.Location.Rotate(axis, RadiansOfRotation);
                        double RadiansOfRotationTest2 = ElementDir.AngleTo(direction);
                        double AngleOfRotationTest2 = RadiansOfRotationTest2 * 180 / Math.PI;

                        el.Location.Rotate(axis, RadiansOfRotation);
                        double RadiansOfRotationTest3 = ElementDir.AngleTo(direction);
                        double AngleOfRotationTest3 = RadiansOfRotationTest3 * 180 / Math.PI;
                        goto end;
                    }
                }
            end:
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Set_Annotation_Crop : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            // Get selected elements from current document.;
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            string s = "You Picked:" + "\n";

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Set Annotation Crop size");

                foreach (ElementId selectedid in selectedIds)
                {
                    Autodesk.Revit.DB.View e = doc.GetElement(selectedid) as Autodesk.Revit.DB.View;

                    s += " View id = " + e.Id + "\n";

                    s += " View name = " + e.Name + "\n";

                    s += "\n";

                    ViewCropRegionShapeManager regionManager = e.GetCropRegionShapeManager();

                    regionManager.BottomAnnotationCropOffset = 0.01;

                    regionManager.LeftAnnotationCropOffset = 0.01;

                    regionManager.RightAnnotationCropOffset = 0.01;

                    regionManager.TopAnnotationCropOffset = 0.01;
                }
                {
                    TaskDialog.Show("Basic Element Info", s);
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Elevation_A : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            SketchPlane skplane = SketchPlane.Create(doc, plane);
            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Wall wall = null;
            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);
                wall = e as Wall;
                LocationCurve lc = wall.Location as LocationCurve;

                Autodesk.Revit.DB.Line line = lc.Curve as Autodesk.Revit.DB.Line;
                if (null == line)
                {
                    message = "Unable to retrieve wall location line.";

                    return Result.Failed;
                }

                XYZ pntCenter = line.Evaluate(0.5, true);
                XYZ normal = line.Direction.Normalize();
                XYZ dir = new XYZ(0, 0, 1);
                XYZ cross = normal.CrossProduct(dir * -1);
                XYZ pntEnd = pntCenter + cross.Multiply(2);
                XYZ poscross = normal.CrossProduct(dir);
                XYZ pospntEnd = pntCenter + poscross.Multiply(2);

                XYZ vect1 = line.Direction * (-1100 / 304.8);
                vect1 = vect1.Negate();
                XYZ vect2 = vect1 + line.Evaluate(0.5, true);

                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(line.Evaluate(0.5, true), vect2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pntCenter, pospntEnd);

                double angle4 = XYZ.BasisY.AngleTo(line2.Direction);
                double angleDegrees4 = angle4 * 180 / Math.PI;
                if (pntCenter.X < line2.GetEndPoint(1).X)
                {
                    angle4 = 2 * Math.PI - angle4;
                }
                double angleDegreesCorrected4 = angle4 * 180 / Math.PI;
                Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(pntEnd, XYZ.BasisZ);

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Create Wall Section View");

                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, pntEnd/*line.Evaluate(0.5,true)*/, 100);
                    ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, 1);

                    if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                    {
                        angle4 = angle4 / 2;

                        marker.Location.Rotate(axis, angle4);
                    }
                    marker.Location.Rotate(axis, angle4);
                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Elevation_B : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            SketchPlane skplane = SketchPlane.Create(doc, plane);
            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Wall wall = null;
            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);
                wall = e as Wall;
                LocationCurve lc = wall.Location as LocationCurve;

                Autodesk.Revit.DB.Line line = lc.Curve as Autodesk.Revit.DB.Line;
                if (null == line)
                {
                    message = "Unable to retrieve wall location line.";

                    return Result.Failed;
                }

                XYZ pntCenter = line.Evaluate(0.5, true);
                XYZ normal = line.Direction.Normalize();
                XYZ dir = new XYZ(0, 0, 1);
                XYZ cross = normal.CrossProduct(dir );
                XYZ pntEnd = pntCenter + cross.Multiply(2);
                XYZ poscross = normal.CrossProduct(dir);
                XYZ pospntEnd = pntCenter + poscross.Multiply(2);

                XYZ vect1 = line.Direction * (-1100 / 304.8);
                vect1 = vect1.Negate();
                XYZ vect2 = vect1 + line.Evaluate(0.5, true);

                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(line.Evaluate(0.5, true), vect2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pntCenter, pospntEnd);

                double angle4 = XYZ.BasisY.AngleTo(line2.Direction);
                double angleDegrees4 = angle4 * 180 / Math.PI;
                if (pntCenter.X < line2.GetEndPoint(1).X)
                {
                    angle4 = 2 * Math.PI - angle4;
                }
                double angleDegreesCorrected4 = angle4 * 180 / Math.PI;
                Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(pntEnd, XYZ.BasisZ);

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Create Wall Section View");
                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, pntEnd/*line.Evaluate(0.5,true)*/, 100);
                    ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, 1);

                    if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                    {
                        angle4 = angle4 / 2;
                        marker.Location.Rotate(axis, angle4 );
                    }
                    marker.Location.Rotate(axis, angle4 );
                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }
    }


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Angle_To_EleMarker : IExternalCommand
    {

        public XYZ GetElevationMarkerCenter(Autodesk.Revit.DB.Document doc, Autodesk.Revit.DB.View symbolView, ElevationMarker marker, ViewSection elevation)
        {
            // 1/4" elevation marker radius.
            double elevationMarkerRadius = .02083333;
            double scaledElevationMarkerRadius = Math.Round(elevationMarkerRadius * symbolView.Scale, 6);

            // Offset from the centre of the bounding box for the viewer marker and the edge of the elevation marker circle.
            double offsetFromViewerEdge = 0.006102;
            double scaledOffsetFromViewer = Math.Round(offsetFromViewerEdge * symbolView.Scale, 6);
            XYZ center = null;
            // Get list of viewers in view.
            List<Element> viewers = new FilteredElementCollector(doc, symbolView.Id)
                .OfCategory(BuiltInCategory.OST_Viewers)
                .Where(v => v.Name == elevation.Name)
                .ToList();
            // Checking specific viewer on marker in case of reference views.
            foreach (Element e in viewers)
            {
                if (IsViewerInMarker(doc, marker, e))
                {
                    BoundingBoxXYZ vbb = e.get_BoundingBox(symbolView);
                    center = (vbb.Max + vbb.Min) / 2;
                    break;
                }
            }
            if (center != null)
            {
                center = center + (elevation.ViewDirection * (scaledElevationMarkerRadius - scaledOffsetFromViewer));
            }
            return center;
        }
        private bool IsViewerInMarker(Autodesk.Revit.DB.Document doc, ElevationMarker marker, Element viewer)
        {
            for (int i = 0; i < 4; i++)
            {
                ElementId viewSectionId = marker.GetViewId(i);
                if (viewSectionId != null)
                {
                    Element viewSection = doc.GetElement(viewSectionId);
                    if (viewSection?.Name == viewer.Name)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            SketchPlane skplane = SketchPlane.Create(doc, plane);
            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-5609-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            XYZ DIR = null;
            XYZ LineWallDir = null;
            XYZ DirEnd;
            double angle3 ;

            TaskDialog.Show("!", "Select a Wall and an elevation");

            //Autodesk.Revit.DB.Grid Grid_ = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Grid;
            ////Autodesk.Revit.DB.Wall el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Wall;
            Element el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));
            Element el2 = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));
            Element ViewSection_ = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));

            //List<Element> viewers = new FilteredElementCollector(doc, doc.ActiveView.Id)
            //   .OfCategory(BuiltInCategory.OST_Viewers)
            //   .Where(v => v.Name == ViewSection_.Name)
            //   .ToList();

            IEnumerable<Autodesk.Revit.DB.View> views = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v1 => !v1.IsTemplate)
                .Where(v1 => v1.CanUseTemporaryVisibilityModes())
                .Where(v1 => v1.Name == ViewSection_.Name)
                .ToList();

            Autodesk.Revit.DB.LocationPoint lp = null;
            ElevationMarker marker = el2 as ElevationMarker;
            LocationCurve lc = el.Location as LocationCurve;
            
            if (lc != null)
            {
                LineWallDir = (lc.Curve as Autodesk.Revit.DB.Line).Direction;
            }
            else
            {
                Location loc = el.Location;
                lp = loc as Autodesk.Revit.DB.LocationPoint;

                var fi = el as FamilyInstance;
                DIR = fi.FacingOrientation;
                DirEnd = DIR * (1100 / 304.8);
            }


            XYZ pntCenter = lc.Curve.Evaluate(0.5, true);
            XYZ normal = (lc.Curve as Autodesk.Revit.DB.Line).Direction;

            XYZ cross = normal.CrossProduct(XYZ.BasisZ * -1);
            XYZ pntEnd = pntCenter + cross.Multiply(2);
            XYZ poscross = normal.CrossProduct(XYZ.BasisZ);
            XYZ pospntEnd = pntCenter + poscross.Multiply(2);

            XYZ vect1 = (lc.Curve as Autodesk.Revit.DB.Line).Direction * (-1100 / 304.8);
            vect1 = vect1.Negate();
            XYZ vect2 = vect1 + lc.Curve.Evaluate(0.5, true);

            Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(lc.Curve.Evaluate(0.5, true), vect2);
            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pntCenter, pospntEnd);
            //Makeline(doc, pntCenter, vect2);

            double angle4 = XYZ.BasisY.AngleTo(line2.Direction);
            double angleDegrees4 = angle4 * 180 / Math.PI;
            if (pntCenter.X < line2.GetEndPoint(1).X)
            {
                angle4 = 2 * Math.PI - angle4;
            }
            double angleDegreesCorrected4 = angle4 * 180 / Math.PI;
            

            
            XYZ CentreEleMarker = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
            BoundingBoxXYZ vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
            XYZ center = (vbb.Max + vbb.Min) / 2;
            Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(center, XYZ.BasisZ);

            Autodesk.Revit.DB.Line EleMarkLine = Autodesk.Revit.DB.Line.CreateBound(CentreEleMarker, center);
            double EleMarkLineAgleToY  = XYZ.BasisY.AngleTo(EleMarkLine.Direction);
            double EleMarkLineAgleToYDegrees = EleMarkLineAgleToY * 180 / Math.PI;
           

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Wall Section View");

                if (EleMarkLine.Direction.X > 0)
                {
                    marker.Location.Rotate(axis, EleMarkLineAgleToY);
                }
                else
                {
                    marker.Location.Rotate(axis, EleMarkLineAgleToY * -1);
                }
                if (!line2.Direction.IsAlmostEqualTo(line3.Direction))
                {

                }

                if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                {
                    angle4 = angle4 / 2;

                    marker.Location.Rotate(axis, angle4);
                }
                marker.Location.Rotate(axis, angle4);
                tx.Commit();
            }



            //using (Transaction tx = new Transaction(doc))
            //{
            //    tx.Start("Create Wall Section View");
            //    if (marker != null)
            //    {
            //        BoundingBoxXYZ vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
            //        XYZ center = (vbb.Max + vbb.Min) / 2;

            //        XYZ norm = LineWallDir.CrossProduct(XYZ.BasisZ);

            //        XYZ InterNorm = LineWallDir.CrossProduct(XYZ.BasisZ.Negate());

            //        Autodesk.Revit.DB.Line line_ = Autodesk.Revit.DB.Line.CreateUnbound(lc.Curve.Evaluate(0.5, true), norm);
            //        XYZ vect1 = line_.Direction * (1100 / 304.8);
            //        XYZ PerpVect = vect1 + lc.Curve.Evaluate(0.5, true);
            //        ModelCurve mc = Makeline(doc, lc.Curve.Evaluate(0.5, true), PerpVect);
            //        Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(lc.Curve.Evaluate(0.5, true), PerpVect);

            //        Autodesk.Revit.DB.Line Interline_ = Autodesk.Revit.DB.Line.CreateUnbound(lc.Curve.Evaluate(0.5, true), InterNorm);
            //        XYZ Intervect1 = Interline_.Direction * (1100 / 304.8);
            //        XYZ InterPerpVect = Intervect1 + lc.Curve.Evaluate(0.5, true);
            //        ModelCurve Intermc = Makeline(doc, lc.Curve.Evaluate(0.5, true), InterPerpVect);
            //        Autodesk.Revit.DB.Line Interline3 = Autodesk.Revit.DB.Line.CreateBound(lc.Curve.Evaluate(0.5, true), InterPerpVect);

            //        XYZ vect2 = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
            //        Makeline(doc, vect2, center);
            //        Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(vect2, center);
                    
            //        angle3 = line2.Direction.AngleTo(line3.Direction);


            //        Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(vect2, XYZ.BasisZ);

            //        if (lc != null)
            //        {
            //            double AngleOfRotation = angle3 * 180 / Math.PI;

            //            if (AngleOfRotation > 180)
            //            {
            //                double PrevAngle = angle3;
            //                marker.Location.Rotate(axis, (angle3));

            //                vect2 = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
            //                vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
            //                center = (vbb.Max + vbb.Min) / 2;

            //                line2 = Autodesk.Revit.DB.Line.CreateBound(vect2, center);

            //                angle3 = line2.Direction.AngleTo(line3.Direction);
            //                AngleOfRotation = angle3 * 180 / Math.PI;

            //                XYZ Inverted = InvCoord(line3.Direction);
            //                if (line2.Direction.IsAlmostEqualTo(line3.Direction) || line2.Direction.IsAlmostEqualTo(Inverted))
            //                {
            //                    if (true)
            //                    {

            //                    }

            //                    goto end;
            //                }
            //                else 
            //                {
            //                    marker.Location.Rotate(axis, PrevAngle * -1);
            //                }
            //            }
            //            else
            //            {
            //                angle3 = Math.PI - angle3;
            //                AngleOfRotation = angle3 * 180 / Math.PI;
            //            }
            //            if (line2.Direction.IsAlmostEqualTo(line3.Direction))
            //            {
            //                goto end;
            //            }
            //            else
            //            {
            //                marker.Location.Rotate(axis, (angle3 ));

            //                vect2 = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
            //                vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
            //                center = (vbb.Max + vbb.Min) / 2;

            //                line2 = Autodesk.Revit.DB.Line.CreateBound(vect2, center);

            //                angle3 = line2.Direction.AngleTo(line3.Direction);
            //                AngleOfRotation = angle3 * 180 / Math.PI;

            //                if (line2.Direction.IsAlmostEqualTo(line3.Direction))
            //                {
            //                    goto end;
            //                }
            //                XYZ Inverted = InvCoord(line3.Direction);
            //                if (line2.Direction.IsAlmostEqualTo(Inverted))
            //                {
            //                    marker.Location.Rotate(axis, angle3 * -1);
            //                    goto end;
            //                }
            //                else
            //                {
            //                    marker.Location.Rotate(axis, angle3 * -1);
            //                    //marker.Location.Rotate(axis, angle3 * -1);;
            //                    goto end;
            //                }
            //            }
            //        }
            //    }
            //end:
            //    tx.Commit();
            //}
            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Find_Perpendicular_Wall : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();
            List<Element> ele2 = new List<Element>();
            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("!", "Select a reference Grid to find orthogonal walls");

            Autodesk.Revit.DB.Grid levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Grid;

            Autodesk.Revit.DB.Curve dircurve = levelBelow.Curve;
            Autodesk.Revit.DB.Line line = dircurve as Autodesk.Revit.DB.Line;
            XYZ dir = line.Direction;

            XYZ origin = line.Origin;
            XYZ viewdir = line.Direction;
            XYZ up = XYZ.BasisZ;
            XYZ right = up.CrossProduct(viewdir);

            foreach (Element wall in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Transform curveTransform = lc.Curve.ComputeDerivatives(0.5, true);

                try
                {
                    XYZ origin2 = curveTransform.Origin;
                    XYZ viewdir2 = curveTransform.BasisX.Normalize();
                    XYZ viewdir2_back = curveTransform.BasisX.Normalize() * -1;

                    XYZ up2 = XYZ.BasisZ;
                    XYZ right2 = up.CrossProduct(viewdir2);
                    XYZ left2 = up.CrossProduct(viewdir2 * -1);

                    double y_onverted = Math.Round(-1 * viewdir2.X);

                    if (viewdir.IsAlmostEqualTo(right2/*, 0.3333333333*/))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(left2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2_back))
                    {
                        ele.Add(wall);
                    }


                }
                catch (Exception)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }
            uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Export_View_Details : IExternalCommand
    {
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            SketchPlane skplane = SketchPlane.Create(doc, plane);
            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }


        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            string Sync_Manager = @"T:\Lopez\1.xlsx";
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            using (Transaction tx2 = new Transaction(doc))
            {
                tx2.Start("Create Wall Section View");
                foreach (ElementId selectedid in selectedIds)
                {
                    Autodesk.Revit.DB.View e = doc.GetElement(selectedid) as Autodesk.Revit.DB.View;

                    BoundingBoxXYZ box = e.get_BoundingBox(/*e*/null);
                    Autodesk.Revit.DB.Transform transform = box.Transform;

                    XYZ min = box.Min;
                    XYZ max = box.Max;

                    XYZ symBoxBL = box.Min;  
                    XYZ symBoxTR = box.Max;  
                    XYZ symBoxTL = new XYZ(symBoxBL.X, symBoxTR.Y, box.Min.Z);
                    XYZ symBoxBR = new XYZ(symBoxTR.X, symBoxBL.Y, box.Min.Z);

                    XYZ symBoxA = new XYZ(symBoxTR.X, symBoxTR.Y, box.Min.Z);

                    XYZ coordBL = transform.OfPoint(symBoxBL);  // 1) BL = bottom left 
                    XYZ coordTR = transform.OfPoint(symBoxTR);  // 2) TR = top right    
                    XYZ coordTL = transform.OfPoint(symBoxTL);  // 3) TL = top left
                    XYZ coordBR = transform.OfPoint(symBoxBR);  // 4) BL = bottom right

                    XYZ coordB = new XYZ(coordTR.X, coordTR.Y, coordBL.Z);
                    XYZ coordA = new XYZ(coordBL.X, coordTR.Y, coordBL.Z);

                    XYZ coordC = new XYZ(coordTR.X, coordBL.Y, coordBL.Z);
                    XYZ coordD = new XYZ(coordBL.X, coordTR.Y, coordBL.Z);

                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(coordA, coordB);
                    //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(coordC, coordD);

                    XYZ walldir = line.Direction ;
                    XYZ up = XYZ.BasisZ;
                    XYZ viewdir = walldir.CrossProduct(up);
                    double distance = coordA.DistanceTo(coordBL);
                    Autodesk.Revit.DB.Line lineAperp = Autodesk.Revit.DB.Line.CreateUnbound(coordB, viewdir);

                    XYZ vect1 = lineAperp.Direction * (distance *-1 /*/ 304.8*/);
                    XYZ vect2 = vect1 + coordB;
                    Autodesk.Revit.DB.Line lineB = Autodesk.Revit.DB.Line.CreateBound(coordB, vect2);
                    ModelCurve bot_left = Makeline(doc, coordB, vect2);

                    Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                    t.Origin = vect2;
                    t.BasisX = walldir;
                    t.BasisY = up;
                    t.BasisZ = viewdir;

                    BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                    sectionBox.Transform = t;
                    sectionBox.Min = min;
                    sectionBox.Max = max;


                    //Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                    ////t.Origin = coordBR;
                    ////t.BasisX = InvCoord(transform.BasisX);
                    ////t.BasisY = transform.BasisY;
                    ////t.BasisZ = InvCoord(transform.BasisZ);
                    //t.Origin = transform.Origin;
                    //t.BasisX = transform.BasisX;
                    //t.BasisY = transform.BasisY;
                    //t.BasisZ = transform.BasisZ;

                    //BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                    //sectionBox.Transform = t;
                    //sectionBox.Min = min;
                    //sectionBox.Max = max;

                    ModelCurve New_cen_max = Makeline(doc, transform.OfPoint(sectionBox.Min), transform.OfPoint(sectionBox.Max));
                    ModelCurve bot_left2 = Makeline(doc, coordC, coordD);

                    ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
                    ViewSection.CreateSection(doc, vft.Id, sectionBox /*box*/);



                    //Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(coordBL, coordBR);

                    //BoundingBoxXYZ Cbox = e.CropBox;
                    //Autodesk.Revit.DB.Transform Cboxtransform = Cbox.Transform;
                    //sectionBox.Max = transform.OfPoint(box.Max);
                    //sectionBox.Min = transform.OfPoint(box.Min);
                    //t.BasisX = transform.BasisX; //Rightdir
                    //t.BasisY = transform.BasisY; //up
                    //t.BasisZ = transform.BasisX.CrossProduct(transform.BasisY); //viewdir
                    //double TransX = transform.Origin.X;
                    //double TransY = transform.Origin.Y;
                    //double TransZ = transform.Origin.Z;
                    //XYZ TransBasisX = InvCoord(transform.BasisX);
                    //XYZ TransBasisY = InvCoord(transform.BasisY);
                    //XYZ TransBasisZ = InvCoord(transform.BasisZ);

                    //ModelCurve min_max = Makeline(doc, transform.OfPoint(box.Min), transform.OfPoint(box.Max));
                    //ModelCurve cen_max = Makeline(doc, transform.Origin, transform.OfPoint(box.Max));

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        //ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                        //sheet.Cells[1, 1].Value = TransX;
                        //sheet.Cells[1, 2].Value = TransY;
                        //sheet.Cells[1, 3].Value = TransZ;
                        //sheet.Cells[2, 1].Value = Math.Round(TransBasisX.X, 0);
                        //sheet.Cells[2, 2].Value = Math.Round(TransBasisX.Y, 0);
                        //sheet.Cells[2, 3].Value = Math.Round(TransBasisX.Z, 0);
                        //sheet.Cells[2, 4].Value = Math.Round(TransBasisY.X, 0);
                        //sheet.Cells[2, 5].Value = Math.Round(TransBasisY.Y, 0);
                        //sheet.Cells[2, 6].Value = Math.Round(TransBasisY.Z, 0);
                        //sheet.Cells[2, 7].Value = Math.Round(TransBasisZ.X, 0);
                        //sheet.Cells[2, 8].Value = Math.Round(TransBasisZ.Y, 0);
                        //sheet.Cells[2, 9].Value = Math.Round(TransBasisZ.Z, 0);
                        //sheet.Cells[3, 1].Value = min.X;
                        //sheet.Cells[3, 2].Value = min.Y;
                        //sheet.Cells[3, 3].Value = min.Z;
                        //sheet.Cells[3, 4].Value = max.X;
                        //sheet.Cells[3, 5].Value = max.Y;
                        //sheet.Cells[3, 6].Value = max.Z;
                        
                        package.Save();
                    }
                }
                tx2.Commit();
            } 
            {
                //TaskDialog.Show("Basic Element Info", s);
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Import_View_Details : IExternalCommand
    {
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            string Sync_Manager = @"T:\Lopez\1.xlsx";


            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                //s += " transform.Origin = " + transform.Origin + "\n";
                //s += " right = " + transform.BasisX + "\n";
                //s += " up = " + transform.BasisY + "\n";
                //s += " viewdir = " + transform.BasisZ + "\n";
                //s += "\n";

                //int number = Convert.ToInt32(sheet.Cells[2, column].Value);

                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                t.Origin = new XYZ((double)sheet.Cells[1, 1].Value, (double)sheet.Cells[1, 2].Value, (double)sheet.Cells[1, 3].Value);

                XYZ x_ = new XYZ((double)sheet.Cells[2, 1].Value, (double)sheet.Cells[2, 2].Value, (double)sheet.Cells[2, 3].Value) ;
                XYZ y_ = new XYZ((double)sheet.Cells[2, 4].Value, (double)sheet.Cells[2, 5].Value, (double)sheet.Cells[2, 6].Value);
                XYZ z_ = new XYZ((double)sheet.Cells[2, 7].Value, (double)sheet.Cells[2, 8].Value, (double)sheet.Cells[2, 9].Value) ;

                t.BasisX = x_;
                t.BasisY = y_;
                t.BasisZ = z_;

                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                //sectionBox.Transform = t;
                //sectionBox.Transform.Origin = new XYZ((double)sheet.Cells[1, 1].Value, (double)sheet.Cells[1, 2].Value, (double)sheet.Cells[1, 3].Value);
                sectionBox.Min = new XYZ((double)sheet.Cells[3, 1].Value, (double)sheet.Cells[3, 2].Value, (double)sheet.Cells[3, 3].Value);
                sectionBox.Max = new XYZ((double)sheet.Cells[3, 4].Value, (double)sheet.Cells[3, 5].Value, (double)sheet.Cells[3, 6].Value);

                using (Transaction tx2 = new Transaction(doc))
                {
                    tx2.Start("Create Wall Section View");

                    ViewSection.CreateSection(doc, vft.Id, sectionBox);

                    tx2.Commit();
                }


                package.Save();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Sync_Manager : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        static public Autodesk.Revit.ApplicationServices.Application m_app;
        public float abort = 0;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            IList<UIView> openViews = uidoc.GetOpenUIViews();
            foreach (UIView uiv in openViews)
            {
                if (uiv.ViewId != uidoc.ActiveView.Id)
                {
                    uiv.Close();
                }
            }

            SyncListUpdater SyncListUpdater_ = new SyncListUpdater();
            SyncListUpdater_.label3.Text = "Checking in 30 seconds";
            SyncListUpdater_.label4.Text = "Logged on at:";
            SyncListUpdater_.label5.Text = "Waiting list to sync";

            SyncListUpdater_.ShowDialog();


            if (doc.Title == "RAC_basic_sample_project"  /*"RHR_BUILDING_A22"*/)
            {
               
                string user = doc.Application.Username;
                var lastSaveTime = DateTime.Now;
                var CheckTime = DateTime.Now;
               

            beggining:
                string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";

                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show("Excel Sync Manager file not found, try Sync the normal way", "Sync Warning");
                    MessageBox.Show("Another user is using the Excel Sync Manager file, Try again" ,"Sync Warning");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                        var Time_ = DateTime.Now;

                        if (sheet.Cells[1, 2].Value != null)
                        {
                            if (sheet.Cells[1, 2].Value.ToString() == user)
                            {
                                goto finish_;
                            }

                        }
                        for (int row = 1; row < 20; row++)
                        {
                            if (sheet.Cells[row, 1].Value == null)
                            {
                                break;
                            }

                            if (sheet.Cells[row, 1].Value != null)
                            {
                                var Value1 = sheet.Cells[row, 1].Value;
                                var Value2 = sheet.Cells[row, 2].Value;
                                //s += Value1 + " + " + Value2.ToString() + "\n";
                                SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString() + "\n");
                                if ((sheet.Cells[row, 1].Value == null))
                                {
                                    sheet.Cells[1, 1].Value = Time_.ToString();
                                    sheet.Cells[1, 2].Value = user;
                                }
                            }

                        }
                        SyncListUpdater_.listBox1.Items.Clear();
                        SyncListUpdater_.listBox1.Refresh();
                        SyncListUpdater_.listBox1.Update();
                        SyncListUpdater_.Show();
                        SyncListUpdater_.label2.Text = lastSaveTime.ToString();
                        SyncListUpdater_.label2.Refresh();
                        SyncListUpdater_.label2.Update();
                        SyncListUpdater_.label3.Update();
                        SyncListUpdater_.label4.Update();
                        SyncListUpdater_.label5.Update();

                    //TaskDialog.Show("Current Sync List ", s);
                    //MessageBox.Show("Current Sync List ", "");
                    nonvalue:
                        //---------------------------------------------------------
                        try
                        {
                            if (sheet.Cells[1, 1].Value == null)
                            {
                                sheet.Cells[1, 1].Value = Time_.ToString();
                                sheet.Cells[1, 2].Value = user;
                                SyncListUpdater_.listBox1.Refresh();
                                SyncListUpdater_.listBox1.Update();
                                package.Save();
                                goto finish_;
                            }
                        }
                        catch (Exception)
                        {
                            goto nonvalue;
                        }

                        //---------------------------------------------------------
                        if (sheet.Cells[1, 2].Value.ToString() != null)
                        {
                            for (int row = 1; row < 9999; row++)
                            {
                                var thisValue = sheet.Cells[row, 2].Value;
                                if (thisValue != null)
                                {
                                    if (thisValue.ToString() == user)
                                    {
                                        sheet.DeleteRow(row, 2);
                                        //goto finder;
                                    }
                                }

                            }
                        }
                        if (sheet.Cells[1, 2].Value.ToString() != null)
                        {
                            for (int row = 1; row < 9999; row++)
                            {
                                var thisValue = sheet.Cells[row, 1].Value;
                                if (thisValue == null)
                                {
                                    sheet.Cells[row, 1].Value = Time_.ToString();
                                    sheet.Cells[row, 2].Value = user;
                                    package.Save();
                                    goto finder;
                                }
                                else
                                {
                                }
                            }
                        }
                    }

                }
                catch (Exception)
                {
                    //MessageBox.Show("Excel file not found", "");
                    //return;
                    goto beggining;
                }

                try
                {
                    SyncListUpdater_.listBox1.Items.Clear();
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                        for (int row = 1; row < 20; row++)
                        {
                            if (sheet.Cells[row, 1].Value == null)
                            {
                                break;
                            }
                            if (sheet.Cells[row, 1].Value != null)
                            {
                                var Value1 = sheet.Cells[row, 1].Value;
                                var Value2 = sheet.Cells[row, 2].Value;

                                if (lastSaveTime == DateTime.MinValue)
                                {
                                    lastSaveTime = DateTime.Now;
                                }
                                DateTime now = DateTime.Now;
                                TimeSpan elapsedTime = now.Subtract(lastSaveTime);
                                double minutes = elapsedTime.Minutes;
                                if (minutes > 2)
                                {
                                    SyncListUpdater_.Close();
                                    MessageBox.Show("10 minutes have passed, try to sync again", "Warning");
                                    return Autodesk.Revit.UI.Result.Cancelled;
                                }

                                SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString());

                                SyncListUpdater_.label1.Text = elapsedTime.ToString();

                                SyncListUpdater_.label1.Refresh();
                                SyncListUpdater_.label1.Update();

                            }

                        }

                        if (sheet.Cells[1, 1].Value != null && sheet.Cells[1, 2].Value.ToString() != user)
                        {
                            package.Save();
                            package.Dispose();
                            SyncListUpdater_.listBox1.Refresh();
                            SyncListUpdater_.listBox1.Update();
                            goto finder;
                        }
                        else
                        {
                        }
                    }
                }
                catch (Exception)
                {
                    //goto finder;
                }

            finder:

                if (SyncListUpdater_.DialogResult == DialogResult.Cancel)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                SyncListUpdater_.listBox1.Refresh();
                SyncListUpdater_.listBox1.Update();
                while (true)
                {

                    if (CheckTime == DateTime.MinValue)
                    {
                        CheckTime = DateTime.Now;
                    }
                    DateTime nowTocheck = DateTime.Now;

                    TimeSpan elapsedTimeToCheck = nowTocheck.Subtract(CheckTime);
                    double minutestoCheck = elapsedTimeToCheck.TotalSeconds;


                    SyncListUpdater_.label1.Text = elapsedTimeToCheck.ToString();

                    SyncListUpdater_.label1.Refresh();
                    SyncListUpdater_.label1.Update();

                    if (minutestoCheck > 10.0)
                    {
                        CheckTime = DateTime.Now;
                        try
                        {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                            {
                                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                                SyncListUpdater_.listBox1.Items.Clear();

                                for (int row = 1; row < 20; row++)
                                {
                                    if (sheet.Cells[row, 1].Value == null)
                                    {
                                        break;
                                    }
                                    if (sheet.Cells[row, 1].Value != null)
                                    {
                                        var Value1 = sheet.Cells[row, 1].Value;
                                        var Value2 = sheet.Cells[row, 2].Value;

                                        if (lastSaveTime == DateTime.MinValue)
                                        {
                                            lastSaveTime = DateTime.Now;
                                        }
                                        DateTime now = DateTime.Now;
                                        TimeSpan elapsedTime = now.Subtract(lastSaveTime);
                                        double minutes = elapsedTime.Minutes;
                                        if (minutes > 10)
                                        {
                                            SyncListUpdater_.Close();
                                            MessageBox.Show("10 minutes have passed", "Warning");
                                            return Autodesk.Revit.UI.Result.Cancelled;
                                        }

                                        SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString());
                                        //SyncListUpdater_.textBox1.Text = minutestoCheck.ToString() /*DateTime.Now.ToShortTimeString()*/;
                                        SyncListUpdater_.label1.Text = elapsedTime.ToString();
                                        //SyncListUpdater_.textBox1.Refresh();
                                        //SyncListUpdater_.textBox1.Update();
                                        SyncListUpdater_.label1.Refresh();
                                        SyncListUpdater_.label1.Update();

                                    }

                                }

                                if (sheet.Cells[1, 1].Value != null && sheet.Cells[1, 2].Value.ToString() != user)
                                {
                                    package.Save();
                                    package.Dispose();
                                    SyncListUpdater_.listBox1.Refresh();
                                    SyncListUpdater_.listBox1.Update();
                                    goto finder;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //SyncListUpdater_.textBox1.Text = time.ToString("hh':'mm':'ss") /*time.ToString()*/;
                            //SyncListUpdater_.listBox1.Refresh();
                            //SyncListUpdater_.listBox1.Update();
                            goto finder;
                        }
                    }
                }

            //for (int i = 0; i < reset; i++)
            //{


            //    if (i == 99999998)
            //    {

            //    }
            //}
            finish_:;
                SyncListUpdater_.Close();

                TransactWithCentralOptions transact = new TransactWithCentralOptions();
                SynchronizeWithCentralOptions synch = new SynchronizeWithCentralOptions();
                //synch.Comment = "Autosaved by the API at " + DateTime.Now;
                RelinquishOptions relinquishOptions = new RelinquishOptions(true);
                relinquishOptions.CheckedOutElements = true;
                synch.SetRelinquishOptions(relinquishOptions);

                //uiApp.Application.WriteJournalComment("AutoSave To Central", true);
                doc.SynchronizeWithCentral(transact, synch);

                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                        if (sheet.Cells[1, 2].Value != null)
                        {
                            if (sheet.Cells[1, 2].Value.ToString() == user)
                            {
                                sheet.DeleteRow(1, 1);
                                package.Save();
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show("Excel file not found", "");
                    //return;
                }
                return Autodesk.Revit.UI.Result.Succeeded;
            }
            else
            {
                TaskDialog.Show(doc.PathName.ToString(), "This tool is active only for =" + "RHR_BUILDING_A22");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

        }
    }


    class ribbonUI : IExternalApplication
    {
        public static FailureDefinitionId failureDefinitionId = new FailureDefinitionId(new Guid("E7BC1F65-781D-48E8-AF37-1136B62913F5"));
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = System.IO.Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2022\STH_Automation_22\img");
            string dll = Assembly.GetExecutingAssembly().Location;
            string myRibbon_1 = "Test Tools";

            application.CreateRibbonTab(myRibbon_1);
            RibbonPanel panel_1 = application.CreateRibbonPanel(myRibbon_1, "STH");
          
            PushButton Button1 = (PushButton)panel_1.AddItem(new PushButtonData("Align To Grid", "Align To Grid", dll, "STH_Automation_22.Rot_Ele_Angle_To_Grid"));
            Button1.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "RotateTogrid.png"), UriKind.Absolute));
            Button1.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";
            Button1.ToolTipImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGrid.jpg"), UriKind.Absolute));

            PushButton Button2 = (PushButton)panel_1.AddItem(new PushButtonData("Set Annotation Crop", "Set Annotation Crop", dll, "STH_Automation_22.Set_Annotation_Crop"));
            Button2.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "SetCropbox.png"), UriKind.Absolute));
           

            PushButton Button3_A = (PushButton)panel_1.AddItem(new PushButtonData("Elevate wall A", "Elevate wall A", dll, "STH_Automation_22.Wall_Elevation_A"));
            Button3_A.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "ElevateWall.png"), UriKind.Absolute));

            PushButton Button3_B = (PushButton)panel_1.AddItem(new PushButtonData("Elevate wall B", "Elevate wall B", dll, "STH_Automation_22.Wall_Elevation_B"));
            Button3_B.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "ElevateWall.png"), UriKind.Absolute));

            PushButton Button4 = (PushButton)panel_1.AddItem(new PushButtonData("Wall Angle To EleMarker", "Wall Angle To EleMarker", dll, "STH_Automation_22.Wall_Angle_To_EleMarker"));
            Button4.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "ElevateWall.png"), UriKind.Absolute));


            PushButton Button5 = (PushButton)panel_1.AddItem(new PushButtonData("Find Perpendicular Wall", "Find Perpendicular Wall", dll, "STH_Automation_22.Find_Perpendicular_Wall"));
            Button5.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
           

            PushButton Button6 = (PushButton)panel_1.AddItem(new PushButtonData("Sync Manager", "Sync Manager", dll, "STH_Automation_22.Sync_Manager"));
            Button6.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "sync.png"), UriKind.Absolute));

            PushButton Button7 = (PushButton)panel_1.AddItem(new PushButtonData("View BBox", "View BBox", dll, "STH_Automation_22.Export_View_Details"));
            Button7.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "sync.png"), UriKind.Absolute));

            PushButton Button8 = (PushButton)panel_1.AddItem(new PushButtonData("Import View", "Import View", dll, "STH_Automation_22.Import_View_Details"));
            Button8.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "sync.png"), UriKind.Absolute));

            return Autodesk.Revit.UI.Result.Succeeded;
        }
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}

