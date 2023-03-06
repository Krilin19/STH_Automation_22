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
using System.Security.Cryptography;
using Google.Cloud.Firestore;


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

            using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
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

            using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
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

                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
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

                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
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
           

            using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
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

            UIApplication uiapp = commandData.Application;
            DocumentSet documents = uiapp.Application.Documents;

            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;
            

            //string Sync_Manager = @"T:\Lopez\1.xlsx";
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            List<ViewSection> ViewSection_ = new List<ViewSection>();
            List<ViewSheet> ViewSheet_ = new List<ViewSheet>();
            List<Viewport> viewports = new List<Viewport>();
            //List<ElementId> views_ = new List<ElementId>();
            List<ElementId> text_type = new List<ElementId>();
            List<XYZ> text_pt = new List<XYZ>();
            List<TextNote> textnotessearch = new List<TextNote>();
            List<XYZ> vp_centers = new List<XYZ>();

            foreach (Autodesk.Revit.DB.Document doc_ in documents)
            {

                if (doc_.Title != doc.Title)
                {
                    using (Autodesk.Revit.DB.Transaction tx2 = new Autodesk.Revit.DB.Transaction(doc_))
                    {
                        tx2.Start("STH sheet copy");
                        foreach (ElementId selectedid in selectedIds)
                        {


                            Autodesk.Revit.DB.ViewSheet e = doc.GetElement(selectedid) as Autodesk.Revit.DB.ViewSheet;
                            string type = e.GetType().Name;

                            IList<ElementId> rev_Id = e.GetAllRevisionIds();
                            //ICollection<ElementId> viewports = e.GetAllViewports();
                            ICollection<ElementId> views = e.GetAllPlacedViews();

                            IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
                            .Where(q => q.SheetId == e.Id).ToList();

                            foreach (var item in viewports__)
                            {
                                viewports.Add(item);
                            }

                            foreach (var VP in viewports)
                            {

                                XYZ center = VP.GetBoxCenter();
                                vp_centers.Add(center);
                            }

                            if (type == "ViewSheet")
                            {
                                BoundingBoxXYZ box = e.get_BoundingBox(null /*doc.ActiveView*/);
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


                                //XYZ end = new XYZ(coordTR.X, coordTR.Y, coordBL.Z);
                                //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(transform.OfPoint(box.Min), transform.OfPoint(box.Max));
                                //XYZ midCero = new XYZ(line.Evaluate(0.5,true).X, line.Evaluate(0.5, true).Y, coordBL.Z);
                                //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(transform.BasisZ, end);
                                //double dis;
                                //UV uv_;
                                //plane.Project(coordBL, out uv_, out dis);
                                //XYZ SecDir = transform.BasisX; //Rightdir
                                //XYZ up = transform.BasisY;
                                //XYZ viewdir = SecDir.CrossProduct(up);
                                //XYZ vect1 = transform.BasisZ * (dis  /*/ 304.8*/);
                                //XYZ vect2 = vect1 + coordBL;
                                //Autodesk.Revit.DB.Line lineB = Autodesk.Revit.DB.Line.CreateBound(coordBL, vect2);
                                //Autodesk.Revit.DB.Line lineC = Autodesk.Revit.DB.Line.CreateBound(vect2, end);
                                //ModelCurve bot_left = Makeline(doc, coordBL, vect2);
                                //XYZ endvect1 = transform.BasisZ.Negate() * (dis  /*/ 304.8*/);
                                //XYZ endvect2 = endvect1 + end;
                                //Autodesk.Revit.DB.Line lineD = Autodesk.Revit.DB.Line.CreateBound(end,endvect2);
                                //ModelCurve bot_left2 = Makeline(doc, end, endvect2);
                                //ModelCurve bot_left3 = Makeline(doc, coordBL, endvect2);
                                //Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                //t.Origin = endvect2;
                                //t.BasisX = lineC.Direction.Negate();
                                //t.BasisY = up;
                                //t.BasisZ = lineB.Direction.Negate();
                                //BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                                //sectionBox.Transform = t;
                                //sectionBox.Min = min;
                                //sectionBox.Max = max;


                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = coordBR /*box.Transform.Origin*/;
                                t.BasisX = InvCoord(transform.BasisX);
                                t.BasisY = transform.BasisY;
                                t.BasisZ = InvCoord(transform.BasisZ);
                                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                                sectionBox.Transform = t /*transform*/;
                                sectionBox.Min = min;
                                sectionBox.Max = max;
                                ModelCurve New_cen_max = Makeline(doc_, transform.OfPoint(sectionBox.Min), transform.OfPoint(sectionBox.Max));
                                ViewFamilyType vft = new FilteredElementCollector(doc_).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
                                ViewSection view_onproject =  ViewSection.CreateSection(doc_, vft.Id, sectionBox /*box*/);

                                ViewSection_.Add(view_onproject);


                                //double TransX = transform.Origin.X;
                                //double TransY = transform.Origin.Y;
                                //double TransZ = transform.Origin.Z;
                                //using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                                //{
                                //    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                                //    sheet.Cells[1, 1].Value = TransX;
                                //    sheet.Cells[1, 2].Value = TransY;
                                //    sheet.Cells[1, 3].Value = TransZ;
                                //    sheet.Cells[2, 1].Value = Math.Round(t.BasisX.X, 0);
                                //    sheet.Cells[2, 2].Value = Math.Round(t.BasisX.Y, 0);
                                //    sheet.Cells[2, 3].Value = Math.Round(t.BasisX.Z, 0);
                                //    sheet.Cells[2, 4].Value = Math.Round(t.BasisY.X, 0);
                                //    sheet.Cells[2, 5].Value = Math.Round(t.BasisY.Y, 0);
                                //    sheet.Cells[2, 6].Value = Math.Round(t.BasisY.Z, 0);
                                //    sheet.Cells[2, 7].Value = Math.Round(t.BasisZ.X, 0);
                                //    sheet.Cells[2, 8].Value = Math.Round(t.BasisZ.Y, 0);
                                //    sheet.Cells[2, 9].Value = Math.Round(t.BasisZ.Z, 0);
                                //    sheet.Cells[3, 1].Value = min.X;
                                //    sheet.Cells[3, 2].Value = min.Y;
                                //    sheet.Cells[3, 3].Value = min.Z;
                                //    sheet.Cells[3, 4].Value = max.X;
                                //    sheet.Cells[3, 5].Value = max.Y;
                                //    sheet.Cells[3, 6].Value = max.Z;
                                //    package.Save();
                                //}


                                ViewSheet sheet2 = ViewSheet.Create(doc_, familyList.First().Id);

                                ViewSheet_.Add(sheet2);

                                if (sheet2.LookupParameter("Sheet Number") != null)
                                {
                                    string parametro = e.LookupParameter("Sheet Number").AsString();
                                    Parameter param = sheet2.LookupParameter("Sheet Number");
                                    param.Set(parametro);
                                }

                                if (sheet2.LookupParameter("Sheet Name") != null)
                                {
                                    string parametro = e.LookupParameter("Sheet Name").AsString();
                                    Parameter param2 = sheet2.LookupParameter("Sheet Name");
                                    param2.Set(parametro);
                                }
                            }
                        }
                        tx2.Commit();
                        using (Autodesk.Revit.DB.Transaction tx3 = new Autodesk.Revit.DB.Transaction(doc_))
                        {
                            tx3.Start("STH sheet copy");
                            foreach (var sheet2 in ViewSheet_)
                            {
                                Viewport viewport = Viewport.Create(doc_, sheet2.Id, ViewSection_.ToArray()[0].Id, new XYZ() /*vp_centers.ToArray()[centerpt]*/);
                                if (Viewport.CanAddViewToSheet(doc, sheet2.Id, ViewSection_.ToArray()[0].Id))
                                {
                                    
                                }
                            }
                            tx3.Commit();
                        }
                    }
                    {
                        //TaskDialog.Show("Basic Element Info", s);
                    }

                }
                
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

                using (Autodesk.Revit.DB.Transaction tx2 = new Autodesk.Revit.DB.Transaction(doc))
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
    public class Create_Sun_Eye_view : IExternalCommand
    {
        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            XYZ testPoint2 = edge.GetEndPoint(0);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                       

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;
                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;
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

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public double MapValue(double start_n, double end_n, double mapped_n_menusone, double mapped_n_one, double number_tobe_map)
        {
            return mapped_n_menusone + (mapped_n_one - mapped_n_menusone) * ((number_tobe_map - start_n) / (end_n - start_n));
        }
        public ViewOrientation3D GetCurrentViewOrientation(UIDocument doc)
        {
            XYZ UpDir = doc.ActiveView.UpDirection;
            XYZ ViewDir = doc.ActiveView.ViewDirection;
            XYZ ViewInvDir = InvCoord(ViewDir);
            XYZ eye = new XYZ(0, 0, 0);
            XYZ up = UpDir;
            XYZ forward = ViewInvDir;
            ViewOrientation3D MyNewOrientation = new ViewOrientation3D(eye, up, forward);
            return MyNewOrientation;
        }

        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        public XYZ CrossProduct(XYZ v1, XYZ v2)
        {
            double x, y, z;
            x = v1.Y * v2.Z - v2.Y * v1.Z;
            y = (v1.X * v2.Z - v2.X * v1.Z) * -1;
            z = v1.X * v2.Y - v2.X * v1.Y;
            var rtnvector = new XYZ(x, y, z);
            return rtnvector;
        }
        public XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
        {
            double degToRadian = Math.PI * 2 / 360;
            double angleHorizR = angleHorizD * degToRadian;
            double angleVertR = angleVertD * degToRadian;
            double a = Math.Cos(angleVertR);
            double b = Math.Cos(angleHorizR);
            double c = Math.Sin(angleHorizR);
            double d = Math.Sin(angleVertR);
            return new XYZ(a * b, a * c, d);
        }

        public class Vector3D
        {
            public Vector3D(XYZ revitXyz)
            {
                XYZ = revitXyz;
            }
            public Vector3D() : this(XYZ.Zero)
            { }
            public Vector3D(double x, double y, double z)
              : this(new XYZ(x, y, z))
            { }
            public XYZ XYZ { get; private set; }
            public double X => XYZ.X;
            public double Y => XYZ.Y;
            public double Z => XYZ.Z;
            public Vector3D CrossProduct(Vector3D source)
            {
                return new Vector3D(XYZ.CrossProduct(source.XYZ));
            }
            public double GetLength()
            {
                return XYZ.GetLength();
            }
            public override string ToString()
            {
                return XYZ.ToString();
            }
            public static Vector3D BasisX => new Vector3D(
              XYZ.BasisX);
            public static Vector3D BasisY => new Vector3D(
              XYZ.BasisY);
            public static Vector3D BasisZ => new Vector3D(
              XYZ.BasisZ);
        }
        static AddInId appId = new AddInId(new Guid("8D3F5703-A09A-6ED6-864C-5720329D9677"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 2;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

         

            ProjectLocation plCurrent = doc.ActiveProjectLocation;


            Autodesk.Revit.DB.View currentView = uidoc.ActiveView;
            SunAndShadowSettings sunSettings = currentView.SunAndShadowSettings;


            // Set the initial direction of the sun at ground level (like sunrise level)
            XYZ initialDirection = XYZ.BasisY;

            // Get the altitude of the sun from the sun settings
            double altitude = sunSettings.GetFrameAltitude(
              sunSettings.ActiveFrame);

            // Create a transform along the X axis based on the altitude of the sun
            Autodesk.Revit.DB.Transform altitudeRotation = Autodesk.Revit.DB.Transform
              .CreateRotation(XYZ.BasisX, altitude);

            // Create a rotation vector for the direction of the altitude of the sun
            XYZ altitudeDirection = altitudeRotation
              .OfVector(initialDirection);

            // Get the azimuth from the sun settings of the scene
            double azimuth = sunSettings.GetFrameAzimuth(
              sunSettings.ActiveFrame);

            // Correct the value of the actual azimuth with true north

            // Get the true north angle of the project
            Element projectInfoElement
              = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .FirstElement();

            BuiltInParameter bipAtn
              = BuiltInParameter.BASEPOINT_ANGLETON_PARAM;

            Parameter patn = projectInfoElement.get_Parameter(
              bipAtn);

            double trueNorthAngle = patn.AsDouble();

            // Add the true north angle to the azimuth
            double actualAzimuth = 2 * Math.PI - azimuth + trueNorthAngle;

            // Create a rotation vector around the Z axis
            Autodesk.Revit.DB.Transform azimuthRotation = Autodesk.Revit.DB.Transform
              .CreateRotation(XYZ.BasisZ, actualAzimuth);

            // Finally, calculate the direction of the sun
            XYZ sunDirection = azimuthRotation.OfVector(
              altitudeDirection);



            //}
            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;


            XYZ UpDir = uidoc.ActiveView.UpDirection;
           

            ViewOrientation3D viewOrientation3D;
            using (Autodesk.Revit.DB.Transaction tr1 = new Autodesk.Revit.DB.Transaction(doc))
            {
            
                tr1.Start("Place vs in sheet");
                View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);
                tr1.SetName("Create view " + view3D.Name);
   

                XYZ eye = XYZ.Zero;
                XYZ inverted_sun_location = InvCoord(sunDirection);

                XYZ origin_b = new XYZ(0, 0, 0);
                XYZ normal_B = new XYZ(1, 0, 0);


                Autodesk.Revit.DB.Transform trans3;
                Autodesk.Revit.DB.Transform trans_inverted_direction;


                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), sunDirection);
                XYZ vect1 = line2.Direction * (100000 / 304.8);
                XYZ vect2 = vect1 + new XYZ(0, 0, 0);
                ModelCurve mc = Makeline(doc, line2.Origin, vect2);

                Autodesk.Revit.DB.Plane Plane_mirror = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(/*mc.SketchPlane.GetPlane().Normal*/ new XYZ(0, 0, 1), sunDirection);
                Autodesk.Revit.DB.Plane forward_dir = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(mc.SketchPlane.GetPlane().Normal, sunDirection);

                trans3 = Autodesk.Revit.DB.Transform.CreateReflection(Plane_mirror);
                XYZ inv_sun_mirrored = trans3.OfVector(inverted_sun_location);

                trans_inverted_direction = Autodesk.Revit.DB.Transform.CreateReflection(forward_dir);
                XYZ inverted_direction = trans_inverted_direction.OfVector(inverted_sun_location);

                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), inv_sun_mirrored);
                XYZ invertedvect1 = line3.Direction * (100000 / 304.8);
                XYZ invertedvect2 = invertedvect1 + new XYZ(0, 0, 0);
                Makeline(doc, line2.Origin, invertedvect2);


                XYZ cross_product = CrossProduct(/*inv_sun,*/ /*new_p*/ inv_sun_mirrored, inverted_sun_location);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), cross_product);
                XYZ invertedvect3 = line4.Direction * (100000 / 304.8);
                XYZ invertedvect4 = invertedvect3 + new XYZ(0, 0, 0);
                Makeline(doc, line2.Origin, invertedvect4);


                XYZ cross_product_up = CrossProduct(/*inv_sun,*/ /*new_p*/ line2.Direction, cross_product);



                XYZ startPoint = sunDirection;
                XYZ endPoint = inverted_sun_location;
                Autodesk.Revit.DB.Line geomLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
                XYZ pntCenter = geomLine.Evaluate(0.0, true);
                Autodesk.Revit.DB.Line geomLine2 = Autodesk.Revit.DB.Line.CreateBound(doc.ActiveView.Origin, XYZ.BasisZ);
                Autodesk.Revit.DB.Plane geomPlane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(sunDirection, sunDirection);

                SketchPlane sketch = SketchPlane.Create(doc, geomPlane);
                sketch.Name = view3D.Name;
                doc.ActiveView.SketchPlane = sketch;
                doc.ActiveView.ShowActiveWorkPlane();
                view3D.SketchPlane = sketch;
                view3D.ShowActiveWorkPlane();


                Autodesk.Revit.DB.Transform rot2 = Autodesk.Revit.DB.Transform.CreateRotation(/*orientLine.Direction*/mc.GeometryCurve.GetEndPoint(1), -2.60);
                XYZ rotated_vec2 = rot2.OfVector(inverted_direction);

                viewOrientation3D = new ViewOrientation3D(/*eye*/ mc.GeometryCurve.GetEndPoint(0), cross_product_up/*rotated_vec2*/ * -1, /*inverted_direction*/rotated_vec2);
                view3D.SetOrientation(viewOrientation3D);
                view3D.SaveOrientationAndLock();
                tr1.Commit();
            }


            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Sync_Manager : IExternalCommand
    {
        //static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        static public Autodesk.Revit.ApplicationServices.Application m_app;
        public float abort = 0;


        public void Add_Document_with_AustoID()
        {
            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath2 = System.IO.Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2022\STH_Automation_22\");
            string path = /*ppDomain.CurrentDomain.BaseDirectory*/folderPath2 + @"revit-api-test-firebase.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);
            FirestoreDb db = FirestoreDb.Create("revit - api - test");
           

            CollectionReference cool = db.Collection("Add_Document_with_AustoID");
            Dictionary<string, object> data1 = new Dictionary<string, object>()
            {
                { "First name","tacv"},
                { "Lastname","Alexitico"},
                { "age ","8"}
            };
            cool.AddAsync(data1);
            MessageBox.Show("data added successfully");
        }


        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {


            Add_Document_with_AustoID();

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

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rhino_access : IExternalCommand
    {
        public static List<List<Element>> GetMembersRecursive(Autodesk.Revit.DB.Document d, Group g, List<List<Element>> r, List<List<string>> strin_name, List<int> int_)
        {
            if (strin_name == null)
            {
                strin_name = new List<List<string>>();
            }
            if (r == null)
            {
                r = new List<List<Element>>();
            }
            if (int_ == null)
            {
                int_ = new List<int>();
            }
            List<string> lista_nombre_main = new List<string>();
            List<string> lista_nombre_buildings = new List<string>();
            List<string> lista_nombre_floor = new List<string>();
            List<Group> lista_de_buildings = new List<Group>();
            List<Group> lista_de_floor = new List<Group>();
            List<List<Element>> ele_list = new List<List<Element>>();
            List<Element> ceiling_groups1 = new List<Element>();

            List<Element> elems = g.GetMemberIds().Select(q => d.GetElement(q)).ToList();
            lista_nombre_main.Add(g.Name);
            lista_nombre_buildings.Add(g.Name);

            foreach (Element el in elems)
            {
                if (el.GetType() == typeof(Group))
                {
                    Group gp = el as Group;
                    lista_de_buildings.Add(gp);
                    lista_nombre_buildings.Add(el.Name);
                }
                if (el.GetType() == typeof(Ceiling))
                {
                    ceiling_groups1.Add(el);
                }
            }
            r.Add(ceiling_groups1);
            for (int i = 0; i < lista_de_buildings.ToArray().Length; i++)
            {
                Group gp2 = lista_de_buildings.ToArray()[i];

                List<Element> elems2 = gp2.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                ele_list.Add(elems2);
            }

            for (int i = 0; i < ele_list.ToArray().Length; i++)
            {
                List<Element> lista1 = ele_list.ToArray()[i];
                List<Element> ceiling_groups = new List<Element>();
                foreach (var item in lista1)
                {
                    if (item.GetType() == typeof(Group))
                    {
                        Group gp4 = item as Group;
                        List<Element> elems3 = gp4.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                        foreach (var item2 in elems3)
                        {
                            if (item2.GetType() == typeof(Group))
                            {
                                Group gp5 = item2 as Group;
                                List<Element> elems4 = gp5.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                                foreach (var item3 in elems4)
                                {
                                    if (item3.GetType() == typeof(Ceiling))
                                    {
                                        ceiling_groups.Add(item3);
                                        lista_nombre_floor.Add(item.Name);
                                    }
                                }
                            }
                            if (item2.GetType() == typeof(Ceiling))
                            {
                                ceiling_groups.Add(item2);
                                lista_nombre_floor.Add(item.Name);
                            }
                        }

                        //lista_de_floor.Add(gp4);
                        //lista_nombre_floor.Add(item.Name);
                    }
                    if (item.GetType() == typeof(Ceiling))
                    {
                        ceiling_groups.Add(item);
                        lista_nombre_floor.Add(item.Name);

                    }
                }
                r.Add(ceiling_groups);
            }
            strin_name.Add(lista_nombre_main);
            strin_name.Add(lista_nombre_buildings);
            strin_name.Add(lista_nombre_floor);

            return r;

        }

        public static List<List<Face>> GetFaces(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            if (normal.Z == 0 && normal.Y > -0.8 /*&& normal.X < 0*/)
                            {
                                Element e = doc_.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                faces_list.Add(face);
                            }
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }

        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            double lenght = Math.Round(edge.ApproximateLength, 0);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;

                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;

        }

        static AddInId appId = new AddInId(new Guid("D044091D-29A4-4F70-8FE5-84FBD4ED0D73"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

          

            //Form1 form3 = new Form1();
            //form3.ShowDialog();




            //if (form3.radioButton1.Checked)
            //{
            //    Form22 form4 = new Form22();
            //    form4.ShowDialog();

            //    Form20 form2 = new Form20();

            //    form2.ShowDialog();
            //}



            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Form1 form = new Form1();

            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            //Group grp_Lot = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
            ///this code was use to select ceilings and explore is facing direction so they can be reproduce in rhino geometry///
            try
            {
                ICollection<Reference> my_faces = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");
                foreach (var item_myRefWall in my_faces)
                {
                    Element e = doc.GetElement(item_myRefWall);
                    GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                    Face face = geoobj as Face;
                    PlanarFace planarFace = face as PlanarFace;
                    XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                    name_of_roof.Add("roof");
                    faces_picked.Add(face);


                    if (item_myRefWall == my_faces.ToArray().Last())
                    {
                        Faces_lists_excel.Add(faces_picked);

                        //names.Add(name_of_roof);
                    }

                }
            }
            catch (Exception)
            {

                MessageBox.Show("no surfaces were selected", "Warning");
            }





            Group grpExisting = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
            GetMembersRecursive(doc, grpExisting, elemente_selected, names, numeros_);

            foreach (var item in elemente_selected)
            {
                foreach (var item2 in item)
                {
                    if (!form.listBox1.Items.Contains(item2.Name))
                    {
                        form.listBox1.Items.Add(item2.Name);
                    }
                }
            }

            //Form23 form5 = new Form23();
            //form5.ShowDialog();

            //Form24 form6 = new Form24();
            //form6.ShowDialog();


            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            List<List<Element>> elemente_selected_to_bedeleted = new List<List<Element>>();
            foreach (var item in elemente_selected)
            {
                List<Element> ele_sel = new List<Element>();
                foreach (var item2 in item)
                {
                    foreach (var item3 in form.listBox2.Items)
                    {
                        if (item3.ToString() == item2.Name)
                        {
                            ele_sel.Add(item2);
                        }
                    }
                }
                elemente_selected_to_bedeleted.Add(ele_sel);
            }

            string name_of_group = names.ToArray()[0].ToArray()[0].ToString();
            names.ToArray()[1].Insert(1, "roof" + name_of_group);
            GetFaces(doc, elemente_selected_to_bedeleted, Faces_lists_excel);
            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);
            TaskDialog.Show("info faces", info);
            //string filename = Path.Combine(Path.GetTempPath(), "Book1.xlsx"); /// this line was used to automatically look for this excel file///

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            int numero = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {
                package.Workbook.Worksheets.Delete(1);
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("my_data");
                int row = 1;
                for (int i = 0; i < xyz_faces.ToArray().Length; i++)
                {
                    numero = 0;
                    foreach (var item in xyz_faces.ToArray()[i])
                    {
                        if (item == null)
                        {
                            numero += 1;
                            sheet.Cells[row, 1].Value = names.ToArray()[0][0];
                            sheet.Cells[row, 2].Value = names.ToArray()[1][i + 1];
                            sheet.Cells[row, 3].Value = numero;
                            row++;
                        }
                        else
                        {
                            sheet.Cells[row, 1].Value = Math.Round(item.X, 1);
                            sheet.Cells[row, 2].Value = Math.Round(item.Y, 1);
                            sheet.Cells[row, 3].Value = Math.Round(item.Z, 1);
                            row++;
                        }

                        if (item == xyz_faces.ToArray()[i].Last())
                        {
                            sheet.Cells[row, 1].Value = "Next";
                            sheet.Cells[row, 2].Value = ".";
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }
                    }
                }
                package.Save();
            }
            Process.Start(filename2);
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rhino_access_faces : IExternalCommand
    {

        public static List<List<Face>> GetFaces_individual(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            Element e = doc_.GetElement(item3.Reference);
                            GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                            Face face = geoobj as Face;
                            faces_list.Add(face);
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }
        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            XYZ testPoint2 = edge.GetEndPoint(0);
                            double lenght = Math.Round(edge.ApproximateLength, 0);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                            XYZ dir = new XYZ(0, 0, 0) - testPoint1;

                            //ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face,doc_.ActiveView as View3D);
                            //ReferenceWithContext referenceWithContext = refIntersector.FindNearest(testPoint1 , dir);

                            //Reference reference = referenceWithContext.GetReference();
                            //XYZ intersection = reference.GlobalPoint;
                            //XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);
                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;
                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;
        }

        static AddInId appId = new AddInId(new Guid("D031092D-29A4-4F70-8FE5-84FBD4ED0D73"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 4;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "Createsheet" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<Face> element = new List<Face>();
            List<Reference> my_faces = new List<Reference>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Form1 form = new Form1();

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.checkBox1.Checked == false)
            {
                foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Ceiling)))
                {
                    if (item.Name.Contains(form.textBox1.Text))
                    {
                        Options op = new Options();
                        op.ComputeReferences = true;
                        foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item2.Faces)
                            {
                                PlanarFace planarFace = item3 as PlanarFace;
                                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                                Element e = doc.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                element.Add(face);
                            }
                        }

                        //element.Clear();
                    }
                }
                Faces_lists_excel.Add(element);
            }
            else
            {
                //Group grp_Lot = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
                ///this code was use to select ceilings and explore is facing direction so they can be reproduce in rhino geometry///
                ICollection<Reference> my_faces_ = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

                foreach (var item in my_faces_)
                {
                    my_faces.Add(item);
                }
            }

            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            foreach (var item_myRefWall in my_faces)
            {
                Element e = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                PlanarFace planarFace = face as PlanarFace;
                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                name_of_roof.Add("roof");
                faces_picked.Add(face);


                if (item_myRefWall == my_faces.ToArray().Last())
                {
                    Faces_lists_excel.Add(faces_picked);

                    //names.Add(name_of_roof);
                }

            }

            //names.ToArray()[1].Insert(1, "individual faces" );
            //GetFaces_individual(doc, elemente_selected, Faces_lists_excel);
            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);
            TaskDialog.Show("Excel writting", "Writting faces information in a temporary excel file");

            //string filename = Path.Combine(Path.GetTempPath(), "Book1.xlsx"); /// this line was used to automatically look for this excel file///

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            int numero = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {
                package.Workbook.Worksheets.Delete(1);
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("my_data");
                int row = 1;
                for (int i = 0; i < xyz_faces.ToArray().Length; i++)
                {
                    numero = 0;

                    foreach (var item in xyz_faces.ToArray()[i])
                    {

                        if (item == null)
                        {
                            numero += 1;

                            sheet.Cells[row, 1].Value = form.textBox1.Text;
                            sheet.Cells[row, 2].Value = form.textBox1.Text;
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }
                        else
                        {
                            sheet.Cells[row, 1].Value = Math.Round(item.X, 1);
                            sheet.Cells[row, 2].Value = Math.Round(item.Y, 1);
                            sheet.Cells[row, 3].Value = Math.Round(item.Z, 1);
                            row++;
                        }

                        if (item == xyz_faces.ToArray()[i].Last())
                        {
                            sheet.Cells[row, 1].Value = "Next";
                            sheet.Cells[row, 2].Value = ".";
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }

                    }

                }

                package.Save();
            }

            Process.Start(filename2);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    public class TextTypeUpdater : IUpdater
    {
        static AddInId m_appId;
        static UpdaterId m_updaterId;
        FailureDefinitionId _failureId = null;

        // constructor takes the AddInId for the add-in associated with this updater
        public TextTypeUpdater(AddInId id)
        {
            m_appId = id;
            // every Updater must have a unique ID
            m_updaterId = new UpdaterId(m_appId, new Guid("FBFBF6B2-4C06-42d4-97C1-D1B4EB593EFF"));

            
            _failureId = new FailureDefinitionId(new Guid("A1882A5F-CE11-4302-A416-F22310F1730F"));

            FailureDefinition failureDefinition
             = FailureDefinition.CreateFailureDefinition(_failureId, FailureSeverity.Error,
               "PreventDeletion: sorry, this element cannot be deleted.");


        }
        public void Execute(UpdaterData data)
        {
            Autodesk.Revit.DB.Document doc = data.GetDocument();


            //foreach (ElementId addedElemId in data.GetAddedElementIds())
            //{
            //    //TextNoteType textNoteType = doc.GetElement(addedElemId) as TextNoteType;
            //    Autodesk.Revit.DB.View textNoteType = doc.GetElement(addedElemId) as Autodesk.Revit.DB.View;
            //    string name = textNoteType.Name;
            //    doc.Delete(addedElemId);
            //    //TaskDialog.Show("New Element Deleted!", "Text type '" + name + "' has been deleted. Please use existing text types in the template.");
            //    //var paramnames =  doc.GetElement(addedElemId).Parameters;

            //    //string name_ = doc.GetElement(addedElemId).LookupParameter("View Type").AsString();

            //    TaskDialog.Show("!!!", "Do not forget to fill View type parameter for this view");
            //}
            
           
            foreach (ElementId addedElemId in data.GetDeletedElementIds())
            {
                FailureMessage failureMessage = new FailureMessage(_failureId);

                failureMessage.SetFailingElement(addedElemId);
                doc.PostFailure(failureMessage);
                //TextNoteType textNoteType = doc.GetElement(addedElemId) as TextNoteType;
                var textNoteType = doc.GetElement(addedElemId) /*as Autodesk.Revit.DB.View*/;
                string name = textNoteType.Name;

                //TaskDialog.Show("New Element Deleted!", "Text type '" + name + "' has been deleted. Please use existing text types in the template.");
                //var paramnames =  doc.GetElement(addedElemId).Parameters;

                //string name_ = doc.GetElement(addedElemId).LookupParameter("View Type").AsString();

                TaskDialog.Show("!!!", "Do not forget to fill View type parameter for this view");

            }
        }
        public string GetAdditionalInformation() { return "Text note type check"; }
        public ChangePriority GetChangePriority() { return ChangePriority.FloorsRoofsStructuralWalls; }
        public UpdaterId GetUpdaterId() { return m_updaterId; }
        public string GetUpdaterName() { return "Text note type"; }
    }

    

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Register : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            TextTypeUpdater updater = new TextTypeUpdater(uidoc.Application.ActiveAddInId);
            UpdaterRegistry.RegisterUpdater(updater);




            // Trigger will occur only for TextNoteType elements
            //ElementClassFilter textNoteTypeFilter = new ElementClassFilter(typeof(TextNoteType));
            ElementClassFilter textNoteTypeFilter = new ElementClassFilter(typeof(Autodesk.Revit.DB.View));

            // GetChangeTypeElementAddition specifies that the triggger will occur when elements are added
            // Other options are GetChangeTypeAny, GetChangeTypeElementDeletion, GetChangeTypeGeometry, GetChangeTypeParameter
            //UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), textNoteTypeFilter, Element.GetChangeTypeElementAddition());
            UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), textNoteTypeFilter, Element.GetChangeTypeElementDeletion());

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class UnRegister : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            TextTypeUpdater updater = new TextTypeUpdater(uidoc.Application.ActiveAddInId);
            UpdaterRegistry.UnregisterUpdater(updater.GetUpdaterId());

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Data_base_Rego : IExternalCommand
    {
        public FirestoreDb datab_()
        {
            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath2 = System.IO.Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2022\STH_Automation_22\");
            string path = /*ppDomain.CurrentDomain.BaseDirectory*/folderPath2 + @"revit-api-test-firebase.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);
            FirestoreDb db = FirestoreDb.Create("revit-api-test");
            return db;
        }


        public void Add_Document_with_AustoID(FirestoreDb db)
        {

            CollectionReference cool = db.Collection(/*"Add_Document_with_AustoID"*/"Sync Manager");
            Dictionary<string, object> data1 = new Dictionary<string, object>()
            {
                { "STH Project","Ryde Hospital"},
                { "User Sync","Grant Larson"},
                { "Time","6/03/2023 11:19:37 AM"}
            };
            cool.AddAsync(data1);
            MessageBox.Show("data added successfully");
        }

        public void Add_Document_with_CustomsID(FirestoreDb db, Autodesk.Revit.DB.Document doc)
        {

            DocumentReference docdb = db.Collection("Sync Manager 2").Document(doc.Title);
            string user = doc.Application.Username;
            var lastSaveTime = DateTime.Now;
            var CheckTime = DateTime.Now;

            Dictionary<string, object> data1 = new Dictionary<string, object>()
            {
                

                { "Waiting state", "false"},
                { "User Sync", user.ToString()},
                { "Time",DateTime.Now.ToString()}
            };
            docdb.SetAsync(data1);
            //MessageBox.Show("data added successfully");
        }

        async void getAllData(FirestoreDb db)
        {
            DocumentReference docRef = db.Collection("Sync Manager").Document("Ryde Hospital");
            DocumentSnapshot snap = await docRef.GetSnapshotAsync();
            Dictionary<string, object> data2 = snap.ToDictionary();
        }


        async void All_Documunets_FromACollection(FirestoreDb db)
        {
            Query docRef = db.Collection("Sync Manager");
            QuerySnapshot snap = await docRef.GetSnapshotAsync();
            string s = "You Picked:" + "\n";
            foreach (DocumentSnapshot project in snap)
            {
                Dictionary<string, object> data2 = project.ToDictionary();
                s += " Doc Id = " + project.Id + "\n";
                s += " coll = " + data2.Keys + " = " + data2.Values +  "\n";

            }
            TaskDialog.Show("Basic Element Info", s);
        }


        static AddInId appId = new AddInId(new Guid("7F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            
            FirestoreDb db = datab_();

            //Add_Document_with_AustoID();

            //getAllData(db);

            //All_Documunets_FromACollection(db);

            Add_Document_with_CustomsID(db, doc);


            return Autodesk.Revit.UI.Result.Succeeded;
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

            PushButton Button9 = (PushButton)panel_1.AddItem(new PushButtonData("Suneye View", "Suneye View", dll, "STH_Automation_22.Create_Sun_Eye_view"));
            Button9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "sun_eyes_view_2.png"), UriKind.Absolute));
            Button9.ToolTip = "Create a isometric Revit view from sun position towards project origin";
            Button9.LongDescription = "...";

            PushButton Button10 = (PushButton)panel_1.AddItem(new PushButtonData("Family Geo Exporter", "Family Geo Exporter", dll, "STH_Automation_22.Rhino_access"));
            Button10.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButton Button11 = (PushButton)panel_1.AddItem(new PushButtonData("Family Geo Search", "Family Geo Search", dll, "STH_Automation_22.Rhino_access_faces"));
            Button11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButton Button12 = (PushButton)panel_1.AddItem(new PushButtonData("Register", "Register", dll, "STH_Automation_22.Register"));
            Button12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButton Button13 = (PushButton)panel_1.AddItem(new PushButtonData("UnRegister", "UnRegister", dll, "STH_Automation_22.UnRegister"));
            Button13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButton Button14 = (PushButton)panel_1.AddItem(new PushButtonData("Data base Rego", "Data base Rego", dll, "STH_Automation_22.Data_base_Rego"));
            Button14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));


            return Autodesk.Revit.UI.Result.Succeeded;
        }
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}

