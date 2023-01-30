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
            Document doc = uidoc.Document;

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
    public class Wall_Elevation : IExternalCommand
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
    public class Wall_Angle_To_EleMarker : IExternalCommand
    {

        public XYZ GetElevationMarkerCenter(Document doc, Autodesk.Revit.DB.View symbolView, ElevationMarker marker, ViewSection elevation)
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
        private bool IsViewerInMarker(Document doc, ElevationMarker marker, Element viewer)
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


           


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Wall Section View");
                if (marker != null)
                {
                    BoundingBoxXYZ vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
                    XYZ center = (vbb.Max + vbb.Min) / 2;

                  

                    //Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(location, XYZ.BasisY);
                    //XYZ vect1 = line2.Direction * (1100 / 304.8);
                    XYZ vect2 = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
                    Makeline(doc, vect2, center);

                    Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(vect2, center);
                    //XYZ vect1 = line2.Direction * (1100 / 304.8);

                    if (lc != null)
                    {
                        Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(vect2, XYZ.BasisZ);
                        angle3 = line2.Direction.AngleTo(LineWallDir);
                        double AngleOfRotation = angle3 * 180 / Math.PI;

                        if (line2.Direction.IsAlmostEqualTo(LineWallDir))
                        {
                            goto end;
                        }
                        else
                        {
                            marker.Location.Rotate(axis, (angle3 ));

                            vect2 = GetElevationMarkerCenter(doc, doc.ActiveView, marker, views.ToArray()[0] as ViewSection);
                            vbb = ViewSection_.get_BoundingBox(doc.ActiveView);
                            center = (vbb.Max + vbb.Min) / 2;
                            line2 = Autodesk.Revit.DB.Line.CreateBound(vect2, center);
                            angle3 = line2.Direction.AngleTo(LineWallDir);
                            AngleOfRotation = angle3 * 180 / Math.PI;

                            if (line2.Direction.IsAlmostEqualTo(LineWallDir))
                            {
                                goto end;
                            }

                            XYZ Inverted = InvCoord(LineWallDir);
                            if (line2.Direction.IsAlmostEqualTo(Inverted))
                            {
                                goto end;
                            }
                            else
                            {
                                marker.Location.Rotate(axis, angle3 * -1);
                                //double RadiansOfRotationTest2 = ElementDir.AngleTo(direction);
                                //double AngleOfRotationTest2 = RadiansOfRotationTest2 * 180 / Math.PI;

                                marker.Location.Rotate(axis, angle3 * -1    );
                                //double RadiansOfRotationTest3 = ElementDir.AngleTo(direction);
                                //double AngleOfRotationTest3 = RadiansOfRotationTest3 * 180 / Math.PI;
                                goto end;
                            }
                        }
                    }
                    else
                    {
                        Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(vect2, XYZ.BasisZ);

                        angle3 = line2.Direction.AngleTo(DIR);
                        if (line2.Direction.IsAlmostEqualTo(DIR))
                        {
                            goto end;
                        }
                        else
                        {
                            marker.Location.Rotate(axis, (angle3 ));


                            if (line2.Direction.IsAlmostEqualTo(DIR))
                            {
                                goto end;
                            }

                            XYZ Inverted = InvCoord(DIR);
                            if (line2.Direction.IsAlmostEqualTo(Inverted))
                            {
                                goto end;
                            }
                            else
                            {
                                marker.Location.Rotate(axis, angle3 * -1);
                                //double RadiansOfRotationTest2 = ElementDir.AngleTo(direction);
                                //double AngleOfRotationTest2 = RadiansOfRotationTest2 * 180 / Math.PI;

                                marker.Location.Rotate(axis, angle3 * -1);
                                //double RadiansOfRotationTest3 = ElementDir.AngleTo(direction);
                                //double AngleOfRotationTest3 = RadiansOfRotationTest3 * 180 / Math.PI;
                                goto end;
                            }
                        }
                    }
                }

                if (lc != null)
                {


                }
            end:
                tx.Commit();
            }
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


    class ribbonUI : IExternalApplication
    {
        public static FailureDefinitionId failureDefinitionId = new FailureDefinitionId(new Guid("E7BC1F65-781D-48E8-AF37-1136B62913F5"));
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = System.IO.Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2022\STH_Automation_22\img");
            string dll = Assembly.GetExecutingAssembly().Location;
            string myRibbon_1 = "Modelling Tools";

            application.CreateRibbonTab(myRibbon_1);
            RibbonPanel panel_1 = application.CreateRibbonPanel(myRibbon_1, "STH");
          
            PushButton Button1 = (PushButton)panel_1.AddItem(new PushButtonData("Align To Grid", "Align To Grid", dll, "STH_Automation_22.Rot_Ele_Angle_To_Grid"));
            Button1.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "RotateTogrid.png"), UriKind.Absolute));
            Button1.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";
            Button1.ToolTipImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGrid.jpg"), UriKind.Absolute));

            PushButton Button2 = (PushButton)panel_1.AddItem(new PushButtonData("Set Annotation Crop", "Set Annotation Crop", dll, "STH_Automation_22.Set_Annotation_Crop"));
            Button2.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            //Button2.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";

            PushButton Button3 = (PushButton)panel_1.AddItem(new PushButtonData("Elevate wall", "Elevate wall", dll, "STH_Automation_22.Wall_Elevation"));
            Button3.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            //Button3.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";

            PushButton Button4 = (PushButton)panel_1.AddItem(new PushButtonData("Wall Angle To EleMarker", "Wall Angle To EleMarker", dll, "STH_Automation_22.Wall_Angle_To_EleMarker"));
            Button4.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "RotateEle.png"), UriKind.Absolute));
            //Button4.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";

            PushButton Button5 = (PushButton)panel_1.AddItem(new PushButtonData("Find Perpendicular Wall", "Find Perpendicular Wall", dll, "STH_Automation_22.Find_Perpendicular_Wall"));
            Button5.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            //Button4.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";

            return Autodesk.Revit.UI.Result.Succeeded;
        }
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}

