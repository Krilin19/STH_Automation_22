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
            Button1.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            Button1.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";
            Button1.ToolTipImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGrid.jpg"), UriKind.Absolute));

            PushButton Button2 = (PushButton)panel_1.AddItem(new PushButtonData("Set Annotation Crop", "Set Annotation Crop", dll, "STH_Automation_22.Set_Annotation_Crop"));
            Button2.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            Button2.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";
            Button2.ToolTipImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGrid.jpg"), UriKind.Absolute));

            PushButton Button3 = (PushButton)panel_1.AddItem(new PushButtonData("Elevate wall", "Elevate wall", dll, "STH_Automation_22.Wall_Elevation"));
            Button3.LargeImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGridIcon.png"), UriKind.Absolute));
            Button3.LongDescription = "This tool helps to make parallel any line base element or family to a selected grid angle";
            Button3.ToolTipImage = new BitmapImage(new Uri(System.IO.Path.Combine(folderPath, "AlignToGrid.jpg"), UriKind.Absolute));

            //try
            //{
            //    foreach (Autodesk.Windows.RibbonTab tab in Autodesk.Windows.ComponentManager.Ribbon.Tabs)
            //    {
            //        if (tab.Title == "Insert")
            //        {
            //            tab.IsVisible = false;
            //        }
            //    }
            //    adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;
            //    //ImageSource imgbg = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            //    //    "gradient.png"), UriKind.Relative));
            //    ImageBrush picBrush = new ImageBrush();
            //    //picBrush.ImageSource = imgbg;
            //    picBrush.AlignmentX = AlignmentX.Left;
            //    picBrush.AlignmentY = AlignmentY.Top;
            //    picBrush.Stretch = Stretch.None;
            //    picBrush.TileMode = TileMode.FlipXY;
            //    LinearGradientBrush gradientBrush = new LinearGradientBrush();
            //    gradientBrush.StartPoint = new System.Windows.Point(0, 0);
            //    gradientBrush.EndPoint = new System.Windows.Point(0, 1);
            //    gradientBrush.GradientStops.Add(new GradientStop(Colors.White, 0.0));
            //    gradientBrush.GradientStops.Add(new GradientStop(Colors.Orange, 1.0));
            //}
            //catch (Exception ex)
            //{
            //    winform.MessageBox.Show(
            //      ex.StackTrace + "\r\n" + ex.InnerException,
            //      "Error", winform.MessageBoxButtons.OK);
            //    return Result.Failed;
            //}
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}

