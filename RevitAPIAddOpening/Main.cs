using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIAddOpening
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document arDoc = uidoc.Document;
            Document ovDoc = arDoc.Application.Documents
                .OfType<Document>()
                .Where(x => x.Title.Contains("OV"))
                .FirstOrDefault();

            Document vkDoc = arDoc.Application.Documents
                .OfType<Document>()
                .Where(x => x.Title.Contains("VK"))
                .FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show ("Error", "Not find file OV");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(vkDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("Opening"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Error", "Not find Opening.rfa");
                return Result.Cancelled;
            }

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Error", "Not find 3D view");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction ts1 = new Transaction(arDoc, "activate opening");
            ts1.Start();
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            ts1.Commit();

            Transaction ts = new Transaction(arDoc, "Add opening");
            ts.Start();
            foreach (Duct duct in ducts)
            {
                Line line = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;

                List<ReferenceWithContext> referenceWithContexts = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct (new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext reference in referenceWithContexts)
                {
                    double proximity = reference.Proximity;
                    Reference refer = reference.GetReference();
                    Wall wall = arDoc.GetElement(refer.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointOpening = point + (direction * proximity);

                    FamilyInstance opening = arDoc.Create.NewFamilyInstance(pointOpening, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = opening.LookupParameter("Width");
                    Parameter heigth = opening.LookupParameter("Heigth");

                    width.Set(duct.Diameter);
                    heigth.Set(duct.Diameter);
                }
            }

            foreach (Pipe pipe in pipes)
            {
                Line linePipe = (pipe.Location as LocationCurve).Curve as Line;
                XYZ pointPipe = linePipe.GetEndPoint(0);
                XYZ directionPipe = linePipe.Direction;

                List<ReferenceWithContext> intersectionsP = referenceIntersector.Find(pointPipe, directionPipe)
                    .Where(x => x.Proximity <= linePipe.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext reference in intersectionsP)
                {
                    double proximity = reference.Proximity;
                    Reference refer = reference.GetReference();
                    Wall wall = arDoc.GetElement(refer.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointOpening = pointPipe + (directionPipe * proximity);

                    FamilyInstance opening = arDoc.Create.NewFamilyInstance(pointOpening, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = opening.LookupParameter("Ширина");
                    Parameter heigth = opening.LookupParameter("Высота");

                    width.Set(pipe.Diameter);
                    heigth.Set(pipe.Diameter);
                }
            }
            ts.Commit();
            return Result.Succeeded;
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
