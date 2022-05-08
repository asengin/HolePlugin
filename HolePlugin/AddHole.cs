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

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc
                .Application.Documents
                .OfType<Document>()
                .Where(x => x.Title.Contains("ОВ"))
                .FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семество \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D Вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element,
                view3D);

            using (Transaction tr0 = new Transaction(arDoc, "Активация семейства"))
            {
                tr0.Start();

                if (!familySymbol.IsActive)
                    familySymbol.Activate();

                tr0.Commit();
            }

            using (Transaction tr = new Transaction(arDoc, "Установка отверстий"))
            {
                tr.Start();
                
                foreach (Duct duct in ducts)
                {
                    Line curve = (duct.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0);
                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= curve.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                        Parameter width = hole.LookupParameter("Ширина");
                        Parameter height = hole.LookupParameter("Высота");
                        width.Set(duct.Width + 0.16);
                        height.Set(duct.Height + 0.16);
                    }
                }

                foreach (Pipe pipe in pipes)
                {
                    Line curve = (pipe.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0);
                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= curve.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                        Parameter width = hole.LookupParameter("Ширина");
                        Parameter height = hole.LookupParameter("Высота");
                        width.Set(pipe.Diameter + 0.16);
                        height.Set(pipe.Diameter + 0.16);
                    }
                }

                tr.Commit();
            }

            return Result.Succeeded;
        }
    }
}
