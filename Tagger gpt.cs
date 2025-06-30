using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Plumbing;
using System.Windows.Controls;
using Grid = Autodesk.Revit.DB.Grid;
using System.Windows;
using System.Xml.Linq;
using Изоляция;
using System;
using System.Windows.Shapes;
using Line = Autodesk.Revit.DB.Line;

[Transaction(TransactionMode.Manual)]
public class CreateDimensions : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;
        View activeView = doc.ActiveView;

        // Проверяем, является ли активный вид планом этажа
        if (activeView.ViewType != ViewType.FloorPlan)
        {
            message = "Активный вид должен быть планом этажа.";
            return Result.Failed;
        }
        
        List<FamilyInstance> cleanouts = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .OfType<FamilyInstance>()
                    .Where(it => it.Name == "DN 110" || it.Name == "DN 110 HL" || it.Name == "HL98")
                    .ToList();

        List<Grid> grids_list = new FilteredElementCollector(doc, activeView.Id)
                .OfCategory(BuiltInCategory.OST_Grids)
                .OfType<Grid>()
                .ToList();     
        
        // Начинаем транзакцию
        using (Transaction trans = new Transaction(doc, "Create Dimensions"))
        {
            trans.Start();
            
            List<Grid> gridsX = new List<Grid>();
            List<Grid> gridsY = new List<Grid>();

            foreach (var grid in grids_list)
            {
                XYZ direction = (grid.Curve as Line).Direction;
                if (direction.IsAlmostEqualTo(new XYZ(1, 0, 0)) || direction.IsAlmostEqualTo(new XYZ(-1, 0, 0)))
                {
                    gridsX.Add(grid);
                }
                else if (direction.IsAlmostEqualTo(new XYZ(0, 1, 0)) || direction.IsAlmostEqualTo(new XYZ(0, -1, 0)))
                {
                    gridsY.Add(grid);
                }
            }

            XYZ FindNearestPointOnGrid(FamilyInstance element, Grid grid)
            {
                XYZ location = (element.Location as LocationPoint).Point;                
                Line gridLine = grid.Curve as Line; 

                // Находим ближайшую точку на оси
                XYZ nearestPoint = gridLine.Project(location).XYZPoint;
                return nearestPoint;
            }

            ReferenceArray dimXRefArray = new ReferenceArray();
            ReferenceArray dimYRefArray = new ReferenceArray();

            // Создаем размеры
            foreach (FamilyInstance cleanout in cleanouts)
            {
                XYZ location = (cleanout.Location as LocationPoint).Point;              

                Parameter level = cleanout.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);

                double form_level = level.AsDouble();

                XYZ location_Z = new XYZ(location.X, location.Y, form_level);

                dimXRefArray.Insert(cleanout.GetReferences(FamilyInstanceReferenceType.WeakReference)[0], dimXRefArray.Size);
                dimYRefArray.Insert(cleanout.GetReferences(FamilyInstanceReferenceType.WeakReference)[0], dimYRefArray.Size);

                List<ElementId> justListX = new List<ElementId>();
                List <ElementId> justListY = new List<ElementId>();

                // Найти ближайшие параллельные оси
                Grid nearestGridX = gridsX.OrderBy(grid => location.DistanceTo(FindNearestPointOnGrid(cleanout, grid))).First();

                Reference nearestGridXRef = new Reference(nearestGridX);
                dimXRefArray.Append(nearestGridXRef);
                ElementId idGridX = nearestGridX.Id;
                justListX.Add(idGridX);

                Grid nearestGridY = gridsY.OrderBy(grid => location.DistanceTo(FindNearestPointOnGrid(cleanout, grid))).First();
                Reference nearestGridYRef = new Reference(nearestGridY);
                dimYRefArray.Append(nearestGridYRef);
                ElementId idGridY = nearestGridY.Id;
                justListY.Add(idGridY);

                //var simpleform = new SimpleForm(justListY);  ////////////////////////////////////dsdsdsdsds
                //simpleform.ShowDialog();               

                //Создаем размеры
                if (nearestGridX != null && dimXRefArray.Size > 0)
                {
                    XYZ nearestPointX = FindNearestPointOnGrid(cleanout, nearestGridX);

                    // Получение линии оси
                    Line gridLineX = nearestGridX.Curve as Line;

                    // Получение точки пересечения линии оси и семейства
                    XYZ intersectionPointX = gridLineX.Project(nearestPointX).XYZPoint;

                    XYZ intersectionPointX_Z_zero = new XYZ(intersectionPointX.X, intersectionPointX.Y, form_level);

                    if (location_Z.DistanceTo(intersectionPointX_Z_zero) > 0.1)
                    {
                        // Создаем линию для размера
                        Line dimensionLineX = Line.CreateBound(location_Z + new XYZ(3, 0, 0), intersectionPointX_Z_zero + new XYZ(3, 0, 0));

                        // Создаем размер
                        Dimension dimX = doc.Create.NewDimension(activeView, dimensionLineX, dimXRefArray);                      

                    }                    

                    dimXRefArray.Clear();

                }

                if (nearestGridY != null && dimYRefArray.Size > 0)
                {
                    XYZ nearestPointY = FindNearestPointOnGrid(cleanout, nearestGridY);
                    
                    // Получение линии оси
                    Line gridLineY = nearestGridY.Curve as Line;

                    // Получение точки пересечения линии оси и семейства
                    XYZ intersectionPointY = gridLineY.Project(nearestPointY).XYZPoint;

                    XYZ intersectionPointY_Z_zero = new XYZ(intersectionPointY.X, intersectionPointY.Y, form_level);

                    

                    if (location_Z.DistanceTo(intersectionPointY_Z_zero) > 0.1)
                    {
                        // Создаем линию для размера
                        Line dimensionLineY = Line.CreateBound(location_Z + new XYZ(0, 3, 0), intersectionPointY_Z_zero + new XYZ(0, 3, 0));

                        // Создаем размер
                        Dimension dimY = doc.Create.NewDimension(activeView, dimensionLineY, dimYRefArray);

                    }

                    dimYRefArray.Clear();
                }
            }
            trans.Commit();
        }
        return Result.Succeeded;
    }
}
