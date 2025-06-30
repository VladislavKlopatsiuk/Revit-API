using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System;
using Изоляция;
using System.Xml.Linq;
using System.Windows;
using System.Collections;

namespace Спека_труб_горизонтальных_и_вертикальных
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Tagger : IExternalCommand
    {
        
        public Result  Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View activeView = doc.ActiveView;

            List<FamilyInstance> pipeFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .OfType<FamilyInstance>()
                    .Where(it => it.Name == "DN 110" || it.Name == "DN 110 HL")
                    .ToList();
            
            List<FamilyInstance> grids = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Grids)
                    .OfType<FamilyInstance>()
                    .ToList();

            FamilySymbol tagRefernce = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_PipeFittingTags)
                .OfType<FamilySymbol>()
                .Single(it => it.FamilyName == "ADSK_M_Соединители трубопроводов" && it.Name == "ADSK_Наименование краткое_25");

            IEnumerable<Element> modelElements = new FilteredElementCollector(doc, activeView.Id)
            .WhereElementIsNotElementType()
            .ToElements()
            .Where(e => e.Category != null && e.Category.CategoryType == CategoryType.Model);

            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Маркировка прочисток");

                List<ElementId> ids_intersect = new List<ElementId>();

                List<ElementId> tags_to_delite = new List<ElementId>();

                foreach (var element in pipeFittings)
                {
                    var tag_intersect_dictionary = new Dictionary<ElementId, int>()
                    {

                    };

                    int[,] tagOffsets ={{-20,-20,20,20},
                                        {20,-20,-20,20}};

                    //int minIntersections = 500;

                    //IndependentTag bestTag = null;

                    //List<ElementId> ids_Tags = new List<ElementId>();

                    //List<IndependentTag> tag_list = new List<IndependentTag>();

                    for (int i = 0; i < 4; i++)

                    {
                        Reference reference = new Reference(element);
                        LocationPoint locationPoint = (LocationPoint)element.Location;
                        if (locationPoint is null) continue;
                        XYZ point = locationPoint.Point;
                        XYZ tagPoint = new XYZ(point.X + tagOffsets[0,i] * activeView.Scale / 304.8,
                            point.Y - tagOffsets[1,i] * activeView.Scale / 304.8,
                            point.Z);
                        IndependentTag tag = IndependentTag.Create(
                            doc,
                            tagRefernce.Id,
                            activeView.Id,
                            reference,
                            true,
                            TagOrientation.Horizontal,
                            tagPoint
                            );

                        tag.LeaderEndCondition = LeaderEndCondition.Free;
                        tag.TagHeadPosition = tagPoint;
                        tag.SetLeaderEnd(
                            reference,
                            point);

                        Outline tagOutline = GetOutlineFromTag(tag);
                        BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(tagOutline);
                        //ElementIntersectsFilter elem_filter = new ElementIntersectsElementFilter(tag);

                        var intersectingElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                                .WherePasses(filter)
                                .ToElements();                        

                        foreach (var intersectingElement in intersectingElements)
                            {
                                ElementId id_intersect = intersectingElement.Id;
                                ids_intersect.Add(id_intersect);
                            }                        

                        int intersections = ids_intersect.Count;                        

                        tag_intersect_dictionary.Add( tag.Id, intersections);                        

                        //ids_Tags.Add(tag.Id);
                        //tag_list.Add(tag);
                    }                    

                    var minEntry = tag_intersect_dictionary.Aggregate((l, r) => l.Value < r.Value ? l : r);

                    ElementId minKey = minEntry.Key;

                    var tags_to_delite_instance = tag_intersect_dictionary.Keys.Where(key => key != minKey).ToList();

                    foreach (var tag_to_delite in tags_to_delite_instance)
                    {
                        if (tags_to_delite != null)
                        {
                            tags_to_delite.Add(tag_to_delite);
                        }
                    }                    
                }

                foreach (var tag_to_delite in tags_to_delite)
                {
                    if (tags_to_delite != null)
                    {
                        doc.Delete(tag_to_delite);
                    }
                }
                transaction.Commit();

                var simpleForm = new SimpleForm(tags_to_delite);
                simpleForm.ShowDialog();
            }           

            TaskDialog.Show("Труба", "Отработано");
            return Result.Succeeded;
        }
        //private int GetIntersectionCount(Document doc, IndependentTag tag)
        //{
        //    Outline tagOutline = GetOutlineFromTag(tag);
        //    BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(tagOutline);
            

        //    var intersectingElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
        //        .WherePasses(filter)
        //        .ToElements();

        //    foreach (var intersectingElement in intersectingElements)
        //    {
        //        ElementIntersectsElementFilter elem_filter = new ElementIntersectsElementFilter(intersectingElement);
        //    }

        //    return intersectingElements.Count;
        //}

        private Outline GetOutlineFromTag(IndependentTag tag)
        {
            // Получение габаритов марки
            BoundingBoxXYZ boundingBox = tag.get_BoundingBox(null);
            XYZ min = XYZ.Zero; // Значение по умолчанию
            XYZ max = XYZ.Zero; // Значение по умолчанию

            if (boundingBox != null)
            {
                if (boundingBox.Min != null)
                {
                    min = boundingBox.Min;
                }
                if (boundingBox.Max != null)
                {
                    max = boundingBox.Max;
                }
            }

            return new Outline(min, max);
        }
    }
}





