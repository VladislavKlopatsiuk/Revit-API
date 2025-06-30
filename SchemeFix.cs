using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Изоляция;
using System.Windows;
using System.Data.Common;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace Спека_труб_горизонтальных_и_вертикальных
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SchemeFix : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;

            Document doc = uiDoc.Document;

            List<Instance> list_elements = new FilteredElementCollector(doc)
                .OfClass(typeof(Instance))
                .Cast<Instance>()
                .ToList();

            Schema schema = Schema.Lookup(new Guid("99445e32-4a81-441e-9879-4d4cfdef7aaa"));

            using (Transaction transaction = new Transaction(doc))

            {
                transaction.Start("Исправление схем");

                //Обработка схем
                foreach (var element in list_elements)
                {                  

                    if (schema != null && element != null)
                    {

                        doc.EraseSchemaAndAllEntities(schema);

                    }

                }

                transaction.Commit();
            }

            TaskDialog.Show("Исправление схем", "Отработано");
            return Result.Succeeded;
        }
    }
}