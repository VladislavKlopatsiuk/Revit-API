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
using System.Diagnostics;

namespace Спека_труб_горизонтальных_и_вертикальных
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BranchingWhithoutIntersects : IExternalCommand
    {
        public Result  Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            //View activeView = doc.ActiveView;

            List<Pipe> pipe = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            var pipe_types = new Dictionary<string, string>()
            {
                ["Труба полипропиленовая армированная стекловолокном PN20"] ="Крепление для полипропиленового трубопровода",
                ["Труба полипропиленовая PN20"] ="Крепление для полипропиленового трубопровода",
                ["Труба полипропиленовая PN20"] ="Крепление для полипропиленового трубопровода",
                ["Труба полиэтиленовая напорная ПЭ100 SDR11"] ="Крепление для полиэтиленового трубопровода",
                ["Труба полиэтиленовая напорная ПЭ100 SDR17"] ="Крепление для полиэтиленового трубопровода",
                ["Труба полиэтиленовая напорная ПЭ100 SDR17"] ="Крепление для полиэтиленового трубопровода",
                ["Труба из сшитого полиэтилена PN20"] ="Крепление для труб из сшитого полиэтилена",
                ["Труба из сшитого полиэтилена PN20"] ="Крепление для труб из сшитого полиэтилена",
                ["Труба стальная водогазопроводная оцинкованная"] = "Крепление для стального трубопровода",
                ["Труба стальная электросварная прямошовная"] = "Крепление для стального трубопровода",
                ["Трубы канализационные услиенные, PP-MD SN10"] = "Крепление для полипропиленового трубопровода",
                ["Трубы канализационные услиенные, PP-MD SN10"] = "Крепление для полипропиленового трубопровода",
                ["Труба НПВХ с раструбом под клеевое соединение SDR21 PN10"] = "Крепление для ПВХ трубопровода",
                ["Труба НПВХ с раструбом под клеевое соединение SDR21 PN10"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR17 1,6 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR21 1,25 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR26 1,0 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR33 0,8 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR41 0,63 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба напорная раструбная НПВХ SDR41 0,63 МПа"] = "Крепление для ПВХ трубопровода",
                ["Труба канализационная ПВХ SN4"] = "Крепление для ПВХ трубопровода",
                ["Труба канализационная ПВХ SN8"] = "Крепление для ПВХ трубопровода",
                ["Труба канализационная полипропиленовая"] = "Крепление для полипропиленового трубопровода",
                ["Труба чугунная SML"] = "Крепление для полиэтиленового трубопровода",
                ["Труба чугунная канализационная"] = "Крепление для чугунного трубопровода",
            };

            var pipe_diameter_option = new Dictionary<string, int>()
            {
                ["Труба полипропиленовая армированная стекловолокном PN20"] = 1,
                ["Труба полипропиленовая PN20"] = 1,
                ["Труба полипропиленовая PN20"] = 1,
                ["Труба полиэтиленовая напорная ПЭ100 SDR11"] = 1,
                ["Труба полиэтиленовая напорная ПЭ100 SDR17"] = 1,
                ["Труба полиэтиленовая напорная ПЭ100 SDR17"] = 1,
                ["Труба из сшитого полиэтилена PN20"] = 1,
                ["Труба из сшитого полиэтилена PN20"] = 1,
                ["Труба стальная водогазопроводная оцинкованная"] = 4,
                ["Труба стальная электросварная прямошовная"] = 2,
                ["Трубы канализационные услиенные, PP-MD SN10"] = 3,
                ["Трубы канализационные услиенные, PP-MD SN10"] = 3,
                ["Труба НПВХ с раструбом под клеевое соединение SDR21 PN10"] = 1,
                ["Труба НПВХ с раструбом под клеевое соединение SDR21 PN10"] = 1,
                ["Труба напорная раструбная НПВХ SDR17 1,6 МПа"] = 1,
                ["Труба напорная раструбная НПВХ SDR21 1,25 МПа"] = 1,
                ["Труба напорная раструбная НПВХ SDR26 1,0 МПа"] = 1,
                ["Труба напорная раструбная НПВХ SDR33 0,8 МПа"] = 1,
                ["Труба напорная раструбная НПВХ SDR41 0,63 МПа"] = 1,
                ["Труба напорная раструбная НПВХ SDR41 0,63 МПа"] = 1,
                ["Труба канализационная ПВХ SN4"] = 3,
                ["Труба канализационная ПВХ SN8"] = 3,
                ["Труба канализационная полипропиленовая"] = 3,
                ["Труба чугунная SML"] = 2,
                ["Труба чугунная канализационная"] = 2,
            };

            var KTR_diameter_option_steel = new Dictionary<double, int>()
            {
                [15] = 15,                
                [20] = 20,
                [25] = 25,
                [32] = 32,
                [40] = 40,
                [50] = 50,
                [65] = 76,
                [80] = 89,
                [100] = 108,
                [125] = 127,
                [150] = 159,               
            };
            
            var KTR_range_steel = new Dictionary<double, double>()
            {
                [15] = 1.5,                
                [20] = 2.0,
                [25] = 2.0,
                [32] = 2.5,
                [40] = 3.0,
                [50] = 3.0,
                [65] = 4.0,
                [80] = 4.0,
                [100] = 4.5,
                [125] = 5.0,
                [150] = 6.0,               
            };

            var KTR_diameter_option_polimer = new Dictionary<double, int>()
            {
                [16] = 15,
                [20] = 15,
                [25] = 20,
                [32] = 25,
                [40] = 32,
                [50] = 40,
                [63] = 50,
                [75] = 76,
                [90] = 89,
                [110] = 108,
                [125] = 127,
                [160] = 159,
            };
            
            var KTR_range_polimer = new Dictionary<double, double>()
            {
                [16] = 0.5,
                [20] = 0.5,
                [25] = 0.6,
                [32] = 0.6,
                [40] = 0.75,
                [50] = 0.9,
                [63] = 1.0,
                [75] = 1.1,
                [90] = 1.2,
                [110] = 1.2,
                [125] = 1.2,
                [160] = 1.2,
            };

            var PU_diameter_option_steel = new Dictionary<double, int>()
            {
                [15] = 8,
                [20] = 8,
                [25] = 8,
                [32] = 8,
                [40] = 8,
                [50] = 10,
                [65] = 10,
                [80] = 10,
                [100] = 12,
                [125] = 12,
                [150] = 12,               
            };

            var PU_diameter_option_polimer = new Dictionary<double, int>()
            {
                [16] = 8,
                [20] = 8,
                [25] = 8,
                [32] = 8,
                [40] = 8,
                [50] = 10,
                [63] = 10,
                [75] = 10,
                [90] = 10,
                [110] = 12,
                [125] = 12,
                [160] = 12,
            };

            using (Transaction transaction = new Transaction(doc))

            {
                Stopwatch sw = Stopwatch.StartNew();
                sw.Start();
                transaction.Start("Запись о положении труб в пространстве");
               
                //Обработка труб
                foreach (var curveElement in pipe)
                {
                    Curve pipeLine = (curveElement.Location as LocationCurve).Curve;

                    XYZ hightPoint = pipeLine.GetEndPoint(0),
                        lowPoint = pipeLine.GetEndPoint(1);

                    int xHight = ((int)hightPoint.X);
                    int yHight = ((int)hightPoint.Y);
                    int xLow = ((int)lowPoint.X);
                    int yLow = ((int)lowPoint.Y);

                    bool vertical = (xHight == xLow) && (yHight == yLow);

                    if (vertical)
                    {
                        curveElement.get_Parameter(new Guid("478da18d-2221-4c17-8b50-e2b5be5800d0")).Set("Вертикальный участок");
                    }
                    else
                    {
                        curveElement.get_Parameter(new Guid("478da18d-2221-4c17-8b50-e2b5be5800d0")).Set("Горизонтальный участок");
                    }

                    double level = curveElement.get_Parameter(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION).AsDouble();

                    double level_offset = UnitUtils.ConvertFromInternalUnits(level, UnitTypeId.Millimeters);

                    bool branching_to_wall = level_offset < 1250 && !vertical;

                    //Удаление уже заполненных значений при повторном прожатии

                    Parameter bracing_name = curveElement.get_Parameter(new Guid("87662a0c-8ba8-4a5e-886f-525a1201632f"));

                    Parameter bracing_quantity = curveElement.get_Parameter(new Guid("b4f903ef-cf0c-4bc1-88b5-2bd4789d51e4"));

                    Parameter bracing_comments = curveElement.get_Parameter(new Guid("b5ab5728-547c-4995-abfb-8ee7107a7a81"));

                    if(bracing_name != null|| bracing_quantity != null || bracing_comments != null)
                    {
                        bracing_name.Set("");

                        bracing_quantity.Set(0.0);

                        bracing_comments.Set("");
                    }

                    //Проверка на необходимость устройства крепления по пользовательскому параметру "галочка"

                    bool necessity_user_parameter = curveElement.get_Parameter(new Guid("90010061-1e4e-4e71-8cfe-4cbf75189b12"))?.AsInteger() == 1;

                    if (necessity_user_parameter)
                    {

                    //Получение типа трубы

                    PipeType pipe_type = curveElement.PipeType;

                    //Получение параметра "коментарии к типоразмеру" и поиск значения в словаре

                    Parameter type_comment = pipe_type.LookupParameter("Комментарии к типоразмеру");

                    string pipe_type_size = type_comment.AsString();

                    //Получение условного диаметра трубы и перевод из имперской в метрическую

                    Parameter parameter_diametr_general = curveElement.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

                    double diametr_general_imperial = parameter_diametr_general.AsDouble();

                    double diametr_general_metric = UnitUtils.ConvertFromInternalUnits(diametr_general_imperial, UnitTypeId.Millimeters);

                    //Получение внешнего диаметра трубы и перевод из имперской в метрическую

                    Parameter parameter_diametr_outside = curveElement.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);

                    double diametr_outside_imperial = parameter_diametr_outside.AsDouble();

                    double diametr_outside_metric = UnitUtils.ConvertFromInternalUnits(diametr_outside_imperial, UnitTypeId.Millimeters);                    

                    //Получение условного диаметра трубы и перевод из имперской в метрическую

                    Parameter parameter_diametr_inside = curveElement.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);

                    double diametr_inside_imperial = parameter_diametr_inside.AsDouble();

                    double diametr_inside_metric = UnitUtils.ConvertFromInternalUnits(diametr_inside_imperial, UnitTypeId.Millimeters);

                    //Получение параметра ADSK_Количество

                    double adsk_quantity = curveElement.get_Parameter(new Guid("8d057bb3-6ccd-4655-9165-55526691fe3a")).AsDouble();                                   

                    //расчёт толщины стенки трубопровода

                    double pipe_wall_thickness = (diametr_outside_metric - diametr_inside_metric) / 2;

                    //Проверка на наличие ключа в словаре

                    bool contain_value = pipe_types.ContainsKey(pipe_type_size);

                    bool contain_value_pipe_option = pipe_diameter_option.ContainsKey(pipe_type_size);

                    //Определение объёма пересечки и сравнение его с общим объёмом трубы для определения рациональности простановки крепления    
                                        
                    //Запись значений параметров 

                    if (contain_value && contain_value_pipe_option)
                    {
                        int pipe_option = pipe_diameter_option[pipe_type_size];                        

                        string bracing_type = pipe_types[pipe_type_size];
                        if (!vertical && !branching_to_wall)
                        {
                            switch (pipe_option)
                        {
                            case 1:

                                bool contain_value_KTR_polimer = KTR_diameter_option_polimer.ContainsKey(diametr_general_metric);

                                if (contain_value_KTR_polimer)
                                {
                                    int KTR_DN_polimer = KTR_diameter_option_polimer[diametr_general_metric];
                                    int PU_DN_polimer = PU_diameter_option_polimer[diametr_general_metric];

                                        bracing_name.Set($"Крепление полимерного трубопровода ⌀{diametr_general_metric}x{pipe_wall_thickness} в комплекте из:" +
                                        $"\n -КТР-{KTR_DN_polimer}" +
                                        $"\n -Подвеска ПР-{PU_DN_polimer} L=300мм" +
                                        $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_polimer}" +
                                        $"\n -Планка ПУ-{PU_DN_polimer}");

                                        double KTR_range_for_diameter = KTR_range_polimer[diametr_general_metric];

                                        double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                        bracing_quantity.Set(KTR_quantity);

                                        bracing_comments.Set("К потолку");
                                }
                                break;                                

                            case 2:

                                bool contain_value_KTR_steel = KTR_diameter_option_steel.ContainsKey(diametr_general_metric);

                                if (contain_value_KTR_steel)
                                {
                                    int KTR_DN_steel = KTR_diameter_option_steel[diametr_general_metric];
                                    int PU_DN_steel = PU_diameter_option_steel[diametr_general_metric];

                                        bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                        $"\n -КТР-{KTR_DN_steel}" +
                                        $"\n -Подвеска ПР-{PU_DN_steel} L=300мм" +
                                        $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_steel}" +
                                        $"\n -Планка ПУ-{PU_DN_steel}");

                                        double KTR_range_for_diameter = KTR_range_steel[diametr_general_metric];

                                        double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                        bracing_quantity.Set(KTR_quantity);

                                        bracing_comments.Set("К потолку");
                                    }
                                break;
                            
                            case 3:

                                bool contain_value_KTR_polimer_sewage = KTR_diameter_option_polimer.ContainsKey(diametr_general_metric);

                                if (contain_value_KTR_polimer_sewage)
                                {
                                    int KTR_DN_polimer = KTR_diameter_option_polimer[diametr_general_metric];
                                    int PU_DN_polimer = PU_diameter_option_polimer[diametr_general_metric];

                                        bracing_name.Set($"Крепление полимерного трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                        $"\n -КТР-{KTR_DN_polimer}" +
                                        $"\n -Подвеска ПР-{PU_DN_polimer} L=300мм" +
                                        $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_polimer}" +
                                        $"\n -Планка ПУ-{PU_DN_polimer}");

                                        double KTR_range_for_diameter = KTR_range_polimer[diametr_general_metric];

                                        double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                        bracing_quantity.Set(KTR_quantity);

                                        bracing_comments.Set("К потолку");
                                    }
                                break;

                            case 4:

                                bool contain_value_KTR_VGP = KTR_diameter_option_steel.ContainsKey(diametr_general_metric);

                                if (contain_value_KTR_VGP)
                                {
                                    int KTR_DN_steel = KTR_diameter_option_steel[diametr_general_metric];
                                    int PU_DN_steel = PU_diameter_option_steel[diametr_general_metric];

                                    bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                        $"\n -КТР-{KTR_DN_steel}" +
                                        $"\n -Подвеска ПР-{PU_DN_steel} L=300мм" +
                                        $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_steel}" +
                                        $"\n -Планка ПУ-{PU_DN_steel}");

                                        double KTR_range_for_diameter = KTR_range_steel[diametr_general_metric];

                                        double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                        bracing_quantity.Set(KTR_quantity);

                                        bracing_comments.Set("К потолку");
                                    }

                                else
                                {
                                    bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                                                            $"\n -КТР-{15}" +
                                                                            $"\n -Подвеска ПР-{8} L=300мм" +
                                                                            $"\n -Дюбель-втулка распорная латунная ДВ-М{8}" +
                                                                            $"\n -Планка ПУ-{8}");

                                        double KTR_quantity = adsk_quantity / 1.5;

                                        bracing_quantity.Set(KTR_quantity);

                                        bracing_comments.Set("К потолку");
                                    }

                                break;
                        };
                        }

                        else
                        {
                            switch (pipe_option)
                            {
                                case 1:

                                    bool contain_value_KTR_polimer = KTR_diameter_option_polimer.ContainsKey(diametr_general_metric);

                                    if (contain_value_KTR_polimer)
                                    {
                                        int KTR_DN_polimer = KTR_diameter_option_polimer[diametr_general_metric];
                                        int PU_DN_polimer = PU_diameter_option_polimer[diametr_general_metric];

                                        bracing_name.Set($"Крепление полимерного трубопровода ⌀{diametr_general_metric}x{pipe_wall_thickness} в комплекте из:" +
                                            $"\n -КТР-{KTR_DN_polimer}" +
                                            $"\n -Подвеска ПР-{PU_DN_polimer} L=300мм" +
                                            $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_polimer}");

                                        if(branching_to_wall)
                                            {
                                                double KTR_range_for_diameter = KTR_range_polimer[diametr_general_metric];

                                                double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                        else
                                            {
                                                double KTR_quantity = adsk_quantity;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }

                                    }
                                    break;

                                case 2:

                                    bool contain_value_KTR_steel = KTR_diameter_option_steel.ContainsKey(diametr_general_metric);

                                    if (contain_value_KTR_steel)
                                    {
                                        int KTR_DN_steel = KTR_diameter_option_steel[diametr_general_metric];
                                        int PU_DN_steel = PU_diameter_option_steel[diametr_general_metric];

                                        bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                            $"\n -КТР-{KTR_DN_steel}" +
                                            $"\n -Подвеска ПР-{PU_DN_steel} L=300мм" +
                                            $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_steel}");

                                            if (branching_to_wall)
                                            {
                                                double KTR_range_for_diameter = KTR_range_steel[diametr_general_metric];

                                                double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                            else
                                            {
                                                double KTR_quantity = adsk_quantity / 3; ;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                            
                                    }
                                    break;

                                case 3:

                                    bool contain_value_KTR_polimer_sewage = KTR_diameter_option_polimer.ContainsKey(diametr_general_metric);

                                    if (contain_value_KTR_polimer_sewage)
                                    {
                                        int KTR_DN_polimer = KTR_diameter_option_polimer[diametr_general_metric];
                                        int PU_DN_polimer = PU_diameter_option_polimer[diametr_general_metric];

                                        bracing_name.Set($"Крепление полимерного трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                            $"\n -КТР-{KTR_DN_polimer}" +
                                            $"\n -Подвеска ПР-{PU_DN_polimer} L=300мм" +
                                            $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_polimer}");

                                            if (branching_to_wall)
                                            {
                                                double KTR_range_for_diameter = KTR_range_polimer[diametr_general_metric];

                                                double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                            else
                                            {
                                                double KTR_quantity = adsk_quantity;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                        }
                                    break;

                                case 4:

                                    bool contain_value_KTR_VGP = KTR_diameter_option_steel.ContainsKey(diametr_general_metric);

                                    if (contain_value_KTR_VGP)
                                    {
                                        int KTR_DN_steel = KTR_diameter_option_steel[diametr_general_metric];
                                        int PU_DN_steel = PU_diameter_option_steel[diametr_general_metric];

                                        bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                            $"\n -КТР-{KTR_DN_steel}" +
                                            $"\n -Подвеска ПР-{PU_DN_steel} L=300мм" +
                                            $"\n -Дюбель-втулка распорная латунная ДВ-М{PU_DN_steel}");

                                            if (branching_to_wall)
                                            {
                                                double KTR_range_for_diameter = KTR_range_steel[diametr_general_metric];

                                                double KTR_quantity = adsk_quantity / KTR_range_for_diameter;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                            else
                                            {
                                                double KTR_quantity = adsk_quantity / 3; ;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                        }

                                    else
                                    {
                                        bracing_name.Set($"Крепление стального трубопровода ⌀{diametr_general_metric} в комплекте из:" +
                                                                                $"\n -КТР-{15}" +
                                                                                $"\n -Подвеска ПР-{8} L=300мм" +
                                                                                $"\n -Дюбель-втулка распорная латунная ДВ-М{8}");

                                            if (branching_to_wall)
                                            {                                              

                                                double KTR_quantity = adsk_quantity / 1.5;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                            else
                                            {
                                                double KTR_quantity = adsk_quantity / 3; ;

                                                bracing_quantity.Set(KTR_quantity);

                                                bracing_comments.Set("К стене");
                                            }
                                        }

                                    break;
                            };
                        }
                    }

                }

                }
                transaction.Commit();
                sw.Stop();
                TimeSpan generaltime = sw.Elapsed;
                TaskDialog.Show("Запись о положении труб в пространстве", generaltime.TotalSeconds.ToString());
                TaskDialog.Show("Время работы", "Отработано");
            }
            

            
            return Result.Succeeded;
        }     

    }
}

