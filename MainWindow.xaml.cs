using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using SolidWorks.Interop.sldworks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Runtime.InteropServices;
using System.Windows.Media;
using WormGearGenerator.Helpers;
using SolidWorks.Interop.swconst;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Threading;
using System.Windows.Controls.Primitives;

namespace WormGearGenerator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //глобальные значения геометрических параметров
        float Module;
        float PressureAngle;
        float Peredat;
        float aw, Alpha, dae2, da1, da2, d1, d2, df1, df2, dw1, dw2, h1, ha1;
        public SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");

        //Статус загрузки главного окна
        public bool Status;
        //инициализация пути сохранения и базового пути
        public string InitialPath = null;
        string _pathFromProject;
        private string baseDirectory = System.Environment.CurrentDirectory;
        //пути сохранения компонентов
        public static string _pathAssembly { get; set; }
        public static string _pathWorm { get; set; }
        public static string _pathGear { get; set; }
        //имена компонентов
        public static string _nameAssembly { get; set; }
        public static string _nameWorm { get; set; }
        public static string _nameGear { get; set; }

        public object swFaceObject = null;
        public string selectedEntity;
        public string selectedComponent;
        public bool GearOrWorm;

        public static double[] pointsOfOriginComponent = new double[3];
        Face2[] faceArray = new Face2[2];

        //Токены отмены выполнения ансинхронных методов
        CancellationTokenSource _tokenSourceCylinder = new CancellationTokenSource();
        CancellationTokenSource _tokenSourceFace = new CancellationTokenSource();

        //подтверждение построения
        public static bool buildAccepted { get; set; }
        //инициализация объекта класса проверки полей программы
        private DataValidation validation;
        //инициализация объекта класса MaterialHelper
        private MaterialHelper MatObj;
        //инициализация объекта класса Worm
        private Worm worm;
        //инициализация объекта класса Gear
        private Gear gear;
        //инициализация объекта класса SolidHelper
        private SolidHelper sldobj;
        //инциализация объекта класса ConfirmFolders
        private ConfirmFolders foldersWindow;


        public MainWindow()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            ModelDoc2 swModel = default(ModelDoc2);
            AssemblyDoc swNewPart = default(AssemblyDoc);

            sldobj = new SolidHelper();
            swNewPart = (AssemblyDoc)swApp.ActiveDoc;
            swModel = (ModelDoc2)swNewPart;

            //Чтение пути сохранения открытой сборки 
            _pathFromProject = swModel.GetPathName().ToLower();

            //Проверка наличия сборки
            if (_pathFromProject == String.Empty)
            {
                //Сохранение нвоого файла сборки
                InitialPath = InitializeFolder();
                if (InitialPath == null)
                    return;
                else
                {
                    //Сохранение файла новой сборки
                    sldobj.saveInitialAssembly(InitialPath);
                    InitialPath = InitialPath.ToLower();
                    //Сохранение пути файла новой сборки
                    _pathFromProject = InitialPath;
                    //Место сохранения сборки
                    InitialPath = InitialPath.Replace("\\" + swModel.GetTitle().ToLower() + ".sldasm", "");
                }
            }
            else
            {
                //Сохранение пути сборки если она существует
                InitialPath = _pathFromProject.Replace("\\" + swModel.GetTitle().ToLower() + ".sldasm", "");
            }

            //Инициализация компонентов формы 
            InitializeComponent();

            //Контекст данных программы
            validation = new DataValidation(this);
            DataContext = validation;

            //Заполнение полей программы начальными данными
            InitializeStartData();
            //Запуск расчета с начальными данными
            InitializeCalculate();
        }

        /// <summary>
        /// Иницализация начальной папки для сохранения компонентов
        /// </summary>
        private string InitializeFolder()
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "Assembly File (*.sldasm)|*.sldasm";
            sfd.FileName = "NewAssembly.sldasm"; ;

            //Обработка нажатия в диалогвоом окне кнопки "ОК"
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                return sfd.FileName;
            else
                return null;
        }

        /// <summary>
        ///Инициализация начальных данных, заполнение полей
        /// </summary>
        private void InitializeStartData()
        {
            //Начальные размеры  
            this.Width = 705;
            this.Height = 560;

            //Добавление в комбо передаточного отношения
            string PeredatValue = "20.0000";
            string[] ItemsPeredat = { "5.6000", "6.3000", "7.1000",
            "8.0000", "9.0000", "10.0000", "11.2000", "12.5000",
            "14.0000", "16.0000", "18.0000", "20.0000", "22.4000",
            "25.0000", "28.0000", "31.5000", "40.0000", "45.0000",
            "50.0000", "56.0000", "63.0000", "71.0000", "80.0000", "90.0000", "100.0000"};
            foreach (string element in ItemsPeredat)
                Peredat_combo.Items.Add(element);
            Peredat_combo.SelectedValue = PeredatValue;

            //Добавление в комбо модуля
            string ModuleValue = "4.000 мм";
            string[] ItemsModule = { "1.000 мм", "1.150 мм", "1.250 мм",
            "1.400 мм", "1.600 мм", "1.800 мм", "2.000 мм", "2.240 мм",
            "2.500 мм", "2.800 мм", "3.150 мм", "3.550 мм", "4.000 мм",
            "4.500 мм", "5.000 мм", "5.600 мм", "6.300 мм", "7.100 мм",
            "8.000 мм", "9.000 мм", "10.000 мм", "11.200 мм", "12.500 мм", "14.000 мм", "16.000 мм", "18.000 мм"};
            foreach (string element in ItemsModule)
                Module_combo.Items.Add(element);
            Module_combo.SelectedValue = ModuleValue;

            //Добавление значений в комбо угла профиля
            string PressureAngleValue = "20.0000 град";
            string[] ItemsPresAngle = { "14.5000 град", "17.5000 град", "20.0000 град", "22.5000 град", "25.0000 град", "30.0000 град" };
            foreach (string element in ItemsPresAngle)
                ProfileDeg_combo.Items.Add(element);
            ProfileDeg_combo.SelectedValue = PressureAngleValue;

            //Добавление в комбо построения червяка 
            Worm_combo.Items.Add("Построить");
            Worm_combo.Items.Add("Не строить");
            Worm_combo.SelectedIndex = 0;
            //Добавление в комбо построения колеса
            Gear_combo.Items.Add("Построить");
            Gear_combo.Items.Add("Не строить");
            Gear_combo.SelectedIndex = 0;

            //Значения червяка
            Kol_vitkovValue.Text = validation.kol_vitkov.ToString();
            LengthValue.Text = validation.length_Worm.ToString();
            Koef_diamValue.Text = validation.koef_diamWorm.ToString();

            //Значения колеса
            Teeth_gearValue.Text = validation.teeth_Gear.ToString();
            Width_gearValue.Text = validation.width_Gear.ToString();
            Koef_smeshValue.Text = validation.koef_Smesh.ToString();
            Hole_widthValue.Text = validation.hole_Gear.ToString();

            Hole_widthValue.Visibility = Visibility.Hidden;

            //Значения для силового расчета
            PowerValue.Text = validation.PowerWorm.ToString(); ;
            VelocityValue.Text = validation.VelocityWorm.ToString(); ;

            //Заполнение материалов
            MatObj = new MaterialHelper(swApp);
            var lister = MatObj.GetMaterials();
            lister.Insert(0, new Material
            {
                Name = "Нет материала",
                Classification = "Не задано",
                DatabaseFileFound = false,
                Elastic_modulus = "206000000000",
                Poisson_ratio = "0.3",
                Tensile_strength = "250000000",
                Yield_strength = "160000000"
            });

            MatGear_combo.ItemsSource = MatWorm_combo.ItemsSource = lister;
            MatWorm_combo.DisplayMemberPath = MatGear_combo.DisplayMemberPath = "DisplayName";
            MatGear_combo.SelectedIndex = 0;
            MatWorm_combo.SelectedIndex = 0;


            //Значения силовых характеристик
            KPDValue.IsEnabled = false;

            KvValue.Text = validation.Kv;
            timeValue.Text = validation.time;

            radioParam1.IsChecked = true;
            radioWorm.IsChecked = true;

            radioWormSel.IsChecked = true;
            radioWormSel.Checked += RadioSelectComp_Checked;
            radioGearSel.Checked += RadioSelectComp_Checked;

            RadioType_Calc.Checked += RadioSelectType_Checked;
            RadioType_Calc1.Checked += RadioSelectType_Checked;
            RadioType_Calc2.Checked += RadioSelectType_Checked;
            RadioType_Calc.IsChecked = true;

        }

        /// <summary>
        /// Расчет параметров и их отображение в таблицах на вкладках "Модель" и "Расчет"
        /// </summary>
        private void InitializeCalculate()
        {
            //Значение модуля
            Module = float.Parse(((Module_combo.Text).Split(' ')[0]));

            //Значение угла профиля
            PressureAngle = float.Parse(((ProfileDeg_combo.Text).Split(' ')[0]));

            //Количество оборотов витков
            Kol_oborotovValue.Text = (Math.Floor(float.Parse(LengthValue.Text) / (Module * Math.PI)) + 3).ToString();

            //Расчет значений передаточного числа и зубьев колеса
            if (radioParam.IsChecked == true)
            {
                Peredat = float.Parse(Peredat_combo.Text);
                Teeth_gearValue.Text = (float.Parse(Peredat_combo.Text) * float.Parse(Kol_vitkovValue.Text)).ToString("0");
            }
            else
            {
                Peredat_Changed.Text = (float.Parse(Teeth_gearValue.Text) / float.Parse(Kol_vitkovValue.Text)).ToString("0.00");
                Peredat = float.Parse(Peredat_Changed.Text);
            }
            //Средний диаметр червяка
            //Угол наклона зуба
            //Коэффициент диаметра червяка 
            if (radioWorm.IsChecked == true)
            {
                Av_diamValue.Text = (float.Parse(Koef_diamValue.Text) * Module).ToString(".00");
                DegTeethValue.Text = (Math.Atan(float.Parse(Kol_vitkovValue.Text)
                    / float.Parse(Koef_diamValue.Text)) * 180 / Math.PI).ToString(".00");
                validation.deg_TeethValue = float.Parse(DegTeethValue.Text);

            }
            else if (radioWorm1.IsChecked == true)
            {
                Koef_diamValue.Text = ((float.Parse(Kol_vitkovValue.Text))
                    / (Math.Tan((float.Parse(DegTeethValue.Text) * Math.PI) / 180))).ToString(".00");
                Av_diamValue.Text = (float.Parse(Koef_diamValue.Text) * Module).ToString(".00");
            }
            else if (radioWorm2.IsChecked == true)
            {
                Koef_diamValue.Text = (float.Parse(Av_diamValue.Text) / Module).ToString(".00");
                DegTeethValue.Text = (Math.Atan(float.Parse(Kol_vitkovValue.Text)
                    / float.Parse(Koef_diamValue.Text)) * 180 / Math.PI).ToString(".00");
                validation.deg_TeethValue = float.Parse(DegTeethValue.Text);
            }

            //Делительные диаметры
            d1 = float.Parse(Koef_diamValue.Text) * Module;
            d2 = float.Parse(Teeth_gearValue.Text) * Module;

            //Начальные диаметры (dw)
            dw1 = Module * (float.Parse(Koef_diamValue.Text) + 2 * float.Parse(Koef_smeshValue.Text));
            dw2 = d2;

            //Диаметры вершин (da)
            da1 = d1 + (2 * Module);
            da2 = d2 + 2 * (1 + float.Parse(Koef_smeshValue.Text)) * Module;

            //Диаметры впадин (df)
            df1 = d1 - (float)2.4 * Module;
            df2 = d2 - 2 * ((float)1.2 - float.Parse(Koef_smeshValue.Text)) * Module;

            //Наибольший диаметр (dae)
            dae2 = da2 + (6 * Module / (float.Parse(Kol_vitkovValue.Text) + 2));

            //Угол зацепления (a)
            float tan_on_cos = (float)(Math.Tan(PressureAngle * Math.PI / 180) * Math.Cos(float.Parse(DegTeethValue.Text) * Math.PI / 180));
            Alpha = (float)(Math.Atan(tan_on_cos) * 180 / Math.PI);

            //Высота витков червяка
            h1 = (float)Math.Round(2.2 * Module);
            ha1 = Module;

            //Межосевое расстояние (aw)
            aw = Module * (float.Parse(Teeth_gearValue.Text) + float.Parse(Koef_diamValue.Text) + 2 * float.Parse(Koef_smeshValue.Text)) / 2;

            //КПД
            float v1 = (float)((Math.PI * dw1 * float.Parse(VelocityValue.Text)) / 60000);
            //Делительный угол подъема
            float tgy = (float)(Math.Atan((float.Parse(Kol_vitkovValue.Text)
                / float.Parse(Koef_diamValue.Text))) * 180 / Math.PI);
            //Начальный угол подъема
            float tgyW = (float)(Math.Atan(float.Parse(Kol_vitkovValue.Text)
                / (float.Parse(Koef_diamValue.Text) + 2 * float.Parse(Koef_smeshValue.Text))) * 180 / Math.PI);

            float v2 = v1 * (float.Parse(Kol_vitkovValue.Text) / float.Parse(Koef_diamValue.Text));

            //Скорость скольжения
            float vk = (float)(v1 / Math.Cos(tgyW * Math.PI / 180));

            float phiz = (float)(Math.Atan(0.02 + (0.03) / (vk)) * 180 / Math.PI);
            //float Kv = (1200 + v2) / (1200);
            float KPD;
            if (boolKPD.IsChecked == true)
                KPD = float.Parse(KPDValue.Text);
            else
                KPD = (float)(((Math.Tan(tgyW * Math.PI / 180)) / (Math.Tan((tgyW + phiz) * Math.PI / 180))) * 0.96);
            KPDValue.Text = KPD.ToString("0.000");
            validation.KPD = KPDValue.Text;

            //значения на червяке
            float Power = float.Parse(PowerValue.Text);
            float Velocity = float.Parse(VelocityValue.Text);
            float Moment = 0;

            //значения на колесе
            float Velocity_WG = 0;
            float Power_WG = 0;
            float Moment_WG = 0;

            //Червяк
            if (radioWormSel.IsChecked == true & RadioType_Calc.IsChecked == true)
            {
                Power = float.Parse(PowerValue.Text);
                Velocity = float.Parse(VelocityValue.Text);
                Moment = (float)((60000 * Power) / (2 * Math.PI * Velocity));
                Power_WG = Power * KPD;
                Velocity_WG = Velocity / Peredat;
                Moment_WG = (float)((60000 * Power_WG) / (2 * Math.PI * Velocity_WG));

            }
            else if (radioWormSel.IsChecked == true & RadioType_Calc1.IsChecked == true)
            {
                Velocity = float.Parse(VelocityValue.Text);
                Moment = float.Parse(MomentValue.Text);
                Moment_WG = Moment * Peredat * KPD;
                Velocity_WG = Velocity / Peredat;
                Power = (float)((Math.PI * Velocity_WG * Moment_WG) / (30 * KPD)) / 1000;
                Power_WG = Power * KPD;
            }
            else if (radioWormSel.IsChecked == true & RadioType_Calc2.IsChecked == true)
            {
                Power = float.Parse(PowerValue.Text);
                Moment = float.Parse(MomentValue.Text);
                Velocity = (float)((60000 * Power) / (2 * Math.PI * Moment));
                Velocity_WG = Velocity / Peredat;
                Power_WG = Power * KPD;
                Moment_WG = Moment * Peredat * KPD;
            }

            //Колесо
            else if (radioGearSel.IsChecked == true & RadioType_Calc.IsChecked == true)
            {
                Power_WG = float.Parse(PowerValue_Gear.Text);
                Velocity_WG = float.Parse(VelocityValue_Gear.Text);
                //Power = Power_WG * KPD;
                Power = Power_WG / KPD;
                Velocity = Velocity_WG * Peredat;
                Moment_WG = (float)((60000 * Power_WG) / (2 * Math.PI * Velocity_WG));
                Moment = (float)((60000 * Power) / (2 * Math.PI * Velocity));

            }
            else if (radioGearSel.IsChecked == true & RadioType_Calc1.IsChecked == true)
            {
                Velocity_WG = float.Parse(VelocityValue_Gear.Text);
                Moment_WG = float.Parse(MomentValue_Gear.Text);
                Power_WG = (float)((Math.PI * Velocity_WG * Moment_WG) / (30)) / 1000;
                Moment = (Moment_WG / Peredat) * KPD;
                Velocity = Velocity_WG * Peredat;
               // Power = Power_WG * KPD;
                Power = Power_WG / KPD;
            }
            else if (radioGearSel.IsChecked == true & RadioType_Calc2.IsChecked == true)
            {
                Power_WG = float.Parse(PowerValue_Gear.Text);
                Moment_WG = float.Parse(MomentValue_Gear.Text);
                Velocity_WG = (float)((60000 * Power_WG) / (2 * Math.PI * Moment_WG));
                //Power = Power_WG * KPD;
                Power = Power_WG / KPD;
                Velocity = Velocity_WG * Peredat;
                Moment = (Moment_WG / Peredat) * KPD;
            }

            VelocityValue.Text = Velocity.ToString("0.00");
            validation.VelocityWorm = VelocityValue.Text;

            PowerValue.Text = Power.ToString("0.000");
            validation.PowerWorm = PowerValue.Text;

            MomentValue.Text = Moment.ToString("0.000");
            validation.MomentWorm = MomentValue.Text;

            VelocityValue_Gear.Text = Velocity_WG.ToString("0.00");
            validation.VelocityGear = VelocityValue_Gear.Text;

            PowerValue_Gear.Text = Power_WG.ToString("0.000");
            validation.PowerGear = PowerValue_Gear.Text;

            MomentValue_Gear.Text = Moment_WG.ToString("0.000");
            validation.MomentGear = MomentValue_Gear.Text;

            //окружная сила на червяке, Осевая на колесе
            float Ft1 = (2000 * float.Parse(MomentValue.Text)) / (dw1);
            float Fa2 = Ft1;

            //осевая сила на червяке, окружная на колесе
            float Fa1 = (2000 * Moment_WG) / (dw2);
            float Ft2 = Fa1;

            //Радиальная сила передачи
            float Fr = (float)(Ft2 * Math.Tan(PressureAngle * (Math.PI / 180)));

            //Нормальная сила передачи
            float Fn = (float)(Ft2 / (Math.Cos(PressureAngle * (Math.PI / 180)) * Math.Cos(tgy * (Math.PI / 180))));

            //Коэффициент для вычисления контактного напряжения
            float Koef = 0;
            if (vk >= 10)
                Koef = (float)1.2;
            else if (vk < 10 & vk > 5)
                Koef = (float)1.1;
            else if (vk <= 5)
                Koef = 1;
            //Контактное напряжение
            float contact_stress = (float)((170 / (float.Parse(Teeth_gearValue.Text) / float.Parse(Koef_diamValue.Text))) *
                Math.Sqrt((float.Parse(MomentValue_Gear.Text) * 1000 * float.Parse(KvValue.Text))
                * Math.Pow(((1 + float.Parse(Teeth_gearValue.Text) / float.Parse(Koef_diamValue.Text)) / (aw)), 3)));
            validation.contact_calc = contact_stress.ToString("0.00");

            //Коэффициент для вычисления напряжения изгиба
            float zv2 = (float)((float.Parse(Teeth_gearValue.Text)) / (Math.Pow(Math.Cos(tgyW * Math.PI / 180), 3)));
            float Yf2 = 0;
            if (zv2 < 37)
                Yf2 = (float)(2.4 - 0.0214 * zv2);
            else if (zv2 >= 37 & zv2 <= 45)
                Yf2 = (float)(2.21 - 0.0162 * zv2);
            else if (zv2 > 45)
                Yf2 = (float)(1.72 - 0.0053 * zv2);
            //Напряжение изгиба
            float bending_stress = (float)(((0.7 * Ft2 * Koef) / (float.Parse(Width_gearValue.Text) * Module * Math.Cos(tgyW * Math.PI / 180))) * Yf2);
            validation.bending_calc = bending_stress.ToString("0.00");

            //Число циклов нагружения
            float Nk = 60 * float.Parse(VelocityValue_Gear.Text) * float.Parse(timeValue.Text);

            //Коэффициент долговечности
            float Khl = (float)(Math.Pow(((Math.Pow(10, 7)) / (Nk)), 0.125));

            //ожидаемое значение скорости
            float v_wait = (float)((4.5 / 10000) * float.Parse(VelocityValue.Text) * Math.Pow(float.Parse(MomentValue_Gear.Text), 1.0 / 3));

            //коэф изнашивания зубьев
            float Cv = (float)(1.46 - (v_wait / 7.29) * (1 - (v_wait / 20.2)));

            //Допускаемые контактные напряжения
            limit_contact.Text = (float.Parse(sigmaV.Text) * 0.82 * Cv * Khl).ToString("0.00");
            validation.contact = limit_contact.Text;

            //предел выносливости зубьев при изгибе
            float limit_bending_Koef = (float)(0.08 * float.Parse(sigmaV.Text) + 0.25 * float.Parse(sigmaT.Text));
            float Yhl = (float)(Math.Pow(((Math.Pow(10, 6)) / (Nk)), 1.0 / 9));
            limit_bending.Text = (limit_bending_Koef * Yhl).ToString("0.00");
            validation.bending = limit_bending.Text;

            //Температура масла при работе
            float temperature_oil = (float)((float.Parse(PowerValue.Text) * 1000 * (1 - KPD)) / (17 * 12.2 * Math.Pow(aw / 1000, 1.71))) + 20;
            temperature.Text = temperature_oil.ToString("0.00");
            validation.temperature = temperature.Text;

            //Отображение расчетных данных в таблицах
            RefreshTable_Model(aw, Module, Alpha, da1, d1, df1, h1, ha1, da2, d2, df2, dae2);
            RefreshTable_Calc(Fr, vk, Khl, Ft1, Fa1, Ft2, Fa2, Fn, contact_stress, bending_stress);

            //Изменение цвета заднего фона на белый
            TableGeneral_Model.RowBackground = Brushes.White;
            TableWorm_Model.RowBackground = Brushes.White;
            TableGear_Model.RowBackground = Brushes.White;
            TableGeneral_Calc.RowBackground = Brushes.White;
            TableWorm_Calc.RowBackground = Brushes.White;
            TableGear_Calc.RowBackground = Brushes.White;

        }

        /// <summary>
        /// Отображение данных в таблицах на вкладке "Модель"
        /// </summary>
        private void RefreshTable_Model(float aw_Table, float Module_Table, float Alpha_Table,
            float DaWorm_Table, float DWorm_Table, float DfWorm_Table, float h1_Table, float h1a_Table, float DaGear_Table, float DGear_Table, float DfGear_Table, float DaeGear_Table)
        {

            ObservableCollection<Parameter_values> general = new ObservableCollection<Parameter_values>();
            general.Add(new Parameter_values { parameter = "Межосев. расст. (aw):", value = aw_Table.ToString("0.00") + " мм" });
            general.Add(new Parameter_values { parameter = "Модуль (m):", value = Module_Table.ToString("0.00") + " мм" });
            general.Add(new Parameter_values { parameter = "Угол профиля (α):", value = Alpha_Table.ToString("0.00") + " град" });
            TableGeneral_Model.ItemsSource = general;
            TableGeneral_Model.RowHeight = TableGeneral_Model.Height / general.Count;

            ObservableCollection<Parameter_values> worm = new ObservableCollection<Parameter_values>();
            worm.Add(new Parameter_values { parameter = "Наружный диаметр (da):", value = DaWorm_Table.ToString("0.00") + " мм" });
            worm.Add(new Parameter_values { parameter = "Средний диаметр (d):", value = DWorm_Table.ToString("0.00") + " мм" });
            worm.Add(new Parameter_values { parameter = "Диаметр впадин (df):", value = DfWorm_Table.ToString("0.00") + " мм" });
            worm.Add(new Parameter_values { parameter = "Высота витка (h1):", value = h1_Table.ToString("0.00") + " мм" });
            worm.Add(new Parameter_values { parameter = "Высота гол. витка (ha1):", value = h1a_Table.ToString("0.00") + " мм" });
            TableWorm_Model.ItemsSource = worm;
            TableWorm_Model.RowHeight = TableWorm_Model.Height / worm.Count;

            ObservableCollection<Parameter_values> gear = new ObservableCollection<Parameter_values>();
            gear.Add(new Parameter_values { parameter = "Наиб. диаметр (dae2):", value = DaeGear_Table.ToString("0.00") + " мм" });
            gear.Add(new Parameter_values { parameter = "Наружный диаметр (da):", value = DaGear_Table.ToString("0.00") + " мм" });
            gear.Add(new Parameter_values { parameter = "Средний диаметр (d):", value = DGear_Table.ToString("0.00") + " мм" });
            gear.Add(new Parameter_values { parameter = "Диаметр впадин (df):", value = DfGear_Table.ToString("0.00") + " мм" });
            TableGear_Model.ItemsSource = gear;
            TableGear_Model.RowHeight = TableGear_Model.Height / gear.Count;
        }

        /// <summary>
        /// Отображение данных в таблицах на вкладке "Расчет"
        /// </summary>
        private void RefreshTable_Calc(float Fr_Table, float Vk_Table, float Khl_Table, float Ft1_Table,
            float Fa1_Table, float Ft2_Table, float Fa2_Table, float Fn_Table, float contactPres_Table, float bending_Table)
        {

            ObservableCollection<Parameter_values> general = new ObservableCollection<Parameter_values>();
            general.Add(new Parameter_values { parameter = "Радиальная сила (Fr):", value = Fr_Table.ToString("0.00") + " Н" });
            general.Add(new Parameter_values { parameter = "Нормальная сила (Fn):", value = Fn_Table.ToString("0.00") + " Н" });
            general.Add(new Parameter_values { parameter = "Скорость скольжения (Vk):", value = Vk_Table.ToString("0.00") + " м/c" });
            general.Add(new Parameter_values { parameter = "Коэф. долговечности:", value = Khl_Table.ToString("0.00") });
            TableGeneral_Calc.ItemsSource = general;
            TableGeneral_Calc.RowHeight = TableGeneral_Calc.Height / general.Count;

            ObservableCollection<Parameter_values> worm = new ObservableCollection<Parameter_values>();
            worm.Add(new Parameter_values { parameter = "Окружная сила (Ft):", value = Ft1_Table.ToString("0.00") + " Н" });
            worm.Add(new Parameter_values { parameter = "Осевая сила (Fa):", value = Fa1_Table.ToString("0.00") + " Н" });
            TableWorm_Calc.ItemsSource = worm;
            TableWorm_Calc.RowHeight = TableWorm_Calc.Height / worm.Count;

            ObservableCollection<Parameter_values> gear = new ObservableCollection<Parameter_values>();
            gear.Add(new Parameter_values { parameter = "Окружная сила (Ft):", value = Ft2_Table.ToString("0.00") + " Н" });
            gear.Add(new Parameter_values { parameter = "Осевая сила (Fa):", value = Fa2_Table.ToString("0.00") + " Н" });
            gear.Add(new Parameter_values { parameter = "Контактное напряжение:", value = contactPres_Table.ToString("0.00") + " МПа" });
            gear.Add(new Parameter_values { parameter = "Напряжение изгиба:", value = bending_Table.ToString("0.00") + " МПа" });
            TableGear_Calc.ItemsSource = gear;
            TableGear_Calc.RowHeight = TableGear_Calc.Height / gear.Count;
        }

        /// <summary>
        /// Выбор компонента на вклаке "Расчет"
        /// </summary>
        private void RadioSelectComp_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            string name = pressed.Content.ToString();
            var tag = int.Parse((panelType.Children.OfType<RadioButton>().FirstOrDefault(r => (bool)r.IsChecked)).Tag.ToString());

            PowerValue_Gear.IsEnabled = false;
            VelocityValue_Gear.IsEnabled = false;
            MomentValue_Gear.IsEnabled = false;
            PowerValue.IsEnabled = false;
            VelocityValue.IsEnabled = false;
            MomentValue.IsEnabled = false;

            //Червяк
            if (name == "Червяк" & tag == 0)
            {
                PowerValue.IsEnabled = true;
                VelocityValue.IsEnabled = true;
                MomentValue.IsEnabled = false;
            }
            else if (name == "Червяк" & tag == 1)
            {
                PowerValue.IsEnabled = false;
                VelocityValue.IsEnabled = true;
                MomentValue.IsEnabled = true;
            }
            else if (name == "Червяк" & tag == 2)
            {
                PowerValue.IsEnabled = true;
                VelocityValue.IsEnabled = false;
                MomentValue.IsEnabled = true;
            }

            //Колесо
            if (name == "Червячное колесо" & tag == 0)
            {
                PowerValue_Gear.IsEnabled = true;
                VelocityValue_Gear.IsEnabled = true;
                MomentValue_Gear.IsEnabled = false;
            }
            else if (name == "Червячное колесо" & tag == 1)
            {
                PowerValue_Gear.IsEnabled = false;
                VelocityValue_Gear.IsEnabled = true;
                MomentValue_Gear.IsEnabled = true;
            }
            else if (name == "Червячное колесо" & tag == 2)
            {
                PowerValue_Gear.IsEnabled = true;
                VelocityValue_Gear.IsEnabled = false;
                MomentValue_Gear.IsEnabled = true;
            }
        }

        /// <summary>
        /// Выбор типа расчета на вкладе "расчет"
        /// </summary>
        private void RadioSelectType_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            var tag = pressed.Content.ToString();
            string name = (panelGeneral.Children.OfType<RadioButton>().FirstOrDefault(r => (bool)r.IsChecked)).Content.ToString();

            PowerValue_Gear.IsEnabled = false;
            VelocityValue_Gear.IsEnabled = false;
            MomentValue_Gear.IsEnabled = false;
            PowerValue.IsEnabled = false;
            VelocityValue.IsEnabled = false;
            MomentValue.IsEnabled = false;

            //Червяк
            if (name == "Червяк" & tag == "Мощность, Скорость -> Момент")
            {
                PowerValue.IsEnabled = true;
                VelocityValue.IsEnabled = true;
                MomentValue.IsEnabled = false;
            }
            else if (name == "Червяк" & tag == "Момент, Скорость -> Мощность")
            {
                PowerValue.IsEnabled = false;
                VelocityValue.IsEnabled = true;
                MomentValue.IsEnabled = true;
            }
            else if (name == "Червяк" & tag == "Мощность, Момент -> Скорость")
            {
                PowerValue.IsEnabled = true;
                VelocityValue.IsEnabled = false;
                MomentValue.IsEnabled = true;
            }

            //Колесо
            if (name == "Червячное колесо" & tag == "Мощность, Скорость -> Момент")
            {
                PowerValue_Gear.IsEnabled = true;
                VelocityValue_Gear.IsEnabled = true;
                MomentValue_Gear.IsEnabled = false;
            }
            else if (name == "Червячное колесо" & tag == "Момент, Скорость -> Мощность")
            {
                PowerValue_Gear.IsEnabled = false;
                VelocityValue_Gear.IsEnabled = true;
                MomentValue_Gear.IsEnabled = true;
            }
            else if (name == "Червячное колесо" & tag == "Мощность, Момент -> Скорость")
            {
                PowerValue_Gear.IsEnabled = true;
                VelocityValue_Gear.IsEnabled = false;
                MomentValue_Gear.IsEnabled = true;
            }
        }

        /// <summary>
        /// Вызво изменения цвета таблицы
        /// </summary>
        private void data_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (((TextBox)sender).IsFocused)
                data_Touched();

        }

        /// <summary>
        /// Обработка изменения какого-либо значения в выпадающих списках
        /// </summary>
        private void combo_SelectionChanged(object sender, SelectionChangedEventArgs e) { data_Touched(); }

        /// <summary>
        /// Изменение цвета таблиц
        /// </summary>
        private void data_Touched()
        {
            var color = Brushes.LightGray;
            TableGeneral_Model.RowBackground = color;
            TableWorm_Model.RowBackground = color;
            TableGear_Model.RowBackground = color;
            TableGeneral_Calc.RowBackground = color;
            TableWorm_Calc.RowBackground = color;
            TableGear_Calc.RowBackground = color;
        }

        /// <summary>
        /// Обработка выбора материала из combobox "Материал червяка"
        /// </summary>
        private void MatWorm_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Считывание выбранного значения из списка как обьект класса Material
            Material item = (Material)MatWorm_combo.SelectedItem;

            //Обработка характеристик выбранного материала
            if (item.Elastic_modulus == null)
                E_wormValue.Text = "0";
            else
                E_wormValue.Text = (float.Parse(item.Elastic_modulus) / 1000000).ToString("0.00");

            if (item.Poisson_ratio == null)
                Puasson_wormValue.Text = "0";
            else
                Puasson_wormValue.Text = item.Poisson_ratio;

            //Изменение цвета таблиц
            data_Touched();
        }

        /// <summary>
        /// Обработка выбора материала из combobox "Материал червячного колеса"
        /// </summary>
        private void MatGear_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Считывание выбранного значения из списка как обьект класса Material
            Material item = (Material)MatGear_combo.SelectedItem;

            //Обработка характеристик выбранного материала
            if (item.Tensile_strength == null)
                sigmaV.Text = "0";
            else
                sigmaV.Text = (float.Parse(item.Tensile_strength) / 1000000).ToString("0.00");

            if (item.Yield_strength == null)
                sigmaT.Text = "0";
            else
                sigmaT.Text = (float.Parse(item.Yield_strength) / 1000000).ToString("0.00");

            if (item.Elastic_modulus == null)
                E_gearValue.Text = "0";
            else
                E_gearValue.Text = (float.Parse(item.Elastic_modulus) / 1000000).ToString("0.00");

            if (item.Poisson_ratio == null)
                Puasson_gearValue.Text = "0";
            else
                Puasson_gearValue.Text = item.Poisson_ratio;

            //Изменение цвета таблиц
            data_Touched();
        }

        /// <summary>
        /// Расчет ширины венца зубчатого колеса
        /// </summary>
        private void calc_WidthGear_Click(object sender, RoutedEventArgs e)
        {
            var teeth = float.Parse(Kol_vitkovValue.Text);
            if (teeth <= 3)
                Width_gearValue.Text = (0.75 * da1).ToString("0.00");
            else if (teeth == 4)
                Width_gearValue.Text = (0.67 * da1).ToString("0.00");

            data_Touched();
        }

        /// <summary>
        /// Расчет длины червяка
        /// </summary>
        private void calc_LengthWorm_Click(object sender, RoutedEventArgs e)
        {
            LengthValue.Text = (2 * Math.Sqrt(Math.Pow(dae2 / 2, 2) -
                Math.Pow((aw - da1 / 2), 2) + (Math.PI * Module / 2))).ToString("0.00");
            data_Touched();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {

            if (this.Height == 660)
            {
                ExpanderModel.IsExpanded = true;
                ExpanderCalculations.IsExpanded = true;
            }
            else if (this.Height == 550.4)
            {
                ExpanderModel.IsExpanded = false;
                ExpanderCalculations.IsExpanded = false;
            }

            if (this.Width == 932)
            {
                var bc = new BrushConverter();
                var color = bc.ConvertFromString("#FFCBE4FF");
                Expander_showCalc.Background = (Brush)color;
                Expander_showCalc1.Background = (Brush)color;
            }
            else if (this.Width == 704)
            {
                Expander_showCalc.Background = Brushes.White;
                Expander_showCalc1.Background = Brushes.White;
            }

        }

        /// <summary>
        /// Подтверждение построения сборки
        /// </summary>
        private void confirm_BuildClick(object sender, RoutedEventArgs e)
        {
            //Обработка построения компонентов
            bool isWorm = Worm_combo.Text == "Построить" ? true : false;
            bool isGear = Gear_combo.Text == "Построить" ? true : false;

            this.Hide();
            //Обработка значений построения компонентов
            if (isWorm == false & isGear == false)
            {
                return;
            }
            else
            {
                //Передача праметров в окно "Подтверждение"
                foldersWindow = new ConfirmFolders(InitialPath, isWorm, isGear);
                foldersWindow.Topmost = true;
                foldersWindow.ShowDialog();

                //Обработка подтверждения построения
                if (buildAccepted == true)
                    create_components(isWorm, isGear);
                else
                {
                    this.Show();
                    return;
                }
            }

        }

        /// <summary>
        /// Построение компонентов, сборки, создание зависимостей
        /// </summary>
        private void create_components(bool isWorm, bool isGear)
        {
            try
            {
                //Считывание значений направления зубьев
                int rightOrLeft = radioLeft.IsChecked == true ? 0 : 1;
                //Считывание значения отверстия червячного оклеса
                float hole_diameter = Hole_bool.IsChecked == true ? hole_diameter = float.Parse(Hole_widthValue.Text) / 2 : 0;
                //Считывание выбранных материалов
                Material materialWorm = MatWorm_combo.SelectedIndex != 0 ? (Material)MatWorm_combo.SelectedItem : null;
                Material materialGear = MatGear_combo.SelectedIndex != 0 ? (Material)MatGear_combo.SelectedItem : null;

                //Инициализация объекта класса червяка
                worm = new Worm();
                //Построение червяка
                if (isWorm)
                {
                    //передача параметров червяка
                    worm._da = da1;
                    worm._d = d1;
                    worm._df = df1;
                    worm._pressureAngle = PressureAngle;
                    worm._length = float.Parse(LengthValue.Text);
                    worm._module = Module;
                    worm._z = float.Parse(Kol_vitkovValue.Text);
                    worm._rightOrLeft = rightOrLeft;
                    worm._aw = aw;
                    worm._path = _pathWorm;
                    worm._material = materialWorm;
                    worm.create();
                }
                //Инициализация объекта класса червячного колеса
                gear = new Gear();
                //Построение червячного колеса
                if (isGear)
                {
                    //передача параметров колеса
                    gear._aw = aw;
                    gear._b = float.Parse(Width_gearValue.Text);
                    gear._z = float.Parse(Teeth_gearValue.Text);
                    gear._df = df2;
                    gear._da = da2;
                    gear._beta = float.Parse(DegTeethValue.Text);
                    gear._pressureAngle = PressureAngle;
                    gear._df1 = df1;
                    gear._da1 = da1;
                    gear._d1 = d1;
                    gear._px = (float)(Module * Math.PI);
                    gear._z1 = float.Parse(Kol_vitkovValue.Text);
                    gear._dw = dw2;
                    gear._dae = dae2;
                    gear._rightOrLeft = rightOrLeft;
                    gear._path = _pathGear;
                    gear._hole_diameter = hole_diameter;
                    gear._material = materialGear;
                    gear.create();
                }
                //Проверка построения компонентов, создание сборки
                //создаем сборку
                sldobj.CreateAssembly(_pathAssembly);

                //добавляем компоненты
                sldobj.AddComponents(worm, gear, _pathAssembly);

                if (isWorm & isGear)
                {
                    //добавление зависимостей
                    var teethGear = float.Parse(Teeth_gearValue.Text);
                    var vitWorm = float.Parse(Kol_vitkovValue.Text);
                    //если строим оба - делаем, если нет - пропускаем, УЧИТЫВАТЬ, ЕСЛИ ВЫБИРАЕМ ОБЕ ГРАНИ - УБИРАЕМ ЗАВИСИМОСТИ (ПЕРЕДАЕМ СЛОВАРЬ)
                    sldobj.AddMates(_pathAssembly, _nameAssembly, _nameWorm, _nameGear, aw, teethGear, vitWorm, rightOrLeft);
                }
                //Добавление в текущую сборку
                sldobj.addToAssembly(_pathFromProject, _pathAssembly);

                //Создание зависимостей внутри текущей сборки
                // изменение условия(если была выбрана только одна цилиндр грань или какой-либо фейс)
                //параметры из словаря со значением червяка
                //параметры из словаря со значением колеса
                //параметры фейса
                //Если выбрана цилиндрическая грань, то передаю туда параметры из словаря
                string GearOrWorm = String.Empty;
                if (((selectCylindrical_Gear.IsChecked == true) || (selectCylindrical_Worm.IsChecked == true)
                    || (selectFace_Gear.IsChecked == true) || (selectFace_Worm.IsChecked == true)) & swFaceObject != null)
                {
                    if (selectCylindrical_Worm.IsChecked == true)
                        GearOrWorm = "WormCylinder";
                    else if (selectCylindrical_Gear.IsChecked == true)
                        GearOrWorm = "GearCylinder";
                    else if (selectFace_Gear.IsChecked == true)
                        GearOrWorm = "GearFace";
                    else if (selectFace_Worm.IsChecked == true)
                        GearOrWorm = "WormFace";

                    sldobj.addMatesToFaces(selectedComponent, selectedEntity,
                        _pathFromProject, _nameAssembly, _nameWorm, _nameGear, GearOrWorm);
                }

                _tokenSourceCylinder.Cancel();
                _tokenSourceFace.Cancel();
                Directory.SetCurrentDirectory(baseDirectory);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка во время построения: " + ex.Message);
                this.Show();
            }
        }

        /// <summary>
        /// Изменение отображения поля для значения ширины отверстия колеса
        /// </summary>
        private void CheckBox_Hole_Changed(object sender, RoutedEventArgs e)
        {
            if (Hole_bool.IsChecked == true)
                Hole_widthValue.Visibility = Visibility.Visible;
            else
                Hole_widthValue.Visibility = Visibility.Hidden;
        }

        /// <summary>
        ///Изменение типа расчета КПД
        /// </summary>
        private void CheckBox_KPD_Changed(object sender, RoutedEventArgs e)
        {
            if (boolKPD.IsChecked == true)
                KPDValue.IsEnabled = true;
            else
                KPDValue.IsEnabled = false;
        }

        private void Expander_Expanded(object sender, RoutedEventArgs e) { this.Height = 660; }

        private void Expander_Collapsed(object sender, RoutedEventArgs e) { this.Height = 550; }

        /// <summary>
        ///Отображение/скрытие таблиц
        /// </summary>
        private void Expander_Calculations_Click(object sender, RoutedEventArgs e)
        {
            if (this.Width == 932)
            {
                this.Width = 705;
                this.Width = (int)this.Width;
            }
            else
            {
                this.Width = 932;
            }
        }

        /// <summary>
        ///Закрытие окна программы
        /// </summary>
        private void buCancel_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите выйти?", "Подтвердите выход", MessageBoxButton.OKCancel, MessageBoxImage.Question).ToString() == "OK")
            {
                //Отмена процесса выбора грани
                _tokenSourceCylinder.Cancel();
                _tokenSourceFace.Cancel();
                //Закрытие окна
                this.Close();
            }
            else
                return;
        }

        /// <summary>
        /// Отмена выбора цилиндрической грани
        /// </summary>
        private void unselectCylinderFace(object sender, RoutedEventArgs e)
        {
            swFaceObject = selectedEntity = selectedComponent = null;
            switch ((sender as ToggleButton).Name)
            {
                case "selectCylindrical_Worm":
                    border_cylinderWorm.Background = Brushes.White;
                    if (selectCylindrical_Gear.IsChecked != true)
                    {
                        selectFace_Worm.IsEnabled = true;
                        if (Gear_combo.Text == "Построить")
                            selectFace_Gear.IsEnabled = true;
                    }
                    break;
                case "selectCylindrical_Gear":
                    border_cylinderGear.Background = Brushes.White;
                    if (selectCylindrical_Worm.IsChecked != true)
                    {
                        selectFace_Gear.IsEnabled = true;
                        if (Worm_combo.Text == "Построить")
                            selectFace_Worm.IsEnabled = true;
                    }
                    break;
            }
            //Очистка выбора
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.ClearSelection2(true);
            //Отмена процесса выбора грани
            _tokenSourceCylinder.Cancel();
        }

        /// <summary>
        ///Выбор цилиндрической грани
        /// </summary>
        private async void selectCylindricalFace(object sender, RoutedEventArgs e)
        {
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;

            selectFace_Gear.IsChecked = selectFace_Worm.IsChecked = false;
            selectFace_Gear.IsEnabled = selectFace_Worm.IsEnabled = false;

            //Закрытие потока выбора плоской грани
            _tokenSourceFace.Cancel();
            //Получение токена для контроля выполнения
            _tokenSourceCylinder = new CancellationTokenSource();
            var token = _tokenSourceCylinder.Token;

            //Очистка предварительного выбора
            if ((sender as ToggleButton).Name == "selectCylindrical_Gear" & selectCylindrical_Worm.IsChecked != true)
                swModel.ClearSelection2(true);
            else if ((sender as ToggleButton).Name == "selectCylindrical_Worm" & selectCylindrical_Gear.IsChecked != true)
                swModel.ClearSelection2(true);

            //Запуск асинхронного процесса выбора грани
            await Task.Run(() => waitForCylinderSelection(token));

            //Визуальная обработка выбора грани
            if (selectedEntity == null || selectedComponent == null)
                selectCylindrical_Gear.IsChecked = selectCylindrical_Worm.IsChecked = false;
            else
            {
                var bc = new BrushConverter();
                var color_selected = bc.ConvertFromString("#FFA2FDAA");
                switch ((sender as ToggleButton).Name)
                {
                    case "selectCylindrical_Worm":
                        border_cylinderWorm.Background = (Brush)color_selected;
                        break;
                    case "selectCylindrical_Gear":
                        border_cylinderGear.Background = (Brush)color_selected;
                        break;
                }
            }
        }

        /// <summary>
        /// обработка выбора цилиндрической грани
        /// </summary>
        private void waitForCylinderSelection(CancellationToken token)
        {
            //Инициализация начальных параметров
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            Face2 swFace;
            SelectionMgr swSelMgr;
            Surface swSurf;
            Entity swEntity;
            Component2 swComponent;
            int FILTER = (int)swSelectType_e.swSelFACES;
            object swObject = null;

            //Проверяем, открыт ли документ
            if (swModel != null)
            {
                //Вызов менеджера выбора
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                //Процесс выполняется пока не выбрана цилиндрическая грань
                while (swObject == null)
                {
                    //Обработка выбранной грани
                    for (int i = 1; i <= swSelMgr.GetSelectedObjectCount2(-1); i++)
                    {
                        //Проверка, что была выбрана грань
                        if (swSelMgr.GetSelectedObjectType3(i, -1) == FILTER)
                        {
                            swObject = swSelMgr.GetSelectedObject6(i, -1);
                            swFace = (Face2)swObject;
                            swSurf = (Surface)swFace.GetSurface();

                            //Обработка выбора нецилиндрической грани
                            if (!swSurf.IsCylinder())
                            {
                                MessageBox.Show("Выберите цилиндрическую грань!");
                                var swEnt = (Entity)swObject;
                                swEnt.DeSelect();
                                swObject = null;
                                return;
                            }
                        }
                        else if (swSelMgr.GetSelectedObjectType3(i, -1) != FILTER)
                        {
                            MessageBox.Show("Выберите цилиндрическую грань!");
                            swModel.ClearSelection2(true);
                            return;
                        }
                    }
                    //Выполнение сторонних операций
                    DoEvents();
                    //Обработка отмены операции пользователем
                    if (token.IsCancellationRequested)
                        return;
                }
                //Объект грани
                swFaceObject = swObject;

                //Получение значения грани
                swFace = (Face2)swObject;

                //Получение компонента у выбранной грани
                swEntity = (Entity)swObject;
                swComponent = (Component2)swEntity.GetComponent();

                //Имя для вызова выбранной грани
                string entityName = getRandomString();

                //Применение имени грани
                swModel.SelectedFaceProperties(0, 0, 0, 0, 0, 0, 0, true, entityName);

                //Передача значения грани и модели в глобальные переменные
                selectedEntity = swModel.GetEntityName(swFace);
                selectedComponent = swComponent.GetSelectByIDString();
            }
        }

        //swFeature = (Feature) swFace.GetFeature();
        //ModelDocExtension swModelDocExt;
        // swModelDocExt = swModel.Extension;
        //if (faceArray.Length == 0)
        //    faceArray[0] = swFace;

        //if (faceArray.Length != 0)
        //{
        //    faceArray[1] = swFace;
        //    swModelDocExt.MultiSelect2(faceArray, true, null);
        //}

        /// <summary>
        ///Отмена выбора грани
        /// </summary>
        private void unselectFace(object sender, RoutedEventArgs e)
        {
            swFaceObject = selectedEntity = selectedComponent = null;
            switch ((sender as ToggleButton).Name)
            {
                case "selectFace_Worm":
                    border_faceWorm.Background = Brushes.White;
                    selectCylindrical_Worm.IsEnabled = true;
                    if (Gear_combo.Text == "Построить")
                    {
                        selectFace_Gear.IsEnabled = true;
                        selectCylindrical_Gear.IsEnabled = true;
                    }
                    break;
                case "selectFace_Gear":
                    border_faceGear.Background = Brushes.White;
                    selectCylindrical_Gear.IsEnabled = true;
                    if (Worm_combo.Text == "Построить")
                    {
                        selectFace_Worm.IsEnabled = true;
                        selectCylindrical_Worm.IsEnabled = true;
                    }
                    break;
            }

            //Очистка выбора
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.ClearSelection2(true);
            //Отмена процесса выбора грани
            _tokenSourceFace.Cancel();
        }

        /// <summary>
        ///Выбор грани
        /// </summary>
        private async void selectFace(object sender, RoutedEventArgs e)
        {
            selectCylindrical_Gear.IsChecked = selectCylindrical_Worm.IsChecked = false;
            selectCylindrical_Gear.IsEnabled = selectCylindrical_Worm.IsEnabled = false;

            switch (((ToggleButton)sender).Name)
            {
                case "selectFace_Worm":
                    selectFace_Gear.IsChecked = false;
                    selectFace_Gear.IsEnabled = false;
                    border_faceGear.Background = Brushes.White;
                    break;
                case "selectFace_Gear":
                    selectFace_Worm.IsChecked = false;
                    selectFace_Worm.IsEnabled = false;
                    border_faceWorm.Background = Brushes.White;
                    break;
            }
            //Завершения выполнения потока цилиндрической грани
            _tokenSourceCylinder.Cancel();
            //Получение токена для контроля выполнения
            _tokenSourceFace = new CancellationTokenSource();
            var token = _tokenSourceFace.Token;

            //Запуск асинхронного прцоесса выбора грани
            await Task.Run(() => waitForFaceSelection(token));

            //Визуальная обработка выбора грани
            if (selectedEntity == null || selectedComponent == null)
                selectFace_Gear.IsChecked = selectFace_Worm.IsChecked = false;
            else
            {
                var bc = new BrushConverter();
                var color_selected = bc.ConvertFromString("#FFA2FDAA");
                switch ((sender as ToggleButton).Name)
                {
                    case "selectFace_Worm":
                        border_faceWorm.Background = (Brush)color_selected;
                        break;
                    case "selectFace_Gear":
                        border_faceGear.Background = (Brush)color_selected;
                        break;
                }
            }
        }

        /// <summary>
        ///Обработка выбора грани
        /// </summary>
        private void waitForFaceSelection(CancellationToken token)
        {
            //Инициализация начальных параметров
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            Face2 swFace;
            SelectionMgr swSelMgr;
            Entity swEntity;
            Component2 swComponent;
            Surface swSurf;

            int FILTER = (int)swSelectType_e.swSelFACES;
            object swObject = null;
            double centerU, centerV;
            double[] Evals = new double[6];

            //Проверяем, открыт ли документ
            if (swModel != null)
            {
                //Отчистка выбора
                swModel.ClearSelection2(true);
                //Вызов менеджера выбора
                swSelMgr = (SelectionMgr)swModel.SelectionManager;
                //Процесс выполняется пока не выбрана грань
                while (swObject == null)
                {
                    //Обработка выбранной грани
                    for (int i = 1; i <= swSelMgr.GetSelectedObjectCount2(-1); i++)
                    {
                        //Проверка, что была выбрана грань
                        if (swSelMgr.GetSelectedObjectType3(i, -1) == FILTER)
                        {
                            //Обработка грани
                            swObject = swSelMgr.GetSelectedObject6(i, -1);
                            swFace = (Face2)swObject;
                            swSurf = (Surface)swFace.GetSurface();

                            //Обработка выбора нецилиндрической грани
                            if (swSurf.IsCylinder())
                            {
                                MessageBox.Show("Выберите плоскую грань!");
                                var swEnt = (Entity)swObject;
                                swEnt.DeSelect();
                                swObject = null;
                                return;
                            }
                            else
                            {
                                double[] vUvBounds = (double[])swFace.GetUVBounds();
                                centerU = ((double)vUvBounds[0] + (double)vUvBounds[1]) / 2;
                                centerV = ((double)vUvBounds[2] + (double)vUvBounds[3]) / 2;
                                //Вычисление центра грани
                                Evals = (double[])swSurf.Evaluate(centerU, centerV, 0, 0);
                            }
                        }
                        else if (swSelMgr.GetSelectedObjectType3(i, -1) != FILTER)
                        {
                            MessageBox.Show("Выберите плоскую грань!");
                            swModel.ClearSelection2(true);
                            return;
                        }
                    }
                    //Выполнение сторонних операций
                    DoEvents();
                    //Обработка отмены операции пользователем
                    if (token.IsCancellationRequested)
                        return;
                }
                MathTransform swTransform;
                MathPoint swmathPt;
                MathUtility swmathUtils;
                swmathUtils = (MathUtility)swApp.GetMathUtility();

                //Объект грани
                swFaceObject = swObject;
                //Получение значения грани
                swFace = (Face2)swObject;
                //Получение компонента у выбранной грани
                swEntity = (Entity)swObject;
                swComponent = (Component2)swEntity.GetComponent();

                //Начальная точка
                swTransform = swComponent.Transform2;
                double[] dOrigPt = new double[3];
                dOrigPt[0] = 0;
                dOrigPt[1] = 0;
                dOrigPt[2] = 0;
                swmathPt = (MathPoint)swmathUtils.CreatePoint(dOrigPt);
                swmathPt = (MathPoint)swmathPt.MultiplyTransform(swTransform);
                double[] vCompOriginPt = (double[])swmathPt.ArrayData;

                pointsOfOriginComponent = vCompOriginPt;
                pointsOfOriginComponent[0] += Evals[0];
                pointsOfOriginComponent[1] += Evals[1];
                pointsOfOriginComponent[2] += Evals[2];

                //Имя для вызова выбранной грани
                string entityName = getRandomString();
                //Применение имени грани
                swModel.SelectedFaceProperties(0, 0, 0, 0, 0, 0, 0, true, entityName);
                //Передача значения грани и модели в глобальные переменные
                selectedEntity = swModel.GetEntityName(swFace);
                selectedComponent = swComponent.GetSelectByIDString();
            }
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background,
                                                  new Action(delegate { }));
        }

        /// <summary>
        ///Вызов расчета параметров
        /// </summary>
        private void buCalc_Click(object sender, RoutedEventArgs e) { InitializeCalculate(); }

        /// <summary>
        ///Обработка построения червяка
        /// </summary>
        private void Chervyak_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (Worm_combo.SelectedIndex)
            {
                case 1:
                    selectCylindrical_Worm.IsEnabled = selectFace_Worm.IsEnabled = false;
                    break;
                case 0:
                    if (selectCylindrical_Gear.IsChecked == true || selectFace_Gear.IsChecked == true)
                    {
                        if (selectCylindrical_Gear.IsChecked == true)
                            selectCylindrical_Worm.IsEnabled = true;
                    }
                    else
                        selectCylindrical_Worm.IsEnabled = selectFace_Worm.IsEnabled = true;
                    break;
            }
        }

        /// <summary>
        /// Обработка построения колеса
        /// </summary>
        private void Gear_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (Gear_combo.SelectedIndex)
            {
                case 1:
                    selectCylindrical_Gear.IsEnabled = selectFace_Gear.IsEnabled = false;
                    break;
                case 0:
                    if (selectCylindrical_Worm.IsChecked == true || selectFace_Worm.IsChecked == true)
                    {
                        if (selectCylindrical_Worm.IsChecked == true)
                            selectCylindrical_Gear.IsEnabled = true;
                    }
                    else
                        selectCylindrical_Gear.IsEnabled = selectFace_Gear.IsEnabled = true;
                    break;
            }
        }

        /// <summary>
        ///Обработка выбора начального параметра
        /// </summary>
        private void radioButtonParam_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            string name = pressed.Content.ToString();
            if (name == "Передаточное отношение")
            {
                Teeth_gearValue.IsEnabled = false;
                Teeth_gearValue.Text = (float.Parse(Peredat_combo.Text, CultureInfo.InvariantCulture.NumberFormat)
                    * float.Parse(Kol_vitkovValue.Text, CultureInfo.InvariantCulture.NumberFormat)).ToString("0");
                Peredat_combo.IsEnabled = true;
                Peredat_combo.Visibility = Visibility.Visible;
                Peredat_Changed.Visibility = Visibility.Hidden;

            }
            else if (name == "Количество зубьев колеса")
            {
                Teeth_gearValue.IsEnabled = true;
                Peredat_Changed.Text = (float.Parse(Teeth_gearValue.Text, CultureInfo.InvariantCulture.NumberFormat)
                    / float.Parse(Kol_vitkovValue.Text, CultureInfo.InvariantCulture.NumberFormat)).ToString("0");
                Peredat_Changed.IsEnabled = false;
                Peredat_Changed.Visibility = Visibility.Visible;
                Peredat_combo.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Обработка выбора параметра для определения размера червяка
        /// </summary>
        private void radioButtonWorm_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            string name = pressed.Content.ToString();
            Module = float.Parse(((Module_combo.Text).Split(' ')[0]));
            if (name == "Коэффициент диаметра")
            {
                Koef_diamValue.IsEnabled = true;
                DegTeethValue.IsEnabled = false;
                Av_diamValue.IsEnabled = false;
                Av_diamValue.Text = (float.Parse(Koef_diamValue.Text) * Module).ToString(".00");
                DegTeethValue.Text = (Math.Atan(float.Parse(Kol_vitkovValue.Text)
                    / float.Parse(Koef_diamValue.Text)) * 180 / Math.PI).ToString(".00");

            }
            else if (name == "Угол наклона зуба")
            {
                Koef_diamValue.IsEnabled = false;
                DegTeethValue.IsEnabled = true;
                Av_diamValue.IsEnabled = false;
                Koef_diamValue.Text = (float.Parse(Kol_vitkovValue.Text)
                    / (Math.Tan(float.Parse(DegTeethValue.Text) * Math.PI / 180))).ToString(".00");
                Av_diamValue.Text = (float.Parse(Koef_diamValue.Text) * Module).ToString(".00");
            }
            else if (name == "Средний диаметр")
            {
                Koef_diamValue.IsEnabled = false;
                DegTeethValue.IsEnabled = false;
                Av_diamValue.IsEnabled = true;
                Koef_diamValue.Text = (float.Parse(Av_diamValue.Text) / Module).ToString(".00");
                DegTeethValue.Text = (Math.Atan(float.Parse(Kol_vitkovValue.Text)
                    / float.Parse(Koef_diamValue.Text)) * 180 / Math.PI).ToString(".00");
            }
        }

        /// <summary>
        ///Генерация случайной строки
        /// </summary>
        private string getRandomString()
        {
            var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var Charsarr = new char[8];
            var random = new Random();

            for (int i = 0; i < Charsarr.Length; i++)
            {
                Charsarr[i] = characters[random.Next(characters.Length)];
            }

            var resultString = new String(Charsarr);

            return resultString;
        }

        /// <summary>
        ///Сохранение данных программного модуля в json формат
        /// </summary>
        private void buSave_json_Click(object sender, RoutedEventArgs e)
        {
            //Сохранение данных для экспорта в json
            var saveParameters = getData(false);

            //Приведение данных в json формат
            string json = JsonConvert.SerializeObject(saveParameters, Newtonsoft.Json.Formatting.Indented);

            //Вызов и инициализация начальных параметров диалогового окна для выбора места сохранения
            System.Windows.Forms.SaveFileDialog SFD = new System.Windows.Forms.SaveFileDialog();
            SFD.FileName = "data";
            SFD.Filter = "Json files (*.json)|*.json";
            SFD.FilterIndex = 2;
            SFD.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            //Обработка нажатия на кнопку ОК в диалоговом окне
            if (SFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                File.WriteAllText(SFD.FileName, json);
                FileInfo file = new FileInfo(SFD.FileName);
                file.Attributes = FileAttributes.ReadOnly;
                MessageBox.Show("Файл успешно сохранен!");
            }
            else
                return;
        }

        /// <summary>
        ///Октрытие файла с параметрами 
        /// </summary>
        private void buOpen_json_Click(object sender, RoutedEventArgs e)
        {
            string path = "";
            string jsonRead = "";

            System.Windows.Forms.OpenFileDialog OFD = new System.Windows.Forms.OpenFileDialog();
            OFD.Filter = "Json files (*.json)|*.json";
            OFD.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (OFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                path = OFD.FileName;
            else
                return;

            using (StreamReader r = new StreamReader(path))
            {
                try { jsonRead = r.ReadToEnd(); }
                catch { MessageBox.Show("Ошибка в процессе чтения файла! Проверьте формат выбранного файла"); }
            };

            try
            {
                var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonRead);

                if (data["initial parameter"] == "Передаточное отношение")
                {
                    radioParam.IsChecked = true;
                    Peredat_combo.SelectedItem = data["gear_ratio"];
                }
                else if (data["initial parameter"] == "Количество зубьев")
                    radioParam1.IsChecked = true;

                if (data["worm_size"] == "Коэффициент диаметра")
                    radioWorm.IsChecked = true;
                else if (data["worm_size"] == "Угол наклона зуба")
                    radioWorm1.IsChecked = true;
                else if (data["worm_size"] == "Средний диаметр")
                    radioWorm2.IsChecked = true;

                Module_combo.SelectedItem = data["module"];
                ProfileDeg_combo.SelectedItem = data["profile_angle"];

                MaterialHelper MatObj = new MaterialHelper(swApp);
                var lister = MatObj.GetMaterials();
                lister.Insert(0, new Material
                {
                    Name = "Нет материала",
                    Classification = "Не задано",
                    DatabaseFileFound = false,
                    Elastic_modulus = "206000000000",
                    Poisson_ratio = "0.3",
                    Tensile_strength = "250000000",
                    Yield_strength = "160000000"
                });

                for (int i = 0; i < lister.Count; i++)
                {
                    if (lister[i].DisplayName == data["material_worm"])
                        MatWorm_combo.SelectedIndex = i;
                    if (lister[i].DisplayName == data["material_gear"])
                        MatGear_combo.SelectedIndex = i;
                }

                DegTeethValue.Text = data["tooth_angle"];
                if (data["tooth_direction"] == "Левое")
                    radioLeft.IsChecked = true;
                else if (data["tooth_direction"] == "Правое")
                    radioRight.IsChecked = true;

                if (float.Parse(data["hole_diameter"].ToString()) != 0)
                {
                    Hole_bool.IsChecked = true;
                    Hole_widthValue.Text = data["hole_diameter"];
                }

                LengthValue.Text = data["worm_length"];
                Kol_vitkovValue.Text = data["number_of_turns"];
                Kol_oborotovValue.Text = data["number_of_turnsVit"];
                Koef_diamValue.Text = data["diameter_factor"];
                Av_diamValue.Text = data["average_diameter"];
                Teeth_gearValue.Text = data["number_of_teeth"];
                Width_gearValue.Text = data["gear_width"];
                Koef_smeshValue.Text = data["bias_factor"];
                PowerValue.Text = data["power_worm"];
                VelocityValue.Text = data["speed_worm"];
                MomentValue.Text = data["torque_worm"];
                PowerValue_Gear.Text = data["power_gear"];
                VelocityValue_Gear.Text = data["speed_gear"];
                MomentValue_Gear.Text = data["torque_gear"];
                sigmaV.Text = data["tensile_strength"];
                sigmaT.Text = data["contact_strength"];
                E_gearValue.Text = data["elastic_modulus gear"];
                E_wormValue.Text = data["elastic_modulus worm"];
                Puasson_gearValue.Text = data["Poisson_ratio gear"];
                Puasson_wormValue.Text = data["Poisson_ratio worm"];
                KvValue.Text = data["koef_velocity"];
                timeValue.Text = data["time_of_work"];

                MessageBox.Show("Успешный импорт данных!", "Информация");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка считывания данных! " + ex.Message);
                return;
            }

        }

        private Dictionary<string, string> getData(bool isForTable)
        {
            var itemsModel = new List<string>();
            var itemsCalc = new List<string>();

            var tableGeneralModel = TableGeneral_Model.Items.OfType<Parameter_values>();
            var tableWormModel = TableWorm_Model.Items.OfType<Parameter_values>();
            var tableGearModel = TableGear_Model.Items.OfType<Parameter_values>();
            var tablesModel = tableGeneralModel.Concat(tableWormModel.Concat(tableGearModel));

            //Данные из таблиц вкладки Модель
            foreach (var item in tablesModel)
            {
                var value = item.value;
                itemsModel.Add(value);
            }
            var tableGeneralCalc = TableGeneral_Calc.Items.OfType<Parameter_values>();
            var tableWormCalc = TableWorm_Calc.Items.OfType<Parameter_values>();
            var tableGearCalc = TableGear_Calc.Items.OfType<Parameter_values>();
            var tablesCalc = tableGeneralCalc.Concat(tableWormCalc.Concat(tableGearCalc));

            //Данные из таблиц вкладки Расчет
            foreach (var item in tablesCalc)
            {
                var value = item.value;
                itemsCalc.Add(value);
            }

            string initialparameter = "";
            string wormSize = "";
            string teeth_direction = "";
            string peredatExp = "";
            Dictionary<string, string> data;


            //Значение исходного параметра
            if (radioParam.IsChecked == true)
            {
                initialparameter = radioParam.Content.ToString();
                peredatExp = Peredat_combo.Text;
            }
            else if (radioParam1.IsChecked == true)
            {
                initialparameter = radioParam1.Content.ToString();
                peredatExp = Peredat.ToString();
            }

            //значение параметра для расчета размера червяка
            if (radioWorm.IsChecked == true)
                wormSize = radioWorm.Content.ToString();
            else if (radioWorm1.IsChecked == true)
                wormSize = radioWorm1.Content.ToString();
            else if (radioWorm2.IsChecked == true)
                wormSize = radioWorm2.Content.ToString();

            //направление зубьев
            if (radioLeft.IsChecked == true)
                teeth_direction = "Левое";
            else if (radioRight.IsChecked == true)
                teeth_direction = "Правое";

            if (isForTable == false)
            {
                data = new Dictionary<string, string>
                {{ "initial parameter", initialparameter },
                {"worm_size", wormSize },
                {"gear_ratio", peredatExp },
                { "module", Module_combo.Text },
                {"profile_angle", ProfileDeg_combo.Text},
                {"tooth_angle", DegTeethValue.Text},
                {"number_of_turns", Kol_vitkovValue.Text},
                {"number_of_turnsVit", Kol_oborotovValue.Text},
                {"worm_length", LengthValue.Text},
                {"diameter_factor", Koef_diamValue.Text},
                {"average_diameter", Av_diamValue.Text},
                {"number_of_teeth", Teeth_gearValue.Text},
                {"gear_width", Width_gearValue.Text},
                {"bias_factor", Koef_smeshValue.Text},
                {"tooth_direction", teeth_direction},
                {"hole_diameter", Hole_widthValue.Text},
                {"power_worm", PowerValue.Text},
                {"speed_worm", VelocityValue.Text},
                {"torque_worm", MomentValue.Text},
                {"power_gear", PowerValue_Gear.Text},
                {"speed_gear", VelocityValue_Gear.Text},
                {"torque_gear", MomentValue_Gear.Text},
                {"material_worm", MatWorm_combo.Text},
                {"material_gear", MatGear_combo.Text},
                {"tensile_strength", sigmaV.Text},
                {"contact_strength", sigmaT.Text},
                {"elastic_modulus gear", E_gearValue.Text},
                {"elastic_modulus worm", E_wormValue.Text},
                {"Poisson_ratio gear", Puasson_gearValue.Text},
                {"Poisson_ratio worm", Puasson_wormValue.Text},
                {"koef_velocity", KvValue.Text},
                {"time_of_work", timeValue.Text} };

                return data;
            }
            else
            {
                data = new Dictionary<string, string>
                {{"Передаточное отношение", peredatExp},
                {"Модуль", Module_combo.Text},
                {"Угол профиля", ProfileDeg_combo.Text},
                {"Угол наклона зуба", DegTeethValue.Text + " °"},
                {"Межосевое расстояние", itemsModel[0]},
                {"Осевой угол зацепления", itemsModel[2].Replace("град", "") + "°"},
                {"Количество витков", Kol_vitkovValue.Text},
                {"Количество оборотов витков", Kol_oborotovValue.Text},
                {"Длина червяка", LengthValue.Text + " мм"},
                {"Коэффициент диаметра", Koef_diamValue.Text},
                {"Наружный диаметр червяка", itemsModel[3]},
                {"Средний диаметр червяка", itemsModel[4]},
                {"Диаметр впадин червяка", itemsModel[5]},
                {"Высота витка", itemsModel[6]},
                {"Высота головки витка", itemsModel[7]},
                {"Количество зубьев колеса", Teeth_gearValue.Text},
                {"Ширина червячного колеса", Width_gearValue.Text + " мм"},
                {"Коэффициент смещения", Koef_smeshValue.Text},
                {"Направление зубьев", teeth_direction},
                {"Диаметр отверстия колеса", Hole_widthValue.Text + " мм"},
                {"Наибольший диаметр колеса", itemsModel[8]},
                {"Наружный диаметр колеса", itemsModel[9]},
                {"Средний диаметр колеса", itemsModel[10]},
                { "Диаметр впадин колеса", itemsModel[11]},
                {"Мощность на червяке", PowerValue.Text + " кВт"},
                {"Скорость на червяке", VelocityValue.Text + " об/мин"},
                {"Крутящий момент на червяке", MomentValue.Text + " Нм"},
                {"Мощность на колесе", PowerValue_Gear.Text + " кВт"},
                {"Скорость на колесе", VelocityValue_Gear.Text + " об/мин"},
                {"Крутящий момент на колесе", MomentValue_Gear.Text + " Нм"},
                {"КПД", KPDValue.Text},
                {"Материал на червяке", MatWorm_combo.Text},
                {"Материал на колесе", MatGear_combo.Text},
                {"Предел устал. прочности изгиба (Sn), МПа", sigmaV.Text},
                {"Контактная усталостная прочность (Kw), МПа", sigmaT.Text },
                {"Модуль упругости червяка (E), МПа", E_wormValue.Text},
                {"Модуль упругости колеса (E), МПа", E_gearValue.Text },
                {"Коэффициент Пуассона, червяк (μ)", Puasson_wormValue.Text},
                {"Коэффициент Пуассона, колесо (μ)", Puasson_gearValue.Text },
                { "Коэффцииент скорости", KvValue.Text},
                { "Радиальная сила (Fr)", itemsCalc[0]},
                { "Нормальная сила (Fn)", itemsCalc[1]},
                { "Скорость скольжения (vk)", itemsCalc[2]},
                { "Коэффициент долговечности:", itemsCalc[3]},
                { "Окружная сила червяка (Ft)", itemsCalc[4]},
                { "Осевая сила червяка (Fa)", itemsCalc[5]},
                { "Окружная сила колеса (Ft)", itemsCalc[6]},
                { "Осевая сила колеса (Fa)", itemsCalc[7]},
                { "Контактное напряжение", itemsCalc[8]},
                { "Напряжение изгиба (σf)", itemsCalc[9]
                    } };

                return data;
            }

        }

        /// <summary>
        ///Экспорт данных в таблицу PDF 
        /// </summary>
        private void buSave_PDF_Click(object sender, RoutedEventArgs e)
        {
            //Вызов и инициализация начальных параметров диалогового окна для выбора места сохранения
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "PDF (*.pdf)|*.pdf";
            sfd.FileName = "Worm Gear Generator.pdf"; ;

            //Обработка нажатия в диалогвоом окне кнопки "ОК"
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    //Формирование таблицы PDF для экспорта, инициализация начальных параметров
                    PdfPTable pdfTable = new PdfPTable(2);
                    pdfTable.DefaultCell.Padding = 5;
                    pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfTable.WidthPercentage = 100;
                    string ttf = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts), "ARIAL.TTF");
                    var baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    var font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                    //Чтение необходимых данных
                    var data = getData(true);

                    //Запись данных в ячейки сформированной таблицы PDF
                    foreach (KeyValuePair<string, string> entry in data)
                    {
                        PdfPCell parameter = new PdfPCell(new Phrase(entry.Key, font));
                        parameter.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfTable.AddCell(parameter);
                        PdfPCell value = new PdfPCell(new Phrase(entry.Value, font));
                        value.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfTable.AddCell(value);
                    }

                    //Создание файла PDF, запись таблицы в файл
                    using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                    {
                        iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 20f, 20f, 10f);
                        PdfWriter.GetInstance(pdfDoc, stream);
                        pdfDoc.Open();
                        pdfDoc.Add(pdfTable);
                        pdfDoc.Close();
                        stream.Close();
                    }

                    MessageBox.Show("Файл успешно сохранен!", "Информация");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка :" + ex.Message);
                }

            }
        }

        /// <summary>
        ///Экспорт данных в таблицу docx 
        /// </summary>
        private void buSave_docx_Click(object sender, RoutedEventArgs e)
        {
            //Вызов и инициализация начальных параметров диалогового окна для выбора места сохранения
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "DOCX (*.docx)|*.docx";
            sfd.FileName = "Worm Gear Generator.docx";

            //Обработка нажатия в диалогвоом окне кнопки "ОК"
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    //Чтение данных
                    var data = getData(true);

                    //Создание таблицы docx
                    var doc = DocX.Create(sfd.FileName);
                    Xceed.Document.NET.Table t = doc.AddTable(data.Count, 2);
                    t.Alignment = Alignment.center;

                    //Запись данных в ячейки сформированной таблицы docx
                    var i = 0;
                    foreach (KeyValuePair<string, string> entry in data)
                    {
                        t.Rows[i].Cells[0].Paragraphs.First().Append(entry.Key);
                        t.Rows[i].Cells[1].Paragraphs.First().Append(entry.Value);
                        i++;
                    }
                    //Добавление таблицы в документ
                    doc.InsertTable(t);
                    //Сохранение файла
                    doc.Save();

                    MessageBox.Show("Файл успешно сохранен!", "Информация");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка :" + ex.Message);
                }
            }
        }

        /// <summary>
        ///Экспорт данных в таблицу xlsx
        /// </summary>
        private void buSave_excel_Click(object sender, RoutedEventArgs e)
        {
            //Вызов и инициализация начальных параметров диалогового окна для выбора места сохранения
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "XLSX (*.xlsx)|*.xlsx";
            sfd.FileName = "Worm Gear Generator.xlsx";

            //Обработка нажатия в диалогвоом окне кнопки "ОК"
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    using (var fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                    {
                        //Чтение данных
                        var data = getData(true);

                        //Создание таблицы xlsx
                        IWorkbook workbook = new XSSFWorkbook();
                        NPOI.SS.UserModel.ISheet excelSheet = workbook.CreateSheet("Sheet1");

                        //Запись данных в ячейки сформированной таблицы excel
                        var i = 0;
                        foreach (KeyValuePair<string, string> entry in data)
                        {
                            IRow row = excelSheet.CreateRow(i);
                            row.CreateCell(0).SetCellValue(entry.Key);
                            row.CreateCell(1).SetCellValue(entry.Value);
                            i++;
                        }
                        //Установка автоматической ширины
                        for (int x = 0; x < 2; x++) { excelSheet.AutoSizeColumn(x); }

                        //Добавление таблицы в файл
                        workbook.Write(fs);
                    }

                    MessageBox.Show("Файл успешно сохранен!", "Информация");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка :" + ex.Message);
                }
            }
        }

        /// <summary>
        ///Экспорт данных в txt
        /// </summary>
        private void buSave_TXT_Click(object sender, RoutedEventArgs e)
        {
            //Вызов и инициализация начальных параметров диалогового окна для выбора места сохранения
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "TXT (*.txt)|*.txt";
            sfd.FileName = "Worm Gear Generator.txt";

            //Обработка нажатия в диалогвоом окне кнопки "ОК"
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    using (FileStream fs1 = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                    {
                        //Чтение данных
                        var data = getData(true);
                        //Инициализация объекта класса StreamWrite для записи символов в файл
                        StreamWriter writer = new StreamWriter(fs1);

                        //Запись данных в файл
                        foreach (KeyValuePair<string, string> entry in data)
                        {
                            writer.WriteLine(entry.Key + ": " + entry.Value);
                        }
                        writer.Close();
                    }
                    MessageBox.Show("Файл успешно сохранен!", "Информация");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка :" + ex.Message);
                }
            }
        }
    };

    public struct Parameter_values
    {
        public string parameter { get; set; }
        public string value { get; set; }
    }
}
