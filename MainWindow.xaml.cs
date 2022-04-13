using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Media;


namespace WormGearGenerator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        float Module;
        float PressureAngle;
        float Peredat;
        float aw, Alpha, dae2, da1, da2, d1, d2, df1, df2, dw1, dw2;
        float Fn, Fr, Ft1, Ft2, Fa1, Fa2, vk, bending_stress, contact_stress, Khl, temperature;
        public SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        string AssemblyPath;
        private string baseDirectory = System.Environment.CurrentDirectory;

        public static string _pathAssembly { get; set; }
        public static string _pathWorm { get; set; }
        public static string _pathGear { get; set; }

        public static string _nameAssembly { get; set; }
        public static string _nameWorm { get; set; }
        public static string _nameGear { get; set; }

        public static bool buildAccepted { get; set; }

        public MainWindow()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

            InitializeFolder();
            InitializeComponent();
            InitializeStartData();
            InitializeCalculate();
        }

        private void InitializeFolder()
        {
            System.Windows.Forms.FolderBrowserDialog target = new System.Windows.Forms.FolderBrowserDialog();
            target.Description = "Выберите папку для сохранения сборки";
            if (target.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            else
                AssemblyPath = target.SelectedPath;
        }

        private void InitializeStartData()
        {

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
            Chervyak_combo.Items.Add("Построить модель");
            Chervyak_combo.Items.Add("Без построения");
            Chervyak_combo.SelectedIndex = 0;
            //Добавление в комбо построения колеса
            Gear_combo.Items.Add("Построить модель");
            Gear_combo.Items.Add("Без построения");
            Gear_combo.SelectedIndex = 0;

            //Значения червяка
            Kol_vitkovValue.Text = "4";
            LengthValue.Text = "140";
            Koef_diamValue.Text = "10";

            //Значения колеса
            Teeth_gearValue.Text = "60";
            Width_gearValue.Text = "20";
            Koef_smeshValue.Text = "1";
            Hole_widthValue.Text = "0";

            Hole_widthValue.Visibility = Visibility.Hidden;

            //Значения для силового расчета
            PowerValue.Text = "0.1";
            VelocityValue.Text = "1000.0";

            //Заполнение материалов
            MaterialHelper MatObj = new MaterialHelper(swApp);
            var lister = MatObj.GetMaterials();
            MatGear_combo.ItemsSource = MatWorm_combo.ItemsSource = lister;
            MatWorm_combo.DisplayMemberPath = MatGear_combo.DisplayMemberPath = "DisplayName";

            //Значения силовых характеристик
            KPDValue.IsEnabled = false;
            sigmaV.Text = "250";
            sigmaT.Text = "160";
            E_wormValue.Text = "206000";
            E_gearValue.Text = "101000";
            Puasson_wormValue.Text = "0.300";
            Puasson_gearValue.Text = "0.310";
            
            KoValue.Text = "1.200";
            KvValue.Text = "1.042";
            yValue.Text = "0.125";
            timeValue.Text = "10000";

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

        private void InitializeCalculate()
        {
            //Значения внутри полей
            Module = float.Parse(((Module_combo.Text).Split(' ')[0]), CultureInfo.InvariantCulture.NumberFormat);

            //Угол профиля
            PressureAngle = float.Parse(((ProfileDeg_combo.Text).Split(' ')[0]), CultureInfo.InvariantCulture.NumberFormat);

            //Количество оборотов витков
            Kol_oborotovValue.Text = (Math.Floor(float.Parse(LengthValue.Text) / (Module * Math.PI)) + 3).ToString();
           

            //Расчет значений передаточного числа и зубьев колеса
            if (radioParam.IsChecked == true)
            {
                Peredat = float.Parse(Peredat_combo.Text, CultureInfo.InvariantCulture.NumberFormat);
                Teeth_gearValue.Text = (float.Parse(Peredat_combo.Text, CultureInfo.InvariantCulture.NumberFormat)
                * float.Parse(Kol_vitkovValue.Text, CultureInfo.InvariantCulture.NumberFormat)).ToString("0");
            }
            else
            {
                Peredat_Changed.Text = (float.Parse(Teeth_gearValue.Text, CultureInfo.InvariantCulture.NumberFormat)
                    / float.Parse(Kol_vitkovValue.Text, CultureInfo.InvariantCulture.NumberFormat)).ToString("0.00");
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

            //Межосевое расстояние (aw)
            aw = Module * (float.Parse(Teeth_gearValue.Text) + float.Parse(Koef_diamValue.Text) + 2*float.Parse(Koef_smeshValue.Text)) / 2;

            //КПД
            float v1 = (float)((Math.PI * dw1 * float.Parse(VelocityValue.Text)) / 60000);
            //Делительный угол подъема
            float tgy = (float)(Math.Atan((float.Parse(Kol_vitkovValue.Text) / float.Parse(Koef_diamValue.Text))) * 180 / Math.PI);
            //Начальный угол подъема
            float tgyW = (float)(Math.Atan(float.Parse(Kol_vitkovValue.Text) /( float.Parse(Koef_diamValue.Text)+2*float.Parse(Koef_smeshValue.Text))) * 180 / Math.PI);

            float v2 = v1 * (float.Parse(Kol_vitkovValue.Text) / float.Parse(Koef_diamValue.Text));

            //Скорость скольжения
            vk = (float)(v1 / Math.Cos(tgyW * Math.PI / 180));

            float phiz = (float)(Math.Atan(0.02 + (0.03) / (vk)) * 180 / Math.PI);
            float Kv = (1200 + v2) / (1200);
            float KPD = 0;
            if (boolKPD.IsChecked == true)
                KPD = float.Parse(KPDValue.Text);
            else
                KPD = (float)(((Math.Tan(tgyW * Math.PI / 180)) / (Math.Tan((tgyW + phiz) * Math.PI / 180))) * 0.96); 
                KPDValue.Text = KPD.ToString("0.000"); 
            
            //Крутящий момент
          //  MomentValue.Text = ((60000 * float.Parse(PowerValue.Text)) / (2 * Math.PI * float.Parse(VelocityValue.Text))).ToString("0.000");      
            //значения на червяке
            float Power = float.Parse(PowerValue.Text);
            float Velocity = float.Parse(VelocityValue.Text);
            float Moment = 0;

            //значения на колесе
            float Velocity_WG = 0;
            float Power_WG = 0;
            float Moment_WG = 0;

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
                Power = (float)((Math.PI * Velocity_WG * Moment_WG) / (30 * KPD))/1000;
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
                Power = Power_WG * KPD;
                Velocity = Velocity_WG * Peredat;
                Moment_WG = (float)((60000 * Power_WG) / (2 * Math.PI * Velocity_WG));
                Moment = (float)((60000 * Power) / (2 * Math.PI * Velocity));

            }
            else if (radioGearSel.IsChecked == true & RadioType_Calc1.IsChecked == true)
            {
                Velocity_WG = float.Parse(VelocityValue_Gear.Text);
                Moment_WG = float.Parse(MomentValue_Gear.Text);
                Power_WG = (float)((Math.PI*Velocity_WG*Moment_WG) / (30))/1000;
                Moment = (Moment_WG / Peredat) * KPD;
                Velocity = Velocity_WG * Peredat;
                Power = Power_WG * KPD;
            }
            else if (radioGearSel.IsChecked == true & RadioType_Calc2.IsChecked == true)
            {
                Power_WG = float.Parse(PowerValue_Gear.Text);
                Moment_WG = float.Parse(MomentValue_Gear.Text);
                Velocity_WG = (float)((60000*Power_WG) / (2*Math.PI*Moment_WG));
                Power = Power_WG * KPD;
                Velocity = Velocity_WG * Peredat;
                Moment = (Moment_WG / Peredat) * KPD;
            }
            VelocityValue.Text = Velocity.ToString("0.00");
            PowerValue.Text = Power.ToString("0.000");
            MomentValue.Text = Moment.ToString("0.000");
            VelocityValue_Gear.Text = Velocity_WG.ToString("0.00");
            PowerValue_Gear.Text = Power_WG.ToString("0.000");
            MomentValue_Gear.Text = Moment_WG.ToString("0.000");

            //окружная сила на червяке, Осевая на колесе
            Ft1 = (2000 * float.Parse(MomentValue.Text)) / (dw1);
            Fa2 = Ft1;

            //осевая сила на червяке, окружная на колесе
            Fa1 = (2000 * Moment_WG) / (dw2);
            Ft2 = Fa1;

            //Радиальная сила передачи
            Fr = (float)(Ft2 * Math.Tan(PressureAngle * (Math.PI / 180)));

            //Контактное напряжение
            float Koef = 0;
            if (vk >= 10)
                Koef = (float)1.2;
            else if (vk < 10 & vk > 5)
                Koef = (float)1.1;
            else if (vk <= 5)
                Koef = 1;


            //float Module_upr = (float.Parse(E_wormValue.Text) + float.Parse(E_gearValue.Text)) / 2;
            //contact_stress = (float)((1.31 / d2) * Math.Sqrt(((Moment_WG * 1000) * Koef * Module_upr ) / dw1));

            contact_stress = (float)((170 / (float.Parse(Teeth_gearValue.Text) / float.Parse(Koef_diamValue.Text))) * 
                Math.Sqrt((float.Parse(MomentValue_Gear.Text)*1000*float.Parse(KvValue.Text))* Math.Pow(((1 + float.Parse(Teeth_gearValue.Text) / float.Parse(Koef_diamValue.Text)) /(aw)), 3)));

            //Напряжение изгиба
            float zv2 = (float)((float.Parse(Teeth_gearValue.Text)) / (Math.Pow(Math.Cos(tgyW * Math.PI / 180), 3)));
            float Yf2 = 0;
            if (zv2 < 37)
                Yf2 = (float)(2.4 - 0.0214 * zv2);
            else if (zv2 >= 37 & zv2 <= 45)
                Yf2 = (float)(2.21 - 0.0162 * zv2);
            else if (zv2 > 45)
                Yf2 = (float)(1.72 - 0.0053 * zv2);

            bending_stress = (float)(((0.7 * Ft2 * Koef) / (float.Parse(Width_gearValue.Text) * Module * Math.Cos(tgyW * Math.PI/180) )) * Yf2);

            //Нормальная сила
            Fn = (float)(Ft2 / (Math.Cos(PressureAngle * (Math.PI / 180)) * Math.Cos(tgy * (Math.PI / 180))));

            //Число циклов нагружения
            float Nk = 60 * float.Parse(VelocityValue_Gear.Text) * float.Parse(timeValue.Text);
            //Коэффициент долговечности
            Khl = (float)(Math.Pow(((Math.Pow(10,7))/(Nk)), 0.125));

            //ожидаемое значение скорости
            float v_wait = (float)((4.5 / 10000) * float.Parse(VelocityValue.Text) * Math.Pow(float.Parse(MomentValue_Gear.Text), 1.0 / 3));

            //коэф изнашивания зубьев
            float Cv = (float)(1.46 - (v_wait / 7.29) * (1 - (v_wait / 20.2)));

            //Допускаемые контактные напряжения
            limit_contact.Text = (float.Parse(sigmaV.Text)*0.82 * Cv * Khl).ToString("0.00");

            //предел выносливости зубьев при изгибе
            float limit_bending_Koef = (float)(0.08 * float.Parse(sigmaV.Text) + 0.25 * float.Parse(sigmaT.Text));
            float Yhl = (float)(Math.Pow(((Math.Pow(10, 6)) / (Nk)), 1.0/9));
            limit_bending.Text = (limit_bending_Koef * Yhl).ToString("0.00");



            //Температура масла при работе
            temperature = (float)((float.Parse(PowerValue.Text)*1000 * (1 - float.Parse(KPDValue.Text))) / (17 * 12.2 * Math.Pow(aw / 1000, 1.17))) + 20; 

            RefreshTable_Model(aw, Module, Alpha, da1, d1, df1, da2, d2, df2, dae2);
            RefreshTable_Calc(Fr, vk, Khl, temperature, Ft1, Fa1, Ft2, Fa2, Fn, contact_stress, bending_stress);

            var bc = new BrushConverter();
            var color = Brushes.White;

            TableGeneral_Model.RowBackground = color;
            TableWorm_Model.RowBackground = color;
            TableGear_Model.RowBackground = color;
            TableGeneral_Calc.RowBackground = color;
            TableWorm_Calc.RowBackground = color;
            TableGear_Calc.RowBackground = color;
        }

        private void RefreshTable_Model(float aw_Table, float Module_Table, float Alpha_Table,
            float DaWorm_Table, float DWorm_Table, float DfWorm_Table, float DaGear_Table, float DGear_Table, float DfGear_Table, float DaeGear_Table)
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

        private void RefreshTable_Calc(float Fr_Table, float Vk_Table, float Khl_Table, double temp_Table, float Ft1_Table,
            float Fa1_Table, float Ft2_Table, float Fa2_Table, float Fn_Table, float contactPres_Table, float bending_Table)
        {

            ObservableCollection<Parameter_values> general = new ObservableCollection<Parameter_values>();
            general.Add(new Parameter_values { parameter = "Радиальная сила (Fr):", value = Fr_Table.ToString("0.00") + " Н" });
            general.Add(new Parameter_values { parameter = "Нормальная сила (Fn):", value = Fn_Table.ToString("0.00") + " Н" });
            general.Add(new Parameter_values { parameter = "Скорость скольжения (Vk):", value = Vk_Table.ToString("0.00") + " м/c" });
            general.Add(new Parameter_values { parameter = "Коэф. долговечности:", value = Khl_Table.ToString("0.00") });
            general.Add(new Parameter_values { parameter = "t° масла при работе:", value = temp_Table.ToString("0.00") + "°" });
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

        private void data_TextChanged(object sender, TextChangedEventArgs e) { changeBack_Touched(); }
        private void combo_SelectionChanged(object sender, SelectionChangedEventArgs e) { changeBack_Touched(); }

        private void changeBack_Touched()
        {
            var bc = new BrushConverter();
            var color = Brushes.LightGray;

            TableGeneral_Model.RowBackground = color;
            TableWorm_Model.RowBackground = color;
            TableGear_Model.RowBackground = color;
            TableGeneral_Calc.RowBackground = color;
            TableWorm_Calc.RowBackground = color;
            TableGear_Calc.RowBackground = color;
        }

        private void MatWorm_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Material item = (Material)MatWorm_combo.SelectedItem;

            if (item.Elastic_modulus == null)
            {
                E_wormValue.Text = "0";

            }
            else
                E_wormValue.Text = (float.Parse(item.Elastic_modulus) / 1000000).ToString("0.00");

            if (item.Elastic_modulus == null)
            {
                Puasson_wormValue.Text = "0";

            }
            else
                Puasson_wormValue.Text = item.Poisson_ratio;

        }

        private void MatGear_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Material item = (Material)MatGear_combo.SelectedItem;

            if (item.Tensile_strength == null)
            { sigmaV.Text = "0";

            }
            else
                sigmaV.Text = (float.Parse(item.Tensile_strength) / 1000000).ToString("0.00");

            if (item.Yield_strength == null)
            {
                sigmaT.Text = "0";
                
            }
            else
                sigmaT.Text = (float.Parse(item.Yield_strength) / 1000000).ToString("0.00");

            if (item.Elastic_modulus == null)
            {
                E_gearValue.Text = "0";
 
            }
            else
                E_gearValue.Text = (float.Parse(item.Elastic_modulus) / 1000000).ToString("0.00");

            if (item.Elastic_modulus == null)
            {
                Puasson_gearValue.Text = "0";

            }
            else
                Puasson_gearValue.Text = item.Poisson_ratio;
        }

        private void calc_WidthGear_Click(object sender, RoutedEventArgs e)
        {
            var teeth = float.Parse(Kol_vitkovValue.Text);
            if (teeth <= 3)
                Width_gearValue.Text = (0.75 * da1).ToString("0.00");
            else if (teeth == 4)
                Width_gearValue.Text = (0.67 * da1).ToString("0.00");
        }

        private void calc_LengthWorm_Click(object sender, RoutedEventArgs e)
        {
            LengthValue.Text = (2 * Math.Sqrt(Math.Pow(dae2/2,2)- Math.Pow((aw - da1/2),2) + (Math.PI*Module/2))).ToString("0.00");
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var bc = new BrushConverter();
            var color = bc.ConvertFromString("#FFCBE4FF");

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
                Expander_showCalc.Background = (Brush)color;
                Expander_showCalc1.Background = (Brush)color;
            }
            else if (this.Width == 704)
            {
                Expander_showCalc.Background = Brushes.White;
                Expander_showCalc1.Background = Brushes.White;
            }

        }

        private void confirm_BuildClick(object sender, RoutedEventArgs e)
        {
            bool isWorm = Chervyak_combo.Text == "Построить модель" ? true : false;             
            bool isGear = Gear_combo.Text == "Построить модель" ? true : false;             

            this.Hide();
            if (isWorm == false & isGear == false)
            {
                return;
            }
            else
            {
                ConfirmFolders foldersWindow = new ConfirmFolders(this, AssemblyPath, isWorm, isGear);
                foldersWindow.Topmost = true;
                foldersWindow.ShowDialog();

                if (buildAccepted == true)
                    create_components(isWorm, isGear);
                else
                {
                    this.Show();
                    return;
                }
            }
            
        }

        private void create_components(bool isWorm, bool isGear )
        {
            try
            {
                Worm worm = new Worm(baseDirectory);
                Gear gear = new Gear(baseDirectory);
                Material materialWorm = null;
                Material materialGear = null;

                int rightOrLeft;
                if (radioLeft.IsChecked == true)
                    rightOrLeft = 0;
                else
                    rightOrLeft = 1;

                float hole_diameter = 0;
                if (Hole_bool.IsChecked == true)
                    hole_diameter = float.Parse(Hole_widthValue.Text);
                

                if(MatWorm_combo.SelectedItem != null)
                    materialWorm = (Material)MatWorm_combo.SelectedItem;

                if(MatGear_combo.SelectedItem != null)
                    materialGear = (Material)MatGear_combo.SelectedItem;

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
                if (isWorm & isGear)
                {
                    SolidWorker sldobj = new SolidWorker();

                    //создаем сборку
                    sldobj.CreateAssembly(_pathAssembly);

                    //добавляем компоненты
                    sldobj.AddComponent(worm, gear, _pathAssembly);

                    //добавление зависимостей
                    var teethGear = float.Parse(Teeth_gearValue.Text);
                    var vitWorm = float.Parse(Kol_vitkovValue.Text);
                    sldobj.AddMates(_pathAssembly, _nameAssembly, _nameWorm, _nameGear, aw, teethGear, vitWorm, rightOrLeft);
                }

                Directory.SetCurrentDirectory(baseDirectory);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка во время построения:" + ex.Message);
                this.Show();
            }
        }

        private void CheckBox_Hole_Changed(object sender, RoutedEventArgs e)
        {
            if (Hole_bool.IsChecked == true)
                Hole_widthValue.Visibility = Visibility.Visible;
            else
                Hole_widthValue.Visibility = Visibility.Hidden;
        }

        private void CheckBox_KPD_Changed(object sender, RoutedEventArgs e)
        {
            if (boolKPD.IsChecked == true)
                KPDValue.IsEnabled = true;
            else
                KPDValue.IsEnabled = false;
        }

        private void Expander_Expanded(object sender, RoutedEventArgs e) { this.Height = 660; }

        private void Expander_Collapsed(object sender, RoutedEventArgs e) { this.Height = 550; }

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

        private void buCancel_Click(object sender, RoutedEventArgs e) { this.Close(); }

        private void buCalc_Click(object sender, RoutedEventArgs e) {  InitializeCalculate();  }

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
        private void radioButtonWorm_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            string name = pressed.Content.ToString();
            Module = float.Parse(((Module_combo.Text).Split(' ')[0]), CultureInfo.InvariantCulture.NumberFormat);
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

        private void buSave_json_Click(object sender, RoutedEventArgs e)
        {
            var saveParameters = getData(false);

            string json = JsonConvert.SerializeObject(saveParameters, Newtonsoft.Json.Formatting.Indented);

            System.Windows.Forms.SaveFileDialog SFD = new System.Windows.Forms.SaveFileDialog();
            SFD.FileName = "data";
            SFD.Filter = "Json files (*.json)|*.json";
            SFD.FilterIndex = 2;
            SFD.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                if (SFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    File.WriteAllText(SFD.FileName, json);
                    MessageBox.Show("Файл успешно сохранен!");
                }
                else
                    return;
            }
            catch { MessageBox.Show("Ошибка в процессе сохранения!"); }


        }

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
                for (int i = 0; i < lister.Count; i++)
                {
                    if(lister[i].DisplayName == data["material_worm"])
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
                
                KoValue.Text = data["koef peregruz"];
                KvValue.Text = data["koef_velocity"];
                yValue.Text = data["koef_luis"];
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
                
                {"koef peregruz", KoValue.Text},
                {"koef_velocity", KvValue.Text},
                {"koef_luis", yValue.Text},
                {"time_of_work", timeValue.Text} };

                return data;
            }
            else
            {
                data = new Dictionary<string, string>
                {{"Передаточное отношение", peredatExp},
                {"Модуль", Module_combo.Text},
                {"Угол профиля", ProfileDeg_combo.Text},
                {"Угол наклона зуба, °", DegTeethValue.Text},
                { "Межосевое расстояние, мм", aw.ToString()},
                {"Угол профиля, °", Alpha.ToString()},
                {"Количество витков", Kol_vitkovValue.Text},
                {"Количество оборотов витков", Kol_oborotovValue.Text},
                {"Длина червяка, мм", LengthValue.Text},
                {"Коэффициент диаметра", Koef_diamValue.Text},
                {"Наружный диаметр червяка", da1.ToString()},
                {"Средний диаметр червяка", Av_diamValue.Text},
                {"Диаметр впадин червяка", df1.ToString()},
                {"Количество зубьев колеса", Teeth_gearValue.Text},
                {"Ширина червячного колеса, мм", Width_gearValue.Text},
                {"Коэффициент смещения", Koef_smeshValue.Text},
                {"Направление зубьев", teeth_direction},
                {"Диаметр отверстия колеса, мм", Hole_widthValue.Text},
                {"Наибольший диаметр колеса, мм", dae2.ToString()},
                {"Наружный диаметр колеса", da2.ToString()},
                {"Средний диаметр колеса", d2.ToString()},
                { "Диаметр впадин колеса", df2.ToString()},
                {"Мощность на червяке, кВт", PowerValue.Text},
                {"Скорость на червяке, об/мин", VelocityValue.Text},
                {"Крутящий момент на червяке, Нм", MomentValue.Text},
                {"Мощность на колесе, кВт", PowerValue_Gear.Text},
                {"Скорость на колесе, об/мин", VelocityValue_Gear.Text},
                {"Крутящий момент на колесе, Нм", MomentValue_Gear.Text},
                {"КПД", KPDValue.Text},
                {"Материал на червяке", MatWorm_combo.Text},
                {"Материал на колесе", MatGear_combo.Text},
                {"Предел устал. прочности изгиба (Sn), МПа", sigmaV.Text},
                {"Контактная усталостная прочность (Kw), МПа", sigmaT.Text },
                {"Модуль упругости червяка (E), МПа", E_wormValue.Text},
                {"Модуль упругости колеса (E), МПа", E_gearValue.Text },
                {"Коэффициент Пуассона, червяк (μ)", Puasson_wormValue.Text},
                {"Коэффициент Пуассона, колесо (μ)", Puasson_gearValue.Text },
                
                {"Коэффициент Льюиса (y)", yValue.Text },
                {"Коэффициент перегрузки", KoValue.Text},
                { "Коэффцииент скорости", KvValue.Text},
                { "Требуемый срок службы ", timeValue.Text},
                { "Радиальная сила (Fr), Н", Fr.ToString()},
                { "Нормальная сила (Fn), Н", Fn.ToString()},
                { "Контактное напряжение, МПа", contact_stress.ToString()},
                { "Скорость скольжения (vk), м/с", vk.ToString()},
                { "Коэффициент долговечности:", Khl.ToString()},
                { "Температура масла при работе, °:", temperature.ToString()},
                { "Окружная сила червяка (Ft), Н", Ft1.ToString()},
                { "Осевая сила червяка (Fa), Н", Fa1.ToString()},
                { "Окружная сила колеса (Ft), Н", Ft2.ToString()},
                { "Осевая сила колеса (Fa), Н", Fa2.ToString()},
                { "Напряжение изгиба (σf), МПа", bending_stress.ToString()} };

                return data;
            }

        }


        private void buSave_PDF_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "PDF (*.pdf)|*.pdf";
            sfd.FileName = "Worm Gear Generator.pdf";
            bool fileError = false;
            var data = getData(true);

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(sfd.FileName))
                {
                    try
                    {
                        File.Delete(sfd.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("Ошибка записи файла." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                        PdfPTable pdfTable = new PdfPTable(2);
                        pdfTable.DefaultCell.Padding = 5;
                        pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfTable.WidthPercentage = 100;
                        string ttf = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts), "ARIAL.TTF");
                        var baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        var font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);


                        foreach (KeyValuePair<string, string> entry in data)
                        {
                            PdfPCell parameter = new PdfPCell(new Phrase(entry.Key, font));
                            parameter.HorizontalAlignment = Element.ALIGN_CENTER;
                            pdfTable.AddCell(parameter);
                            PdfPCell value = new PdfPCell(new Phrase(entry.Value, font));
                            value.HorizontalAlignment = Element.ALIGN_CENTER;
                            pdfTable.AddCell(value);
                        }


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
        }

        private void buSave_docx_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "DOCX (*.docx)|*.docx";
            sfd.FileName = "Worm Gear Generator.docx";
            bool fileError = false;
            var data = getData(true);

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(sfd.FileName))
                {
                    try
                    {
                        File.Delete(sfd.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("Ошибка записи файла." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                            var doc = DocX.Create(sfd.FileName);
                            Xceed.Document.NET.Table t = doc.AddTable(data.Count, 2);
                            t.Alignment = Alignment.center;

                            var i = 0;
                            foreach (KeyValuePair<string, string> entry in data)
                            {
                                t.Rows[i].Cells[0].Paragraphs.First().Append(entry.Key);
                                t.Rows[i].Cells[1].Paragraphs.First().Append(entry.Value);
                                i++;
                            }
                            doc.InsertTable(t);
                            doc.Save();

                        MessageBox.Show("Файл успешно сохранен!", "Информация");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка :" + ex.Message);
                    }
                    
                }
            }
        }

        private void buSave_excel_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "XLSX (*.xlsx)|*.xlsx";
            sfd.FileName = "Worm Gear Generator.xlsx";
            bool fileError = false;
            var data = getData(true);

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(sfd.FileName))
                {
                    try
                    {
                        File.Delete(sfd.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("Ошибка записи файла." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                        using (var fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            IWorkbook workbook = new XSSFWorkbook();
                            NPOI.SS.UserModel.ISheet excelSheet = workbook.CreateSheet("Sheet1");
                            var i = 0;
                            foreach (KeyValuePair<string, string> entry in data)
                            {
                                IRow row = excelSheet.CreateRow(i);
                                row.CreateCell(0).SetCellValue(entry.Key);
                                row.CreateCell(1).SetCellValue(entry.Value);
                                i++;
                            }

                            for (int x = 0; x < 2; x++) { excelSheet.AutoSizeColumn(x); }

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
        }

        private void buSave_TXT_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "TXT (*.txt)|*.txt";
            sfd.FileName = "Worm Gear Generator.txt";
            bool fileError = false;
            var data = getData(true);

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(sfd.FileName))
                {
                    try
                    {
                        File.Delete(sfd.FileName);
                    }
                    catch (IOException ex)
                    {
                        fileError = true;
                        MessageBox.Show("Ошибка записи файла." + ex.Message);
                    }
                }
                if (!fileError)
                {
                    try
                    {
                        using (FileStream fs1 = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            StreamWriter writer = new StreamWriter(fs1);
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
        }


    };


    public class Parameter_values
    {
        public string parameter { get; set; }
        public string value { get; set; }
    }
}
