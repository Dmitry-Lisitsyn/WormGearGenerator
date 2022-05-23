using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WormGearGenerator.Helpers
{
    public class DataValidation : System.ComponentModel.IDataErrorInfo
    {
        //Червяк
        public float deg_TeethValue { get; set; }
        public string kol_vitkov { get; set; } = "4";
        public string length_Worm { get; set; } = "140";
        public string koef_diamWorm { get; set; } = "10";
        public string av_diamWorm { get; set; } = "40";
        //Червячное колесо
        public string teeth_Gear { get; set; } = "60";
        public string width_Gear { get; set; } = "20";
        public string koef_Smesh { get; set; } = "1";
        public string hole_Gear { get; set; } = "0";
        //Общие в силовом
        public string PowerWorm { get; set; } = "0.1"; 
        public string PowerGear { get; set; }
        public string VelocityWorm { get; set; } = "1000";
        public string VelocityGear { get; set; }
        public string MomentWorm { get; set; }
        public string MomentGear { get; set; }
        public string KPD { get; set; } // 0.1 - 1
        //Параметры материалов
        public string sigmaV { get; set; } = "250"; 
        public string sigmaT { get; set; } = "160"; 
        public string E_worm { get; set; } = "206000";
        public string E_gear { get; set; } = "101000";
        public string Puasson_worm { get; set; } = "0.3";
        public string Puasson_gear { get; set; } = "0.31";
        //Коэффициенты
        public string Ko { get; set; } = "1.200"; 
        public string Kv { get; set; } = "1.042"; 
        public string y { get; set; } = "0.125"; 
        public string time { get; set; } = "10000";
        //Проверка на прочность, изгиб и температуру
        public string temperature { get; set; }
        public string contact_calc { get; set; }
        public string contact { get; set; }
        public string bending_calc { get; set; }
        public string bending { get; set; }

        private MainWindow _window;

        private Dictionary<string, string> errors =  new Dictionary<string, string>();
        private Dictionary<string, string> errorCalc =  new Dictionary<string, string>();

        public string this[string columnName]
        {
            get
            {
                string error = String.Empty;
                switch (columnName)
                {
                    case "kol_vitkov":

                        if (int.TryParse(kol_vitkov, out int kol_vitkov_parsed))
                            validateError(errors, "kol_vitkov", "Значение витков должно быть в диапазоне от 1 до 4", () => (kol_vitkov_parsed > 0) & (kol_vitkov_parsed <= 4));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "length_Worm":

                        if (float.TryParse(length_Worm, out float length_Worm_parsed))
                            validateError(errors,"length_Worm", "Значение витков должно быть в диапазоне от 1 до 10000 мм", () => (length_Worm_parsed > 1) & (length_Worm_parsed < 10000));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "av_diamWorm":

                        if (float.TryParse(av_diamWorm, out float av_diamWorm_parsed))
                            validateError(errors, "av_diamWorm", "Значение параметра не должно быть меньше или равно нулю", () => (av_diamWorm_parsed > 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "koef_diamWorm":
                        if (float.TryParse(koef_diamWorm, out float koef_diamWorm_parsed))
                            validateError(errors,"koef_diamWorm", "Значение параметра не должно быть меньше или равно нулю", () => (koef_diamWorm_parsed > 0));
                        else
                            addError(errors,columnName, "Некорректный ввод данных");
                        break;

                    case "teeth_Gear":
                        if (int.TryParse(teeth_Gear, out int teeth_Gear_parsed))
                            validateError(errors,"teeth_Gear", "Значение параметра должно быть в диапазоне от 8 до 1999", () => (teeth_Gear_parsed >= 8) & (teeth_Gear_parsed < 2000));
                        else
                            addError(errors,columnName, "Некорректный ввод данных");
                        break;

                    case "width_Gear":
                        if (float.TryParse(width_Gear, out float width_Gear_parsed))
                            validateError(errors,"width_Gear", "Ширина червячного колеса не должна быть отрицательной или равняться нулю", () => (width_Gear_parsed > 0));
                        else
                            addError(errors,columnName, "Некорректный ввод данных");
                        break;

                    case "koef_Smesh":
                        if (float.TryParse(koef_Smesh, out float koef_Smesh_parsed))
                            validateError(errors,"koef_Smesh", "Значение параметра должно быть в диапазоне от -1 до 1", () => (koef_Smesh_parsed >= -1) & (koef_Smesh_parsed <= 1));
                        else
                            addError(errors,columnName, "Некорректный ввод данных");
                        break;

                    case "hole_Gear":
                        if (float.TryParse(hole_Gear, out float hole_Gear_parsed))
                            validateError(errors,  "hole_Gear", "Значение параметра не должно быть отрицательным", () => (hole_Gear_parsed >= 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "PowerWorm":
                        if (float.TryParse(PowerWorm, out float PowerWorm_parsed))
                            validateError(errors, "PowerWorm", "Значение параметра не должно быть отрицательным", () => (PowerWorm_parsed > 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "VelocityWorm":
                        if (float.TryParse(VelocityWorm, out float VelocityWorm_parsed))
                            validateError(errors, "VelocityWorm", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (VelocityWorm_parsed > 0) & (VelocityWorm_parsed < 1000000));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "MomentWorm":
                        if (float.TryParse(MomentWorm, out float MomentWorm_parsed))
                            validateError(errors, "MomentWorm", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (MomentWorm_parsed > 0) & (MomentWorm_parsed < 1000000));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "PowerGear":
                        if (float.TryParse(PowerGear, out float PowerGear_parsed))
                            validateError(errors, "PowerGear", "Значение параметра не должно быть отрицательным", () => (PowerGear_parsed > 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "VelocityGear":
                        if (float.TryParse(VelocityGear, out float VelocityGear_parsed))
                            validateError(errors, "VelocityGear", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (VelocityGear_parsed > 0) & (VelocityGear_parsed < 1000000));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "MomentGear":
                        if (float.TryParse(MomentGear, out float MomentGear_parsed))
                            validateError(errors, "MomentGear", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (MomentGear_parsed > 0) & (MomentGear_parsed < 1000000));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "KPD":
                        if (float.TryParse(KPD, out float KPD_parsed))
                            validateError(errors, "KPD", "Значение параметра должно быть в диапазоне от 0.1 до 1", () => (KPD_parsed >= 0.1) & (KPD_parsed <= 1));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "sigmaV":
                        if (float.TryParse(sigmaV, out float sigmaV_parsed))
                            validateError(errors, "sigmaV", "Значение параметра не должно быть отрицательным", () => (sigmaV_parsed >= 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "sigmaT":
                        if (float.TryParse(sigmaT, out float sigmaT_parsed))
                            validateError(errors, "sigmaT", "Значение параметра не должно быть отрицательным", () => (sigmaT_parsed >= 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "E_gear":
                        if (float.TryParse(E_gear, out float E_gear_parsed))
                            validateError(errors,  "E_gear", "Значение параметра не должно быть отрицательным", () => (E_gear_parsed >= 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "Puasson_gear":
                        if (float.TryParse(Puasson_gear, out float Puasson_gear_parsed))
                            validateError(errors, "Puasson_gear", "Для всех существующих материалов значение находится в пределах от 0 до 0.5", () => (Puasson_gear_parsed >= 0) & (Puasson_gear_parsed <= 0.5));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "E_worm":
                        if (float.TryParse(E_worm, out float E_worm_parsed))
                            validateError(errors, "E_worm", "Значение параметра не должно быть отрицательным", () => (E_worm_parsed > 0));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "Puasson_worm":
                        if (float.TryParse(Puasson_worm, out float Puasson_worm_parsed))
                            validateError(errors, "Puasson_worm", "Для всех существующих материалов значение находится в пределах от 0 до 0.5", () => (Puasson_worm_parsed >= 0) & (Puasson_worm_parsed <= 0.5));   
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "Ko":
                        if (float.TryParse(Ko, out float Ko_parsed))
                            validateError(errors, "Ko", "Значение параметра должно быть в диапазоне от 1 до 5", () => (Ko_parsed >= 1) & (Ko_parsed <= 5));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "Kv":
                        if (float.TryParse(Kv, out float Kv_parsed))
                            validateError(errors, "Kv", "Значение параметра должно быть в диапазоне от 1 до 6", () => (Kv_parsed >= 1) & (Kv_parsed <= 6));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "y":
                        if (float.TryParse(y, out float y_parsed))
                            validateError(errors, "y", "Значение параметра должно быть в диапазоне от 0.02 до 0.8", () => (y_parsed >= 0.02) & (y_parsed <= 0.8));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "time":
                        if (float.TryParse(time, out float time_parsed))
                            validateError(errors, "time", "Значение параметра не должно быть меньше 1", () => (time_parsed >= 1));
                        else
                            addError(errors, columnName, "Некорректный ввод данных");
                        break;

                    case "contact":
                        if (float.TryParse(contact, out float limit_parsed))
                            validateError(errorCalc, "contact", "Расчетное контактное напряжение больше чем допускаемое", () => (limit_parsed >= float.Parse(contact_calc)));
                        break;

                    case "bending":
                        if (float.TryParse(bending, out float bending_parsed))
                            validateError(errorCalc, "bending", "Расчетное напряжение изгиба больше чем допускаемое", () => (bending_parsed >= float.Parse(bending_calc)));
                        break;

                    case "temperature":
                        if (float.TryParse(temperature, out float temperature_parsed))
                            validateError(errorCalc, "temperature", "Температура масла должна быть менее чем 95 градусов", () => (temperature_parsed <= 95));
                        break;
                }


                bool IsValid = !errors.Values.Any(x => x != null);
                bool IsValidCalc = !errorCalc.Values.Any(x => x != null);

                if (IsValid & IsValidCalc)
                {
                    _window.buCalculateInModel.IsEnabled = _window.buCalculateinCalc.IsEnabled = true;
                    _window.buCreate_ComponentsInModel.IsEnabled = _window.buCreate_ComponentsInCalc.IsEnabled = true;
                   
                }
                else if (!IsValid || !IsValidCalc)
                {
                    if (!IsValid)
                    {
                        _window.buCalculateInModel.IsEnabled = _window.buCalculateinCalc.IsEnabled = false;
                        _window.buCreate_ComponentsInModel.IsEnabled = _window.buCreate_ComponentsInCalc.IsEnabled = false;
                        
                    }
                    else if (!IsValidCalc)
                    {
                        _window.buCalculateInModel.IsEnabled = _window.buCalculateinCalc.IsEnabled = true;
                        _window.buCreate_ComponentsInModel.IsEnabled = _window.buCreate_ComponentsInCalc.IsEnabled = false;
                   
                    }
                }

                if (errors.ContainsKey(columnName))
                    return errors[columnName];
                else if (errorCalc.ContainsKey(columnName))
                    return errorCalc[columnName];
                else
                    return null;
            }
        }

        public string Error
        {
            get { throw new NotImplementedException(); }
        }

        public DataValidation(MainWindow window)
        {
            this._window = window;
        }

        //Проверка словаря на наличие ошибки у поля
        private bool isErrorExist(Dictionary<string, string> dict, string property)
        {
            return dict.ContainsKey(property);
        }

        //Добавление ошибки в словарь
        private void addError(Dictionary<string, string> dict, string property, string message)
        {
            if (isErrorExist(dict, property))
                removeError(dict, property);

            dict.Add(property, message);
        }

        //Проверка поля на ошибку
        public bool validateError( Dictionary<string, string> dict, string property, string message, Func<bool> ruleCheck)
        {
            bool check = ruleCheck();
            if (!check)
            {
                addError(dict, property, message);
            }
            else
            {
                removeError(dict, property);
            }
            return check;
        }

        //Удаление ошибки из словаря
        private void removeError(Dictionary<string, string> dict, string property)
        {
            if (dict.ContainsKey(property))
                dict.Remove(property);
        }

    }
}
