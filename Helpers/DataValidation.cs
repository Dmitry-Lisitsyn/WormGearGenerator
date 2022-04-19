﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WormGearGenerator.Helpers
{
    public class DataValidation : System.ComponentModel.IDataErrorInfo
    {
        public float deg_TeethValue { get; set; }
        //Червяк
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
        public string PowerWorm { get; set; } = "0.1"; //>=0
        public string PowerGear { get; set; }
        public string VelocityWorm { get; set; } = "1000"; //>=1000000
        public string VelocityGear { get; set; }
        public string MomentWorm { get; set; }
        public string MomentGear { get; set; }

        public string KPD { get; set; } // 0.1 - 1

        //Параметры материалов
        public string sigmaV { get; set; } = "250"; //>=0
        public string sigmaT { get; set; } = "160"; //>=0
        public string E_worm { get; set; } = "206000";//>=0
        public string E_gear { get; set; } = "101000";//>=0
        public string Puasson_worm { get; set; } = "0.3";//>=0
        public string Puasson_gear { get; set; } = "0.31";//>=0

        //Коэффициенты
        public string Ko { get; set; } = "1.200"; //1 - 5
        public string Kv { get; set; } = "1.042"; //1 - 6
        public string y { get; set; } = "0.125"; // 0.02 - 0.8
        public string time { get; set; } = "10000"; // >=1


        private MainWindow _window;

        private Dictionary<string, string> errors =  new Dictionary<string, string>();

        public string this[string columnName]
        {
            get
            {
                string error = String.Empty;
                switch (columnName)
                {
                    case "kol_vitkov":

                        if (int.TryParse(kol_vitkov, out int kol_vitkov_parsed))
                            validateError("kol_vitkov", "Значение витков должно быть в диапазоне от 1 до 4", () => (kol_vitkov_parsed > 0) & (kol_vitkov_parsed <= 4));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "length_Worm":

                        if (float.TryParse(length_Worm, out float length_Worm_parsed))
                            validateError("length_Worm", "Значение витков должно быть в диапазоне от 1 до 10000 мм", () => (length_Worm_parsed > 1) & (length_Worm_parsed < 10000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "av_diamWorm":

                        if (float.TryParse(av_diamWorm, out float av_diamWorm_parsed))
                            validateError("av_diamWorm", "Значение параметра не должно быть меньше или равно нулю", () => (av_diamWorm_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "koef_diamWorm":
                        if (float.TryParse(koef_diamWorm, out float koef_diamWorm_parsed))
                            validateError("koef_diamWorm", "Значение параметра не должно быть меньше или равно нулю", () => (koef_diamWorm_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "teeth_Gear":
                        if (int.TryParse(teeth_Gear, out int teeth_Gear_parsed))
                            validateError("teeth_Gear", "Значение параметра должно быть в диапазоне от 8 до 1999", () => (teeth_Gear_parsed >= 8) & (teeth_Gear_parsed < 2000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "width_Gear":
                        if (float.TryParse(width_Gear, out float width_Gear_parsed))
                            validateError("width_Gear", "Ширина червячного колеса не должна быть отрицательной или равняться нулю", () => (width_Gear_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "koef_Smesh":
                        if (float.TryParse(koef_Smesh, out float koef_Smesh_parsed))
                            validateError("koef_Smesh", "Значение параметра должно быть в диапазоне от -1 до 1", () => (koef_Smesh_parsed >= -1) & (koef_Smesh_parsed <= 1));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "hole_Gear":
                        if (float.TryParse(hole_Gear, out float hole_Gear_parsed))
                            validateError("hole_Gear", "Значение параметра не должно быть отрицательным", () => (hole_Gear_parsed >= 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "PowerWorm":
                        if (float.TryParse(PowerWorm, out float PowerWorm_parsed))
                            validateError("PowerWorm", "Значение параметра не должно быть отрицательным", () => (PowerWorm_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "VelocityWorm":
                        if (float.TryParse(VelocityWorm, out float VelocityWorm_parsed))
                            validateError("VelocityWorm", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (VelocityWorm_parsed > 0) & (VelocityWorm_parsed < 1000000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "MomentWorm":
                        if (float.TryParse(MomentWorm, out float MomentWorm_parsed))
                            validateError("MomentWorm", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (MomentWorm_parsed > 0) & (MomentWorm_parsed < 1000000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "PowerGear":
                        if (float.TryParse(PowerGear, out float PowerGear_parsed))
                            validateError("PowerGear", "Значение параметра не должно быть отрицательным", () => (PowerGear_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "VelocityGear":
                        if (float.TryParse(VelocityGear, out float VelocityGear_parsed))
                            validateError("VelocityGear", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (VelocityGear_parsed > 0) & (VelocityGear_parsed < 1000000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "MomentGear":
                        if (float.TryParse(MomentGear, out float MomentGear_parsed))
                            validateError("MomentGear", "Значение параметра должно быть в диапазоне от 0 до 999999", () => (MomentGear_parsed > 0) & (MomentGear_parsed < 1000000));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "KPD":
                        if (float.TryParse(KPD, out float KPD_parsed))
                            validateError("KPD", "Значение параметра должно быть в диапазоне от 0.1 до 1", () => (KPD_parsed >= 0.1) & (KPD_parsed <= 1));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "sigmaV":
                        if (float.TryParse(sigmaV, out float sigmaV_parsed))
                            validateError("sigmaV", "Значение параметра не должно быть отрицательным", () => (sigmaV_parsed >= 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "sigmaT":
                        if (float.TryParse(sigmaT, out float sigmaT_parsed))
                            validateError("sigmaT", "Значение параметра не должно быть отрицательным", () => (sigmaT_parsed >= 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "E_gear":
                        if (float.TryParse(E_gear, out float E_gear_parsed))
                            validateError("E_gear", "Значение параметра не должно быть отрицательным", () => (E_gear_parsed >= 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "Puasson_gear":
                        if (float.TryParse(Puasson_gear, out float Puasson_gear_parsed))
                            validateError("Puasson_gear", "Для всех существующих материалов значение находится в пределах от 0 до 0.5", () => (Puasson_gear_parsed >= 0) & (Puasson_gear_parsed <= 0.5));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "E_worm":
                        if (float.TryParse(E_worm, out float E_worm_parsed))
                            validateError("E_worm", "Значение параметра не должно быть отрицательным", () => (E_worm_parsed > 0));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "Puasson_worm":
                        if (float.TryParse(Puasson_worm, out float Puasson_worm_parsed))
                            validateError("Puasson_worm", "Для всех существующих материалов значение находится в пределах от 0 до 0.5", () => (Puasson_worm_parsed >= 0) & (Puasson_worm_parsed <= 0.5));   
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "Ko":
                        if (float.TryParse(Ko, out float Ko_parsed))
                            validateError("Ko", "Значение параметра должно быть в диапазоне от 1 до 5", () => (Ko_parsed >= 1) & (Ko_parsed <= 5));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "Kv":
                        if (float.TryParse(Kv, out float Kv_parsed))
                            validateError("Kv", "Значение параметра должно быть в диапазоне от 1 до 6", () => (Kv_parsed >= 1) & (Kv_parsed <= 6));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "y":
                        if (float.TryParse(y, out float y_parsed))
                            validateError("y", "Значение параметра должно быть в диапазоне от 0.02 до 0.8", () => (y_parsed >= 0.02) & (y_parsed <= 0.8));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                    case "time":
                        if (float.TryParse(time, out float time_parsed))
                            validateError("time", "Значение параметра не должно быть меньше 1", () => (time_parsed >= 1));
                        else
                            addError(columnName, "Некорректный ввод данных");
                        break;

                }

               bool IsValid = !errors.Values.Any(x => x != null);
                if (IsValid)
                {
                    _window.buCalculateInModel.IsEnabled = _window.buCalculateinCalc.IsEnabled = true;
                    _window.buCreate_ComponentsInModel.IsEnabled = _window.buCreate_ComponentsInCalc.IsEnabled = true;
                }
                else
                {
                    _window.buCalculateInModel.IsEnabled = _window.buCalculateinCalc.IsEnabled = false;
                    _window.buCreate_ComponentsInModel.IsEnabled = _window.buCreate_ComponentsInCalc.IsEnabled = false;   
                }
                return errors.ContainsKey(columnName) ? errors[columnName] : null;
                
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

        private bool isErrorExist(string property)
        {
            return errors.ContainsKey(property);
        }

        private void addError(string property, string message)
        {
            if (isErrorExist(property))
                removeError(property);

            errors.Add(property, message);
        }

        public bool validateError(string property, string message, Func<bool> ruleCheck)
        {
            bool check = ruleCheck();
            if (!check)
            {
                addError(property, message);
            }
            else
            {
                removeError(property);
            }
            return check;
        }

        private void removeError(string property)
        {
            if (errors.ContainsKey(property))
                errors.Remove(property);
        }

        private void Clear()
        {
            errors.Clear();
        }


    }
}
