using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Threading;

namespace WormGearGenerator
{
    class Worm
    {
        //Геометрические параметры червяка
        public float _da { get; set; }
        public float _d { get; set; }
        public float _df { get; set; }
        public float _pressureAngle { get; set; }
        public float _length { get; set; }
        public float _module { get; set; }
        public float _z { get; set; }
        public int _rightOrLeft { get; set; }
        public float _aw { get; set; }
        //Параметр материала
        public Material _material { get; set; }
        //Путь сохранения файла
        public string _path { get; set; }
        //Объект приложения SolidWorks
        SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        //Свойства для работы со шкалой загрузки
        UserProgressBar pb;
        bool retVal;
        int lRet;

        //Создание модели червяка
        public void create()
        {
            string destPath = _path;
            //Проверка существующего файла
            if (!File.Exists(destPath))
                //Содание модели червяка
                File.WriteAllBytes(destPath, Properties.Resources.WormTemp);
            //Изменение параметров модели червяка
            changeModel();
        }

        private void changeModel()
        {
            //Инициализация шкалы загрузки
            retVal = swApp.GetUserProgressBar(out pb);
            pb.Start(0, 100, "Создание компонента...");
            lRet = pb.UpdateProgress(20);

            //Инициализация объектов и переменных
            ModelDoc2 swModel;
            PartDoc swComp;
            int errors = 0;
            int warnings = 0;
            EquationMgr swEqnMgr = default(EquationMgr);
            string equation = null;

            //Открытие файла червяка
            swComp = (PartDoc)swApp.OpenDoc6(_path, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModel = (ModelDoc2)swComp;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            //Увеличение значения шкалы загрузки
            pb.UpdateTitle("Изменение компонента...");
            lRet = pb.UpdateProgress(25);

            //Подключение модуля уравнений
            swEqnMgr = (EquationMgr)swModel.GetEquationMgr();
            if (swEqnMgr == null)
                errorMsg(swApp, "Ошибка подключения к модели");
            swEqnMgr.AutomaticSolveOrder = true;
            swEqnMgr.AutomaticRebuild = true;

            //Изменение параметров модели
            try
            {
                equation = $"\"_da\" = {_da}mm";
                swEqnMgr.Equation[0] = equation;

                equation = $"\"D1@Sketch2\" = {_da}mm";
                swEqnMgr.Equation[1] = equation;

                equation = $"\"_df\"={_df}mm"; 
                swEqnMgr.Equation[2] = equation;

                equation = $"\"_d\"={_d}mm";
                swEqnMgr.Equation[3] = equation;

                equation = $"\"pressure_angle\"={_pressureAngle}";
                swEqnMgr.Equation[4] = equation;

                //Увеличение значения шкалы загрузки
                lRet = pb.UpdateProgress(30);

                equation = $"\"length\" = {_length}mm";
                swEqnMgr.Equation[5] = equation;

                equation = $"\"module\" = {_module}mm";
                swEqnMgr.Equation[6] = equation;

                equation = $"\"px\" = {_module* Math.PI}mm";
                swEqnMgr.Equation[7] = equation;

                float px = (float)(_module * Math.PI);
                float distance = (float)((int)((_length * 0.5 / px) + 1) * px + 0.5 * px);
                equation = $"\"distance\" = {distance}mm";
                swEqnMgr.Equation[8] = equation;

                //Увеличение значения шкалы загрузки
                lRet = pb.UpdateProgress(50);

                equation = $"\"z\" = {_z}";
                swEqnMgr.Equation[9] = equation;

                equation = $"\"rightOrLeft\" = {_rightOrLeft}";
                swEqnMgr.Equation[10] = equation;

                equation = $"\"aw\" = {_aw}mm";
                swEqnMgr.Equation[11] = equation;

                float cham2 = _da - _df;
                equation = $"\"cham2\" = {cham2}";
                swEqnMgr.Equation[12] = equation;

                equation = $"\"n_vitk\" = {(int)(_length / px) + 3}";
                swEqnMgr.Equation[13] = equation;

                float cham1 = (float)(cham2 * 0.5 * Math.Tan(_pressureAngle * (Math.PI / 180)));
                equation = $"\"cham1\" = {cham1}";
                swEqnMgr.Equation[14] = equation;

                //Увеличение значения шкалы загрузки
                lRet = pb.UpdateProgress(70);

                equation = $"\"D1@Boss-Extrude1\" = {_length}mm";
                swEqnMgr.Equation[15] = equation;

                equation = $"\"D1@Chamfer2\" = {cham2 * 0.5}mm";
                swEqnMgr.Equation[16] = equation;

                equation = $"\"D3@Chamfer2\" = {cham1}mm";
                swEqnMgr.Equation[17] = equation;

                equation = $"\"D1@Chamfer4\" = {cham1}mm";
                swEqnMgr.Equation[18] = equation;

                equation = $"\"D3@Chamfer4\" = {cham2 * 0.5}mm";
                swEqnMgr.Equation[19] = equation;

                equation = $"\"D1@Plane4\" = {distance}mm";
                swEqnMgr.Equation[20] = equation;

                equation = $"\"D1@Sketch4\" = {_d}mm";
                swEqnMgr.Equation[21] = equation;

                equation = $"\"D4@RightTeeth\" = {_z * px}mm";
                swEqnMgr.Equation[22] = equation;

                equation = $"\"D5@RightTeeth\" = {(int)(_length / px) + 3}mm";
                swEqnMgr.Equation[23] = equation;

                equation = $"\"D7@RightTeeth\" = {90}deg";
                swEqnMgr.Equation[24] = equation;

                equation = $"\"D1@RightPattern\" = {_z}";
                swEqnMgr.Equation[25] = equation;

                equation = $"\"D4@LeftHelix\" = {_z * px}";
                swEqnMgr.Equation[26] = equation;

                equation = $"\"D5@LeftHelix\" = {(int)(_length / px) + 3}";
                swEqnMgr.Equation[27] = equation;

                //Увеличение значения шкалы загрузки
                lRet = pb.UpdateProgress(80);

                equation = $"\"D7@LeftHelix\" = {90}deg";
                swEqnMgr.Equation[28] = equation;

                equation = $"\"D1@LeftPattern\" = {_z}";
                swEqnMgr.Equation[29] = equation;

                swEqnMgr.EvaluateAll();

                //Увеличение значения шкалы загрузки
                pb.UpdateTitle("Применение материала...");
                lRet = pb.UpdateProgress(90);

                //Применение материала к компоненту 
                if (_material != null)
                    setMaterial(swComp, _material.Name, _material.Database);

                //Перестроение документа
                swModel.ForceRebuild3(false);
                //Сохранение файла компонента
                swModel.SaveAs3(_path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

                //Увеличение значения шкалы загрузки, завершение работы шкалы загрузки
                pb.UpdateTitle("Сохранение компонента...");
                lRet = pb.UpdateProgress(100);
                pb.End();

            }
            //Обработка ошибок
            catch (Exception e)
            {
                errorMsg(swApp, "Ошибка в процессе редактирования модели! " + e.Message);
                pb.End();
            }
        }

        //Создание сообщений с ошибками
        private void errorMsg(SldWorks swApp, string Message)
        {
            swApp.SendMsgToUser2(Message, 0, 0);
            swApp.RecordLine("'*** WARNING - General");
            swApp.RecordLine("'*** " + Message);
            swApp.RecordLine("");
        }

        //Применение материала
        private void setMaterial(PartDoc myPart, string materialName, string database)
        {
            myPart.SetMaterialPropertyName2("default", database, materialName);
        }
    }
}
