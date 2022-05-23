using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Diagnostics;

namespace WormGearGenerator
{
    class Gear
    {
        public float _aw { get; set; } 
        public float _b { get; set; }
        public float _z { get; set; }
        public float _df { get; set; }
        public float _da { get; set; }
        public float _beta { get; set; }
        public float _pressureAngle { get; set; }
        public float _df1 { get; set; }
        public float _da1 { get; set; }
        public float _d1 { get; set; }
        public float _px { get; set; }
        public float _z1 { get; set; }
        public float _dw { get; set; }
        public float _dae { get; set; }
        public int _rightOrLeft { get; set; }
        public float _hole_diameter { get; set; }
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

        public void create()
        {
            string destPath = _path;
            //Проверка существующего файла
            if (!File.Exists(destPath))
                File.WriteAllBytes(destPath, Properties.Resources.GearTemp);
            changeModel();
            //Создание отверстия
            if (_hole_diameter != 0)
                createHole();
        }

        private void changeModel()
        {
            //Прогресс бар
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
            swComp = (PartDoc)swApp.OpenDoc6(_path, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
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
                equation = $"\"g_aw\" = {_aw}mm";
                swEqnMgr.Equation[0] = equation;

                equation = $"\"g_z\"={_z}mm";
                swEqnMgr.Equation[2] = equation;

                lRet = pb.UpdateProgress(35);

                equation = $"\"g_df\"={_df}";
                swEqnMgr.Equation[3] = equation;

                equation = $"\"g_da\" = {_da}mm";
                swEqnMgr.Equation[4] = equation;

                lRet = pb.UpdateProgress(45);

                equation = $"\"g_beta\" = {_beta}mm";
                swEqnMgr.Equation[5] = equation;

                equation = $"\"g_degpres\" = {_pressureAngle}";
                swEqnMgr.Equation[6] = equation;

                lRet = pb.UpdateProgress(55);

                equation = $"\"w_df1\" = {_df1}mm";
                swEqnMgr.Equation[7] = equation;

                equation = $"\"w_da1\" = {_da1}";
                swEqnMgr.Equation[8] = equation;

                lRet = pb.UpdateProgress(65);

                equation = $"\"w_d1\" = {_d1}";
                swEqnMgr.Equation[9] = equation;

                equation = $"\"px\" = {_px}";
                swEqnMgr.Equation[10] = equation;

                lRet = pb.UpdateProgress(75);

                equation = $"\"w_z\" = {_z1}";
                swEqnMgr.Equation[11] = equation;

                equation = $"\"g_dw\" = {_dw}";
                swEqnMgr.Equation[12] = equation;

                lRet = pb.UpdateProgress(85);

                equation = $"\"rightorLeft\" = {_rightOrLeft}";
                swEqnMgr.Equation[13] = equation;

                equation = $"\"g_dae\" = {_dae}mm";
                swEqnMgr.Equation[14] = equation;

                equation = $"\"g_width\"= {_b}mm";
                swEqnMgr.Equation[1] = equation;

                swEqnMgr.EvaluateAll();

                //Увеличение значения шкалы загрузки
                pb.UpdateTitle("Изменение компонента...");
                lRet = pb.UpdateProgress(95);

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
            catch (Exception e)
            {
                errorMsg(swApp, "Ошибка в процессе редактирования модели! " + e.Message);
            }
        }

        private void createHole()
        {
            ModelDoc2 swModel;

            swModel = (ModelDoc2)swApp.ActiveDoc;

            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.CreateCircle(0, 0, 0, _hole_diameter/1000, 0, 0);
            swModel.ClearSelection2(true);

            swModel.FeatureManager.FeatureCut3(false, false, false, (int)swEndConditions_e.swEndCondThroughAllBoth,
            (int)swEndConditions_e.swEndCondThroughAllBoth, _b/1000, _b/1000, false, false, false,
               false, 0, 0, false, false, false, false, false, true, true,
               false, false, false, (int)swStartConditions_e.swStartSketchPlane, 0, false);
        
        }

        private void setMaterial(PartDoc myPart, string materialName, string database)
        {
            myPart.SetMaterialPropertyName2("default", database, materialName);
        }

        private void errorMsg(SldWorks swApp, string Message)
        {
            swApp.SendMsgToUser2(Message, 0, 0);
            swApp.RecordLine("'*** WARNING - General");
            swApp.RecordLine("'*** " + Message);
            swApp.RecordLine("");
        }
    }
}
