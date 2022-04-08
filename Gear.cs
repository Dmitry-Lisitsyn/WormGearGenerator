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

        public string _path { get; set; }
        private string baseDirectory;

        SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        ModelDoc2 swModel;
        PartDoc swComp;

        public Gear(string directory)
        {
            baseDirectory = directory;
        }

        public void create()
        {
            //!Изменение имени компонентов, обработка существующих файлов

            string fileNameWorm = Directory.GetParent(baseDirectory).Parent.FullName + "\\res\\GearTemp.SLDPRT";
            string destPath = _path;
                //+ "\\" + "Gear.sldprt";

            if (!File.Exists(destPath))
                File.Copy(fileNameWorm, destPath);
            else
                Console.WriteLine("Файл с червяком существует, поэтому пока игнорируем");

            changeDimensions();
        }

        private void changeDimensions()
        {
            int errors = 0;
            int warnings = 0;
            EquationMgr swEqnMgr = default(EquationMgr);
            int nCount = 0;
            string equation = null;

           
            swComp = (PartDoc)swApp.OpenDoc6(_path, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModel = (ModelDoc2)swComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;

            swEqnMgr = (EquationMgr)swModel.GetEquationMgr();
            if (swEqnMgr == null)
                ErrorMsg(swApp, "Ошибка подключения к модели");

            swEqnMgr.AutomaticSolveOrder = true;
            swEqnMgr.AutomaticRebuild = true;

            try
            {
                equation = $"\"g_aw\" = {_aw}mm";
                swEqnMgr.Equation[0] = equation;

                //equation = $"\"D1@Sketch2\" = {_da}mm";
                //swEqnMgr.Equation[1] = equation;

                equation = $"\"g_b\"={_b}mm";
                swEqnMgr.Equation[2] = equation;

                equation = $"\"g_z\"={_z}mm";
                swEqnMgr.Equation[3] = equation;

                equation = $"\"g_df\"={_df}";
                swEqnMgr.Equation[4] = equation;

                equation = $"\"g_da\" = {_da}mm";
                swEqnMgr.Equation[5] = equation;

                equation = $"\"g_beta\" = {_beta}mm";
                swEqnMgr.Equation[6] = equation;

                equation = $"\"g_degpres\" = {_pressureAngle}";
                swEqnMgr.Equation[7] = equation;

                equation = $"\"w_df1\" = {_df1}mm";
                swEqnMgr.Equation[8] = equation;

                equation = $"\"w_da1\" = {_da1}";
                swEqnMgr.Equation[9] = equation;

                equation = $"\"w_d1\" = {_d1}";
                swEqnMgr.Equation[10] = equation;

                equation = $"\"px\" = {_px}";
                swEqnMgr.Equation[11] = equation;

                equation = $"\"w_z\" = {_z1}";
                swEqnMgr.Equation[12] = equation;

                equation = $"\"g_dw\" = {_dw}";
                swEqnMgr.Equation[13] = equation;

                equation = $"\"rightorLeft\" = {_rightOrLeft}";
                swEqnMgr.Equation[14] = equation;

                equation = $"\"g_dae\" = {_dae}mm";
                swEqnMgr.Equation[15] = equation;
               
                swEqnMgr.EvaluateAll();

                swModel.Rebuild((int)swRebuildOptions_e.swRebuildAll);
            }
            catch (Exception e)
            {
                ErrorMsg(swApp, "Ошибка в процессе редактирования модели!");
            }
        }

        private void ErrorMsg(SldWorks swApp, string Message)
        {
            swApp.SendMsgToUser2(Message, 0, 0);
            swApp.RecordLine("'*** WARNING - General");
            swApp.RecordLine("'*** " + Message);
            swApp.RecordLine("");
        }
    }
}
