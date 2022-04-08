using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace WormGearGenerator
{
    class Worm
    {
        public float _da { get; set; }
        public float _d { get; set; }
        public float _df { get; set; }
        public float _pressureAngle { get; set; }
        public float _length { get; set; }
        public float _module { get; set; }
        public float _z { get; set; }
        public int _rightOrLeft { get; set; }
        public float _aw { get; set; }
        
        public string _path { get; set; }
        private string baseDirectory;

        SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        ModelDoc2 swModel;
        PartDoc swComp;

        public Worm(string directory)
        {
            baseDirectory = directory;
        }


        public void create()
        {
            //!Изменение имени компонентов, обработка существующих файлов

            string fileNameWorm = Directory.GetParent(baseDirectory).Parent.FullName + "\\res\\WormTemp.SLDPRT";
            string destPath = _path;
                //+ "\\" + "WormGenerate.sldprt";

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

            //nCount = swEqnMgr.GetCount();
            //for (int i = 0; i < nCount; i++)
            //{
            //    Debug.Print("  Equation(" + i + ")  = " + swEqnMgr.get_Equation(i));
            //    Debug.Print("    Value = " + swEqnMgr.get_Value(i));
            //    Debug.Print("    Index = " + swEqnMgr.Status);
            //    Debug.Print("    Global variable? " + swEqnMgr.get_GlobalVariable(i));
            //}

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

                equation = $"\"D7@LeftHelix\" = {90}deg";
                swEqnMgr.Equation[28] = equation;

                equation = $"\"D1@LeftPattern\" = {_z}";
                swEqnMgr.Equation[29] = equation;

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
