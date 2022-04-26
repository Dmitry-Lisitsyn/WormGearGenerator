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
        public Material _material { get; set; }

        public string _path { get; set; }
        private string baseDirectory;

        SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");

        public Gear(string directory)
        {
            baseDirectory = directory;
        }

        public void create()
        {
            //string fileNameWorm = Directory.GetParent(baseDirectory).Parent.FullName + "\\res\\Part1.SLDPRT";
            string destPath = _path;

            if (!File.Exists(destPath))
                File.WriteAllBytes(destPath, Properties.Resources.GearTemp);
                //File.Copy(fileNameWorm, destPath);

            changeModel();
            if (_hole_diameter != 0)
                createHole();
        }

        private void changeModel()
        {
            ModelDoc2 swModel;
            PartDoc swComp;
            int errors = 0;
            int warnings = 0;
            EquationMgr swEqnMgr = default(EquationMgr);
            string equation = null;

           
            swComp = (PartDoc)swApp.OpenDoc6(_path, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModel = (ModelDoc2)swComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;

            swEqnMgr = (EquationMgr)swModel.GetEquationMgr();
            if (swEqnMgr == null)
                errorMsg(swApp, "Ошибка подключения к модели");
         
            swEqnMgr.AutomaticRebuild = true;

            try
            {
                equation = $"\"g_aw\" = {_aw}mm";
                swEqnMgr.Equation[0] = equation;

                equation = $"\"g_z\"={_z}mm";
                swEqnMgr.Equation[2] = equation;

                equation = $"\"g_df\"={_df}";
                swEqnMgr.Equation[3] = equation;

                equation = $"\"g_da\" = {_da}mm";
                swEqnMgr.Equation[4] = equation;

                equation = $"\"g_beta\" = {_beta}mm";
                swEqnMgr.Equation[5] = equation;

                equation = $"\"g_degpres\" = {_pressureAngle}";
                swEqnMgr.Equation[6] = equation;

                equation = $"\"w_df1\" = {_df1}mm";
                swEqnMgr.Equation[7] = equation;

                equation = $"\"w_da1\" = {_da1}";
                swEqnMgr.Equation[8] = equation;

                equation = $"\"w_d1\" = {_d1}";
                swEqnMgr.Equation[9] = equation;

                equation = $"\"px\" = {_px}";
                swEqnMgr.Equation[10] = equation;

                equation = $"\"w_z\" = {_z1}";
                swEqnMgr.Equation[11] = equation;

                equation = $"\"g_dw\" = {_dw}";
                swEqnMgr.Equation[12] = equation;

                equation = $"\"rightorLeft\" = {_rightOrLeft}";
                swEqnMgr.Equation[13] = equation;

                equation = $"\"g_dae\" = {_dae}mm";
                swEqnMgr.Equation[14] = equation;

                equation = $"\"g_width\"= {_b}mm";
                swEqnMgr.Equation[1] = equation;

                swEqnMgr.EvaluateAll();

                if (_material != null)
                    setMaterial(swComp, _material.Name, _material.Database);

                swModel.ForceRebuild3(false);
                
                swModel.SaveAs3(_path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
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
