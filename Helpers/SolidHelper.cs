using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace WormGearGenerator
{
    class SolidHelper
    {
        private SldWorks swApp;

        public SolidHelper()
        {
           swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        }


        public void CreateAssembly(string path)
        {
            ModelDoc2 swModel;

            //Новое окно сборки
            swModel = (ModelDoc2)swApp.NewAssembly();

            //Сохраняем сборку
            swModel.SaveAs3(path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

            //Закрываю сборку
           // swApp.CloseDoc(path); 
        }

        public void AddComponent(Worm worm, Gear gear, string assemblyPath)
        {
            
            ModelDoc2 swModel;
            AssemblyDoc swAssy;

            object components;
            object compNames;
            object transformationMatrix;
            object coorSysNames;
            double[] compXforms = new double[16];

            int errors = 0;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swAssy = (AssemblyDoc)swApp.ActiveDoc;

            string[] xcompnames = new string[2];

            xcompnames[0] = worm._path;

            xcompnames[1] = gear._path;


            string[] xcoorsysnames = new string[2];

            xcoorsysnames[0] = "Coordinate System1";
            xcoorsysnames[1] = "Coordinate System1";

            compXforms[0] = 1.0;
            compXforms[1] = 0.0;
            compXforms[2] = 0.0;

            compXforms[3] = 0.0;
            compXforms[4] = 1.0;
            compXforms[5] = 0.0;

            compXforms[6] = 0.0;
            compXforms[7] = 0.0;
            compXforms[8] = 1.0;

            compXforms[9] = 0.0;
            compXforms[10] = 0.0;
            compXforms[11] = 0.0;

            compXforms[12] = 0.0;

            compNames = xcompnames;
            coorSysNames = xcoorsysnames;
            transformationMatrix = compXforms;

            
            swModel = (ModelDoc2)swApp.ActivateDoc3(assemblyPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
            components = swAssy.AddComponents3(compNames, null, coorSysNames);
            swApp.CloseDoc(xcompnames[0]);
            swApp.CloseDoc(xcompnames[1]);
            swAssy.ForceRebuild();
            swModel.ViewZoomtofit2();
            swAssy = (AssemblyDoc)swApp.ActiveDoc;

        }

        public static string CreateDir(string path)
        {
            string guid = Guid.NewGuid().ToString();
            string root = path + "\\" + guid;

            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            else
            {
                MessageBox.Show("Файл с таким же именем существует");
            }
            return root;

        }

       
        public void AddMates(string assemblyPath, string assemblyName, string wormName, string gearName, float aw, float teethGear, float teethWorm, int rightOrLeft)
        {
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            Feature feat;
            MateFeatureData swMateData;
            ModelDocExtension swModelDocExt;
            DistanceMateFeatureData swDistMateData;
            int errorCode1 = 0;
            int mateSelMark;
            bool flip;
            if (rightOrLeft == 0)
                flip = false;
            else
                flip = true;

            swAssy = (AssemblyDoc)swApp.ActivateDoc3(assemblyPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
            swModel = (ModelDoc2)swAssy;
            swModelDocExt = swModel.Extension;
            mateSelMark = 1;

            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Front Plane@" + gearName +"-1@"+ assemblyName, "PLANE", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);

            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line9@Sketch2@" + gearName + "-1@" + assemblyName, "EXTSKETCHSEGMENT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
           
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Right Plane@" + wormName + "-1@" + assemblyName, "PLANE", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Right Plane", "PLANE", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);

            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line2@Sketch5@" + wormName + "-1@" + assemblyName, "EXTSKETCHSEGMENT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);

            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Point1@Origin@" + gearName + "-1@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Point1@Origin@" + wormName + "-1@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);

            feat = swModel.Extension.GetLastFeatureAdded();
            swMateData = (MateFeatureData)feat.GetDefinition();
            swDistMateData = (DistanceMateFeatureData)swMateData;
            swDistMateData.Distance = aw / 1000;
            swDistMateData.MinimumDistance = aw / 1000;
            swDistMateData.MaximumDistance = aw / 1000;
            feat.ModifyDefinition(swDistMateData, swAssy, null);

            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line2@Sketch5@" + wormName + "-1@" + assemblyName, "EXTSKETCHSEGMENT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Line9@Sketch2@" + gearName + "-1@" + assemblyName, "EXTSKETCHSEGMENT", 0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate5((int)swMateType_e.swMateGEAR, (int)swMateAlign_e.swMateAlignALIGNED, flip, 0, 0, 0, teethWorm/1000, teethGear/1000, 0, 0, 0, false, false, 0, out errorCode1);

            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(false);
            swModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }

    }
}
