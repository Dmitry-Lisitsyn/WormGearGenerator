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
    class SolidWorker
    {
        private SldWorks swApp;

        public SolidWorker()
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
            //worm._path + "\\" + "WormGenerate.sldprt";
            xcompnames[1] = gear._path;
            //gear._path + "\\" + "Gear.sldprt";

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

       
        public void AddMates(string assemblyPath, string assemblyName, string wormName, string gearName, float aw, float teethGear, float teethWorm)
        {
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            Feature feat;
            MateFeatureData swMateData;
            ModelDocExtension swModelDocExt;
            DistanceMateFeatureData swDistMateData;
            int errorCode1 = 0;
            int mateSelMark;

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
            swAssy.AddMate5((int)swMateType_e.swMateGEAR, (int)swMateAlign_e.swMateAlignALIGNED, true, 0, 0, 0, teethWorm/1000, teethGear/1000, 0, 0, 0, false, false, 0, out errorCode1);

            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(false);
            swModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }

    }
}
//public void ImportToAssembly(bool worm, bool gear)
//{

//    //Сохраняю новую сборку и сую туда компоненты
//    ModelDoc2 model = default(ModelDoc2);
//    model = (ModelDoc2)swApp.ActiveDoc;
//    AssemblyDoc swAss = default(AssemblyDoc);
//    PartDoc swWorm = default(PartDoc);

//    ModelDocExtension swModelDocExt = default(ModelDocExtension);

//    int errors = 0;
//    int warnings = 0;
//    string fileNameAss = null;
//    string fileNameWorm = null;

//    // Сохраняю путь для новой сборки
//    string path = null;
//    System.Windows.Forms.SaveFileDialog SFD = new System.Windows.Forms.SaveFileDialog();
//    SFD.Filter = "Assembly (*.sldasm)|*.sldasm";
//    SFD.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
//    if (SFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
//        path = SFD.FileName;
//    else
//        return;

//    //активное окно 
//    model = (ModelDoc2)swApp.ActiveDoc;

//    //Сохраняю сборку 
//    model.SaveAs3(path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

//    //имя новой сборки = че сохранял
//    string workingDirectory = System.Environment.CurrentDirectory;
//    //fileNameAss = Directory.GetParent(workingDirectory).Parent.FullName + "\\res\\Ass.SLDASM";
//    fileNameAss = path;

//    //открываю эту сборку 
//    swAss = (AssemblyDoc)swApp.OpenDoc6(fileNameAss, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
//    model = (ModelDoc2)swAss;

//    //путь до червяка
//    model = null;
//    fileNameWorm = Directory.GetParent(workingDirectory).Parent.FullName + "\\res\\HandWorm.SLDPRT";

//    //открываю червяка
//    swWorm = (PartDoc)swApp.OpenDoc6(fileNameWorm, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
//    model = (ModelDoc2)swWorm;

//    //возвращаюсь к сборке
//    model = (ModelDoc2)swApp.ActivateDoc3(fileNameAss, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);

//    //вставляю червяка в сборку
//    swAss = (AssemblyDoc)model;
//    swAss.AddComponent4(fileNameWorm, "Default", 0, 0, 0);

//    //закрываю док с червяком
//    swApp.CloseDoc(fileNameWorm);
//    model.ViewZoomtofit2();

//}