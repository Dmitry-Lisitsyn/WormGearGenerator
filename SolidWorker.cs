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
            swModel.SaveAs3(path + "\\" + "AssemblyTest.sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

            //Закрываю сборку
           // swApp.CloseDoc(path); 
        }

        public void AddComponent(Worm worm, Gear gear)
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

            xcompnames[0] = worm._path + "\\" + "Worm.sldprt";
            xcompnames[1] = gear._path + "\\" + "Gear.sldprt";

            string[] xcoorsysnames = new string[2];

            xcoorsysnames[0] = "Coordinate System1";
            xcoorsysnames[1] = "Coordinate System1";

            compXforms[0] = 1.0;
            compXforms[1] = 0.0;
            compXforms[2] = 0.0;
            // y-axis components of rotation
            compXforms[3] = 0.0;
            compXforms[4] = 1.0;
            compXforms[5] = 0.0;
            // z-axis components of rotation
            compXforms[6] = 0.0;
            compXforms[7] = 0.0;
            compXforms[8] = 1.0;

            // Add a translation vector to the transform (zero translation)
            compXforms[9] = 0.0;
            compXforms[10] = 0.0;
            compXforms[11] = 0.0;

            // Add a scaling factor to the transform
            compXforms[12] = 0.0;

            compNames = xcompnames;
            coorSysNames = xcoorsysnames;
            transformationMatrix = compXforms;

            
            swModel = (ModelDoc2)swApp.ActivateDoc3(worm._path + "\\" + "AssemblyTest.sldasm", true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
            components = swAssy.AddComponents3(compNames, null, coorSysNames);
            swApp.CloseDoc(xcompnames[0]);
            swApp.CloseDoc(xcompnames[1]);
            swAssy.ForceRebuild();
            swModel.ViewZoomtofit2();


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

        public void ImportToAssembly(bool worm, bool gear)
        {

            //Сохраняю новую сборку и сую туда компоненты
            ModelDoc2 model = default(ModelDoc2);
            model = (ModelDoc2)swApp.ActiveDoc;
            AssemblyDoc swAss = default(AssemblyDoc);
            PartDoc swWorm = default(PartDoc);

            ModelDocExtension swModelDocExt = default(ModelDocExtension);

            int errors = 0;
            int warnings = 0;
            string fileNameAss = null;
            string fileNameWorm = null;

            // Сохраняю путь для новой сборки
            string path = null;
            System.Windows.Forms.SaveFileDialog SFD = new System.Windows.Forms.SaveFileDialog();
            SFD.Filter = "Assembly (*.sldasm)|*.sldasm";
            SFD.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (SFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                path = SFD.FileName;
            else
                return;

            //активное окно 
            model = (ModelDoc2)swApp.ActiveDoc;

            //Сохраняю сборку 
            model.SaveAs3(path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);

            //имя новой сборки = че сохранял
            string workingDirectory = System.Environment.CurrentDirectory;
            //fileNameAss = Directory.GetParent(workingDirectory).Parent.FullName + "\\res\\Ass.SLDASM";
            fileNameAss = path;

            //открываю эту сборку 
            swAss = (AssemblyDoc)swApp.OpenDoc6(fileNameAss, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            model = (ModelDoc2)swAss;

            //путь до червяка
            model = null;
            fileNameWorm = Directory.GetParent(workingDirectory).Parent.FullName + "\\res\\HandWorm.SLDPRT";

            //открываю червяка
            swWorm = (PartDoc)swApp.OpenDoc6(fileNameWorm, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            model = (ModelDoc2)swWorm;

            //возвращаюсь к сборке
            model = (ModelDoc2)swApp.ActivateDoc3(fileNameAss, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);

            //вставляю червяка в сборку
            swAss = (AssemblyDoc)model;
            swAss.AddComponent4(fileNameWorm, "Default", 0, 0, 0);

            //закрываю док с червяком
            swApp.CloseDoc(fileNameWorm);
            model.ViewZoomtofit2();


        }


    }
}

//!если пустая сборка
//при запуске модуля чел выбирает путь сохранения сборки
//нажимает ок, сборка сохраняется
//поменял параметры, посчитал, нажимает построить
//высвечивается мини окно с подтверждением его хуйни, пути куда сохраняется
//нажимает окей, у меня есть путь сохранения его хуйни
//сохраняю туда компоненты, потом вызываю в сборку
//меняю если что надо, у меня есть файл сборки с измененными компонентами
// вставляю эту сборку в документ??? либо есть второе окно со сборкой
//вызов одной модели
//ModelDoc2 model = default(ModelDoc2);
//model = (ModelDoc2)swApp.ActiveDoc;

//AssemblyDoc swPart = default(AssemblyDoc);
//ModelDoc2 swModel = default(ModelDoc2);

//ModelDocExtension swModelDocExt = default(ModelDocExtension);

//int errors = 0;
//int warnings = 0;
//string fileName = "";

//string workingDirectory = System.Environment.CurrentDirectory;
//fileName = Directory.GetParent(workingDirectory).Parent.FullName + "\\res\\Ass.SLDASM";

//swPart = (AssemblyDoc)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
//swModel = (ModelDoc2)swPart;
//swModelDocExt = (ModelDocExtension)swModel.Extension;