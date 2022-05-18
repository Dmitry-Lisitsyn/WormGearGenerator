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
        //Объект приложения SolidWorks
        private SldWorks swApp;

        //Свойства для работы со шкалой загрузки
        UserProgressBar pb;
        bool retVal;
        int lRet;

        public SolidHelper()
        {
            // Считываение текущей сессии запущенного приложения SolidWorks
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
        }

        //Создание файла сборки для компонентов
        public void CreateAssembly(string path)
        {
            ModelDoc2 swModel;

            //Новое окно сборки
            swModel = (ModelDoc2)swApp.NewAssembly();

            //Сохраняем сборку
            swModel.SaveAs3(path, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }
        //Добавление компонентов в сборку
        public void AddComponents(Worm worm, Gear gear, string assemblyPath)
        {
            //начальные данные
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            object components;
            object compNames;
            object coorSysNames;
            double[] compXforms = new double[16];
            int errors = 0;

            //Считывание открытого окна сборки
            swAssy = (AssemblyDoc)swApp.ActiveDoc;

            //Добаление имен компонентов в массив
            string[] xcompnames = new string[2];
            //IF компоненты построены
            xcompnames[0] = worm._path;
            xcompnames[1] = gear._path;

            //Добавление координатных систем компонентов в массив
            string[] xcoorsysnames = new string[2];
            xcoorsysnames[0] = "Coordinate System1";
            xcoorsysnames[1] = "Coordinate System1";

            //Присваивание массивов к объектам
            compNames = xcompnames;
            coorSysNames = xcoorsysnames;

            //Активация документа сборки
            swModel = (ModelDoc2)swApp.ActivateDoc3(assemblyPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errors);
            //Добавление компонентов в документ сборки
            components = swAssy.AddComponents3(compNames, null, coorSysNames);
            //Закрытие открытых файлов компонентов в среде SolidWorks
            swApp.CloseDoc(xcompnames[0]);
            swApp.CloseDoc(xcompnames[1]);
            //Перестроение документа
            swAssy.ForceRebuild();
            //Фокус камеры на компоненты
            swModel.ViewZoomtofit2();
        }
        //Создание зависимостей внутри сборки
        public void AddMates(string assemblyPath, string assemblyName, string wormName, string gearName, float aw, float teethGear, float teethWorm, int rightOrLeft)
        {
            //Инициализация объектов
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            Feature feat;
            MateFeatureData swMateData;
            ModelDocExtension swModelDocExt;
            DistanceMateFeatureData swDistMateData;
            int errorCode1 = 0;
            int mateSelMark;
            //Обработка направления вращения
            bool flip;
            if (rightOrLeft == 0)
                flip = false;
            else
                flip = true;

            //Активация документа со сборкой
            swAssy = (AssemblyDoc)swApp.ActivateDoc3(assemblyPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
            swModel = (ModelDoc2)swAssy;
            //Иницииализация объекта для выбора элементов компонентов 
            swModelDocExt = swModel.Extension;
            mateSelMark = 1;

            //Инициализация шкалы загрузки
            retVal = swApp.GetUserProgressBar(out pb);
            pb.Start(0, 100, "Создание зависимостей...");
            lRet = pb.UpdateProgress(30);


            //Создание зависимости плоскости червячного колеса с плоскостью сборки
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Front Plane@" + gearName + "-1@" + assemblyName, "PLANE", 0, 0, 0,
                true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0,
                true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false,
                0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(40);

            //Создание зависимости центра червячного колеса с центром координат сборки
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line9@Sketch2@" + gearName + "-1@" + assemblyName, "EXTSKETCHSEGMENT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Point1@Origin", "EXTSKETCHPOINT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false,
                0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(50);

            //Создание зависимости плоскости червяка с плоскостью сборки
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Right Plane@" + wormName + "-1@" + assemblyName, "PLANE",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Right Plane", "PLANE",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false,
                0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(60);

            //Создание зависимости центра червяка с плоскостью
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line2@Sketch5@" + wormName + "-1@" + assemblyName, "EXTSKETCHSEGMENT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Front Plane", "PLANE",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false,
                0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(70);

            //Создание зависимости расстояния между центрами компонентов
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Point1@Origin@" + gearName + "-1@" + assemblyName, "EXTSKETCHPOINT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Point1@Origin@" + wormName + "-1@" + assemblyName, "EXTSKETCHPOINT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate3((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignALIGNED, false,
                0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(80);

            //Изменение значения зависимости межосевого расстояния компонентов
            feat = swModel.Extension.GetLastFeatureAdded();
            swMateData = (MateFeatureData)feat.GetDefinition();
            swDistMateData = (DistanceMateFeatureData)swMateData;
            swDistMateData.Distance = aw / 1000;
            swDistMateData.MinimumDistance = aw / 1000;
            swDistMateData.MaximumDistance = aw / 1000;
            feat.ModifyDefinition(swDistMateData, swAssy, null);

            //Создание зависимости вращения между компонентами
            swModel.ClearSelection2(true);
            swModelDocExt.SelectByID2("Line2@Sketch5@" + wormName + "-1@" + assemblyName, "EXTSKETCHSEGMENT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Line9@Sketch2@" + gearName + "-1@" + assemblyName, "EXTSKETCHSEGMENT",
                0, 0, 0, true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
            swAssy.AddMate5((int)swMateType_e.swMateGEAR, (int)swMateAlign_e.swMateAlignALIGNED, flip,
                0, 0, 0, teethWorm / 1000, teethGear / 1000, 0, 0, 0, false, false, 0, out errorCode1);
            //Увеличение значения шкалы загрузки
            lRet = pb.UpdateProgress(90);

            //Сброс выделения
            swModel.ClearSelection2(true);
            //Перестроение документа
            swModel.ForceRebuild3(false);
            //Сохранение документа
            swModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
            //Увеличение значения шкалы загрузки, завершение работы шкалы загрузки
            pb.UpdateTitle("Сохранение сборки...");
            lRet = pb.UpdateProgress(100);
            pb.End();
        }

        public void addToAssembly(string currentAssemblyName, string generateAssemblyPath)
        {
            //Инициализация объектов
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            object components;
            int errorCode1 = 0;

            //Активация документа со сборкой
            swAssy = (AssemblyDoc)swApp.ActivateDoc3(currentAssemblyName, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
            swModel = (ModelDoc2)swAssy;

            //Добавление сборки
            components = swAssy.AddComponent5(generateAssemblyPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "",
                MainWindow.pointsOfOriginComponent[0], MainWindow.pointsOfOriginComponent[1], MainWindow.pointsOfOriginComponent[2]);
            swApp.CloseDoc(generateAssemblyPath);

            //Сохранение документа
            swModel.SaveAs3(currentAssemblyName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }

        public void addMatesToFaces(string selectedComponent, string selectedEntity, string generalAssemblyPath, string assemblyName, string wormName, string gearName, string WGcomponent)
        {
            //Инициализация начальных параметров
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            Face2 swFace;
            Body2 swBody;
            SelectData swSelData;
            ModelDocExtension swModelDocExt;
            Entity swEntity = null;
            SelectionMgr swSelMgr;

            int mateSelMark;
            int errorCode1 = 0;
            string currentFaceName;
            string selectInWG = null;

            //Активация документа со сборкой
            swAssy = (AssemblyDoc)swApp.ActivateDoc3(generalAssemblyPath, true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref errorCode1);
            swModel = (ModelDoc2)swAssy;

            //Иницииализация объектов для выбора элементов компонентов 
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            mateSelMark = 1;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swSelData = swSelMgr.CreateSelectData();

            //Считывание тела детали, выбранной пользователем
            swBody = getByName(swAssy, selectedComponent);

            //Выбор грани
            swFace = (Face2)swBody.GetFirstFace();
            do
            {
                currentFaceName = swModel.GetEntityName(swFace);
                if (currentFaceName == selectedEntity)
                {
                    swEntity = (Entity)swFace;
                    swEntity.Select4(true, swSelData);
                }
                swFace = (Face2)swFace.GetNextFace();
            } while (swFace != null || swEntity != null);

            //Создание зависимости между выбранной гранью и компонентом
            if (WGcomponent != null)
            {
                if (WGcomponent == "WormCylinder")
                {
                    selectInWG = "Line2@Sketch5@" + assemblyName + "-1@" + swModel.GetTitle() + "/" + wormName + "-1@" + assemblyName;
                    swModelDocExt.SelectByID2(selectInWG, "EXTSKETCHSEGMENT", 0, 0, 0,
                        true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
                    swAssy.AddMate5((int)swMateType_e.swMateCONCENTRIC, (int)swMateAlign_e.swMateAlignALIGNED,
                    false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out errorCode1);
                }
                else if (WGcomponent == "GearCylinder")
                {
                    selectInWG = "Line9@Sketch2@" + assemblyName + "-1@" + swModel.GetTitle() + "/" + gearName + "-1@" + assemblyName;
                    swModelDocExt.SelectByID2(selectInWG, "EXTSKETCHSEGMENT", 0, 0, 0,
                        true, mateSelMark, null, (int)swSelectOption_e.swSelectOptionDefault);
                    swAssy.AddMate5((int)swMateType_e.swMateCONCENTRIC, (int)swMateAlign_e.swMateAlignALIGNED,
                    false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out errorCode1);
                }
                else if (WGcomponent == "GearFace")
                {
                    //Считывание тела колеса
                    var swEntityGear = default(Entity);
                    var swBodyGear = getByName(swAssy, gearName + "-1");
                    var swFaceGear = (Face2)swBodyGear.GetFirstFace();
                    do
                    {
                        currentFaceName = swModel.GetEntityName(swFaceGear);
                        if (currentFaceName == "FrontTempFace")
                        {
                            swEntityGear = (Entity)swFaceGear;
                            swEntityGear.Select4(true, swSelData);
                        }
                        swFaceGear = (Face2)swFaceGear.GetNextFace();
                    } while (swFaceGear != null);

                    swAssy.AddMate5((int)swMateType_e.swMateANGLE, (int)swMateAlign_e.swMateAlignALIGNED,
                    false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out errorCode1);
                }
                else if (WGcomponent == "WormFace")
                {
                    //Считывание тела червяка
                    var swEntityWorm = default(Entity);
                    var swBodyWorm = getByName(swAssy, gearName + "-1");
                    var swFaceWorm = (Face2)swBodyWorm.GetFirstFace();
                    do
                    {
                        currentFaceName = swModel.GetEntityName(swFaceWorm);
                        if (currentFaceName == "LeftTempFace")
                        {
                            swEntityWorm = (Entity)swFaceWorm;
                            swEntityWorm.Select4(true, swSelData);
                        }
                        swFaceWorm = (Face2)swFaceWorm.GetNextFace();
                    } while (swFaceWorm != null);

                    swAssy.AddMate5((int)swMateType_e.swMateANGLE, (int)swMateAlign_e.swMateAlignALIGNED,
                    false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out errorCode1);
                }    
            }

            //Сброс выделения
            swModel.ClearSelection2(true);
            //Перестроение документа
            swModel.ForceRebuild3(false);
            //Сохранение документа
            swModel.SaveAs3(generalAssemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }

        private Body2 getByName(AssemblyDoc assy, string name)
        {
            Component2 swComp;
            Body2 swBody = null;
            string[] str;

            if (name.Contains("@"))
                str = name.Split('@');
            else
            {
                var list = new List<string>();
                list.Add(name);
                str = list.ToArray();
            }

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i].Contains("/"))
                {
                    var substr = str[i].Split('/');
                    swComp = assy.GetComponentByName(substr[1]);
                }
                else
                    swComp = assy.GetComponentByName(str[i]);

                swBody = (Body2)swComp.GetBody();
                if (swBody != null)
                    return swBody;

            }
            return swBody;
        }

        //Сохранение пустой сборки
        public void saveInitialAssembly(string filename)
        {
            ModelDoc2 swModel;
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (File.Exists(filename))
                File.Delete(filename);

            swModel.SaveAs3(filename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
        }

    }
}
