using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
            string destPath = _path + "\\" + "Worm.sldprt";

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

            Console.WriteLine("Обрабатываем изменение компонента");
            swComp = (PartDoc)swApp.OpenDoc6(_path + "\\" + "Worm.sldprt", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModel = (ModelDoc2)swComp;
        }




    }
}
