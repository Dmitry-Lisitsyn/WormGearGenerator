using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AngelSix.SolidDna;


namespace WormGearGenerator
{
    class MaterialHelper
    {
        public SldWorks swApp;

        public MaterialHelper(SldWorks solidWorks)
        {
            swApp = solidWorks;

        }

        public List<Material> GetMaterials(string database = null)
        {

            // Пустой лист
            var list = new List<Material>();

            // If we are using a specified database, use that
            if (database != null)
                ReadMaterials(database, ref list);
            else
            {
                // Otherwise, get all known ones
                // Get the list of material databases (full paths to SLDMAT files)
                var databases = (string[])swApp.GetMaterialDatabases();

                // Get materials from each
                if (databases != null)
                    foreach (var d in databases)
                        ReadMaterials(d, ref list);
            }

            // Order the list
            return list.OrderBy(f => f.DisplayName).ToList();

        }

        private static void ReadMaterials(string database, ref List<Material> list)
        {
            // First make sure the file exists
            if (!File.Exists(database))
                throw new SolidDnaException(
                    SolidDnaErrors.CreateError(
                        SolidDnaErrorTypeCode.SolidWorksApplication,
                        SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileNotFoundError,
                        Localization.GetString("SolidWorksApplicationGetMaterialsFileNotFoundError")));

            try
            {
                // File should be an XML document, so attempt to read that
                using (var stream = File.Open(database, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Try and parse the Xml
                    var xmlDoc = XDocument.Load(stream);

                    var materials = new List<Material>();

                    // Iterate all classification nodes and inside are the materials
                    xmlDoc.Root.Elements("classification")?.ToList()?.ForEach(f =>
                    {
                        // Get classification name
                        var classification = f.Attribute("name")?.Value;

                        // Iterate all materials
                        f.Elements("material").ToList().ForEach(material =>
                        {
                            // Add them to the list
                            materials.Add(new Material
                            {
                                Database = database,
                                DatabaseFileFound = true,
                                Classification = classification,
                                Name = material.Attribute("name")?.Value,
                                Description = material.Attribute("description")?.Value,
                            });
                        });
                    });

                    // If we found any materials, add them
                    if (materials.Count > 0)
                        list.AddRange(materials);
                }
            }
            catch (Exception ex)
            {
                // If we crashed for any reason during parsing, wrap in SolidDna exception
                if (!File.Exists(database))
                    throw new SolidDnaException(
                        SolidDnaErrors.CreateError(
                            SolidDnaErrorTypeCode.SolidWorksApplication,
                            SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileFormatError,
                            Localization.GetString("SolidWorksApplicationGetMaterialsFileFormatError"),
                            ex));
            }
        }
    }
}
