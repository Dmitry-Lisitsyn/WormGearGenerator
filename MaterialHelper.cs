using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Windows;

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

            // если выбираем конкретную базу, то читаем ее
            if (database != null)
                ReadMaterials(database, ref list);
            else
            {
                //если нет, то считываем базы из корневой папки
                var databases = (string[])swApp.GetMaterialDatabases();

                // берем оттуда материалы
                if (databases != null)
                    foreach (var d in databases)
                        ReadMaterials(d, ref list);
            }

            // сортируем по имени и возвращаем
            return list.OrderBy(f => f.DisplayName).ToList();

        }

        private static void ReadMaterials(string database, ref List<Material> list)
        {
            // Проверяем, существует ли файл с базой 
            if (!File.Exists(database))
                Console.WriteLine("Указанной базы материалов не существует");

            try
            {
                // если найден, то открываем его
                using (var stream = File.Open(database, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // парсим как хмл
                    var xmlDoc = XDocument.Load(stream);

                    var materials = new List<Material>();

                    // проходим по структуре документа 
                    xmlDoc.Root.Elements("classification")?.ToList()?.ForEach(f =>
                    {
                        // берем классификацию
                        var classification = f.Attribute("name")?.Value;

                        // проходим по всем материалам
                        f.Elements("material").ToList().ForEach(material =>
                        {
                            // добавляем их в лист 
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

                    // Все что нашли добавляем в лист
                    if (materials.Count > 0)
                        list.AddRange(materials);
                }
            }
            catch (Exception ex)
            {
                // если у нас ошиибка, выводим ее
                MessageBox.Show(ex.Message);
            }
        }
    }
    public class Material
        {

        public string Classification { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Database { get; set; }

        public string DisplayName => $"{Name} ({Classification})";

        public bool DatabaseFileFound { get; set; }


    }
}
