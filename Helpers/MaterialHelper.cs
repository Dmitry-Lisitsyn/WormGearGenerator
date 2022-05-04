using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Windows;

namespace WormGearGenerator
{
    /// <summary>
    /// Предоставляет возможность чтения баз данных материалов SolidWorks 
    /// </summary>
    class MaterialHelper
    {
        public SldWorks swApp;

        public MaterialHelper(SldWorks solidWorks)
        {
            swApp = solidWorks;
        }

        /// <summary>
        /// Возвращает все доступные материалы из SolidWorks
        /// </summary>
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

        /// <summary>
        ///Заполняет передаваемый список названиями материалов и их физическими характеристиками
        /// </summary>
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

                        string ex = null;
                        string nuxy = null;
                        string sigxt = null;
                        string sigyld = null;

                        // проходим по всем материалам
                        f.Elements("material").ToList().ForEach(material =>
                        {
                            material.Elements("physicalproperties").ToList().ForEach(physic =>
                            {
                                //Модуль упругости
                                physic.Elements("EX").ToList().ForEach(prop1 => { ex = prop1.Attribute("value")?.Value; });
                                //Коэффициент Пуассона
                                physic.Elements("NUXY").ToList().ForEach(prop2 =>{ nuxy = prop2.Attribute("value")?.Value; });
                                //Предел прочности
                                physic.Elements("SIGXT").ToList().ForEach(prop3 => { sigxt = prop3.Attribute("value")?.Value; });
                                //Предел текучести
                                physic.Elements("SIGYLD").ToList().ForEach(prop4 =>{ sigyld = prop4.Attribute("value")?.Value; });
                            });

                            // добавляем их в лист 
                            materials.Add(new Material
                            {
                                Database = database,
                                DatabaseFileFound = true,
                                Classification = classification,
                                Name = material.Attribute("name")?.Value,
                                Description = material.Attribute("description")?.Value,
                                Elastic_modulus = ex,
                                Poisson_ratio = nuxy,
                                Tensile_strength = sigxt,
                                Yield_strength = sigyld,
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
    /// <summary>
    /// Предоставляет свойства материала
    /// </summary>
    public class Material
        {
        //Классификация матриала
        public string Classification { get; set; }

        //Название материала
        public string Name { get; set; }

        //Описание материала
        public string Description { get; set; }

        //База данных материала
        public string Database { get; set; }

        //Отображаемое имя материала
        public string DisplayName => $"{Name} ({Classification})";

        //Статус поиска базы данных материала
        public bool DatabaseFileFound { get; set; }

        //Модуль упругости
        public string Elastic_modulus { get; set; }

        //Коэффициент пуассона
        public string Poisson_ratio { get; set; }

        //Предел прочности
        public string Tensile_strength { get; set; }

        //Предел текучести
        public string Yield_strength { get; set; }

    }
}
