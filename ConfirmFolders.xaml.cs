using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WormGearGenerator
{
    /// <summary>
    /// Логика взаимодействия для ConfirmFolders.xaml
    /// </summary>
    public partial class ConfirmFolders : Window
    {
        //Начальный путь сохранения, передающийся из MainWindow
        private string _path { get; set; }
        //Свойство, отвечающее за построение червяка
        private bool _isWorm { get; set; }
        //Свойство, отвечающее за построение червячного колеса
        private bool _isGear { get; set; }
        //Путь сохранения сборки
        public string pathAssembly { get; set; }
        //Путь сохранения червяка
        public string pathWorm { get; set; }
        //Путь сохранения червячного колеса
        public string pathGear { get; set; }

        //Начальная инициализация, передача значений из MainWindow
        public ConfirmFolders(string path, bool isWorm, bool isGear)
        {
            _path = path;
            _isWorm = isWorm;
            _isGear = isGear;

            //Инициализация элементов окна
            InitializeComponent();
            //Передача данных в поля окна
            InitializeFolders();
        }

        /// <summary>
        /// Начальная инициализация значений элементов окна
        /// </summary>
        public void InitializeFolders()
        {
            //Передача начального пути для отображения у каждого файла
            pathAssembly = _path;
            pathWorm = _path;
            pathGear = _path;

            //Отображение пути сохранения сборки
            folderAssembly.Text = pathAssembly + "\\" + nameAssembly.Text + ".sldasm";
            //Отображение имени файла сборки
            nameAssembly.Text = "WormGearAssembly";

            //Обработка отображения параметров файла червяка
            if (_isWorm == true)
            {
                //Если червяк строится, то отображается имя и путь сохранения файла
                nameWorm.Text = "GeneratedWorm";
                folderWorm.Text = pathWorm + "\\" + nameWorm.Text + ".sldprt";
            }
            else
            {
                //Если червяк не строится, то отключаются поля, связанные со сборкой и червяком
                //Удаление данных пути сохранения сборки
                folderAssembly.Text = null;

                //Удаление данных имени сборки и отключение поля
                nameAssembly.TextChanged -= nameAssembly_TextChanged;
                nameAssembly.Text = null;
                nameAssembly.IsEnabled = false;

                //Отключение поля имени файла червяка
                nameWorm.TextChanged -= nameWorm_TextChanged;
                nameWorm.IsEnabled = false;

                //Удаление данных пути сохранения файла червяка
                folderWorm.Text = null;

                //Отключение кнопок, для изменения пути сохранения файлов червяка и сборки
                browseFolderWorm.IsEnabled = false;
                browseFolderAssembly.IsEnabled = false;
            }
            //Обработка отображения параметров файла червячного колеса
            if (_isGear == true)
            {
                //Если червячное колесо строится, то отображается имя и путь сохранения файла
                nameGear.Text = "GeneratedGear";
                folderGear.Text = pathGear + "\\" + nameGear.Text + ".sldprt";
            }
            else
            {
                //Если червячное колесо не строится, то отключаются поля, связанные со сборкой и червячным колесом
                //Удаление данных пути сохранения сборки
                folderAssembly.Text = null;

                //Удаление данных имени сборки и отключение поля
                nameAssembly.TextChanged -= nameAssembly_TextChanged;
                nameAssembly.Text = null;
                nameAssembly.IsEnabled = false;

                //Отключение поля имени файла червячного колеса
                nameGear.TextChanged -= nameGear_TextChanged;
                nameGear.IsEnabled = false;

                //Удаление данных пути сохранения файла червячного колеса
                folderGear.Text = null;

                //Отключение кнопок, для изменения пути сохранения файлов червячного колеса и сборки
                browseFolderGear.IsEnabled = false;
                browseFolderAssembly.IsEnabled = false;
            }        
            
        }

        private void nameAssembly_TextChanged(object sender, TextChangedEventArgs e)
        {
            folderAssembly.Text = pathAssembly + "\\" + nameAssembly.Text + ".sldasm";
        }

        private void nameWorm_TextChanged(object sender, TextChangedEventArgs e)
        {
            folderWorm.Text = pathWorm + "\\" + nameWorm.Text + ".sldprt";
        }

        private void nameGear_TextChanged(object sender, TextChangedEventArgs e)
        {
            folderGear.Text = pathGear + "\\" + nameGear.Text + ".sldprt";
        }

        private void browseFolderAssembly_Click(object sender, RoutedEventArgs e)
        {
            var folder = getFolder();
            if (folder != null)
                pathAssembly = folder;
            else
                return;
            folderAssembly.Text = pathAssembly + "\\" + nameAssembly.Text + ".sldasm";
        }

        private void browseFolderWorm_Click(object sender, RoutedEventArgs e)
        {
            var folder = getFolder();
            if (folder != null)
                pathWorm = folder;
            else
                return;
            folderWorm.Text = pathWorm + "\\" + nameWorm.Text + ".sldprt";
        }

        private void browseFolderGear_Click(object sender, RoutedEventArgs e)
        {
            var folder = getFolder();
            if (folder != null)
                pathGear = folder;
            else
                return;
            folderGear.Text = pathGear + "\\" + nameGear.Text + ".sldprt";
        }
        private string getFolder()
        {
            System.Windows.Forms.FolderBrowserDialog target = new System.Windows.Forms.FolderBrowserDialog();
            target.Description = "Выберите папку для сохранения файла";
            if (target.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return null;
            else
                return target.SelectedPath;
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.buildAccepted = false;
            this.Close();

        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            if (!validateFolders(folderAssembly.Text, nameAssembly.Text) ||
            !validateFolders(folderWorm.Text, nameWorm.Text) ||
            !validateFolders(folderGear.Text, nameGear.Text))
                return;

            MainWindow._pathAssembly = folderAssembly.Text;
            MainWindow._pathWorm = folderWorm.Text;
            MainWindow._pathGear = folderGear.Text;

            MainWindow._nameAssembly = nameAssembly.Text;
            MainWindow._nameWorm = nameWorm.Text;
            MainWindow._nameGear = nameGear.Text;

            MainWindow.buildAccepted = true;
            this.Close();
        }

        private bool validateFolders(string path, string name)
        {
            if (File.Exists(path))
                if (MessageBox.Show("Файл с таким именем уже существует ("+name+"). Заменить файл в папке назначения?", "Подтвердите действие",
                    MessageBoxButton.OKCancel, MessageBoxImage.Question).ToString() == "OK")
                    File.Delete(path);
                else
                    return false;
            return true;
        }

    }
}
