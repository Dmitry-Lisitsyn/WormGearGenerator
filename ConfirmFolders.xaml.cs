using System;
using System.Collections.Generic;
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
        public string _path { get; set; }

        public string pathAssembly { get; set; }
        public string pathWorm { get; set; }
        public string pathGear { get; set; }

        public string _nameAssembly { get; set; }
        public string _nameWorm { get; set; }
        public string _nameGear { get; set; }
        
        private Window mainWindow;

        public ConfirmFolders(Window mainWindow , string path)
        {
            this.mainWindow = mainWindow;
            _path = path;

            InitializeComponent();
            InitializeFolders();
        }

        public void InitializeFolders()
        {
            nameAssembly.Text = "WormGearAssembly";
            nameWorm.Text = "GeneratedWorm";
            nameGear.Text = "GeneratedGear";
            pathAssembly = _path;
            pathWorm = _path;
            pathGear = _path;

            folderAssembly.Text = pathAssembly + "\\" + nameAssembly.Text + ".sldasm";
            folderWorm.Text = pathWorm + "\\" + nameWorm.Text + ".sldprt";
            folderGear.Text = pathGear + "\\" + nameGear.Text + ".sldprt";

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
            System.Windows.Forms.FolderBrowserDialog target = new System.Windows.Forms.FolderBrowserDialog();
            target.Description = "Выберите папку для сохранения файла";
            if (target.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            else
                pathAssembly = target.SelectedPath;
            folderAssembly.Text = pathAssembly + "\\" + nameAssembly.Text + ".sldasm";
        }

        private void browseFolderWorm_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog target = new System.Windows.Forms.FolderBrowserDialog();
            target.Description = "Выберите папку для сохранения файла";
            if (target.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            else
                pathWorm = target.SelectedPath;
            folderWorm.Text = pathWorm + "\\" + nameWorm.Text + ".sldprt";
        }

        private void browseFolderGear_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog target = new System.Windows.Forms.FolderBrowserDialog();
            target.Description = "Выберите папку для сохранения файла";
            if (target.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            else
                pathGear = target.SelectedPath;
            folderGear.Text = pathGear + "\\" + nameGear.Text + ".sldprt";
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.buildAccepted = false;
            this.Close();

        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            MainWindow._pathAssembly = folderAssembly.Text;
            MainWindow._pathWorm = folderWorm.Text;
            MainWindow._pathGear = folderGear.Text;

            MainWindow._nameAssembly = nameAssembly.Text;
            MainWindow._nameWorm = nameWorm.Text;
            MainWindow._nameGear = nameGear.Text;

            MainWindow.buildAccepted = true;
            this.Close();
        }


    }
}
