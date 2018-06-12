using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using FuncionalPTD.BusinessLogic;

namespace DogsViewModel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        public MainWindow()
        {
            InitializeComponent();
        }
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void NewProjectButton_OnClick(object sender, RoutedEventArgs e)
        {
            NewProjectWindow newProjectWindow = new NewProjectWindow();
            newProjectWindow.ShowDialog();
        }
        private void NewDirectoryButton_OnClick(object sender, RoutedEventArgs e)
        {
           NewDirectoryWindow newDirectoryWindow = new NewDirectoryWindow();
           newDirectoryWindow.ShowDialog();
        }
        private void OpenProjectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenProjectWindow openProjectWindow = new OpenProjectWindow();
            openProjectWindow.ShowDialog();
        }
        private void OpenDirectoryButton_Click(object sender, RoutedEventArgs e)
        {
            OpenDirectoryWindow openDirectoryWindow = new OpenDirectoryWindow();
            openDirectoryWindow.ShowDialog();
        }
        private void ProjectTreeButton_Click(object sender, RoutedEventArgs e)
        {
            if (TreeOfProject.Visibility == Visibility.Collapsed)
            {
                TreeOfProject.Visibility = Visibility.Visible;
                Button1Vis.Visibility = Visibility.Visible;
                Button2Vis.Visibility = Visibility.Visible;
                Button3Vis.Visibility = Visibility.Visible;
                Button1.Visibility = Visibility.Collapsed;
                Button2.Visibility = Visibility.Collapsed;
                Button3.Visibility = Visibility.Collapsed;
            }
            else
            {
                TreeOfProject.Visibility = Visibility.Collapsed;
                Button1Vis.Visibility = Visibility.Collapsed;
                Button2Vis.Visibility = Visibility.Collapsed;
                Button3Vis.Visibility = Visibility.Collapsed;
                Button1.Visibility = Visibility.Visible;
                Button2.Visibility = Visibility.Visible;
                Button3.Visibility = Visibility.Visible;
            }
        }
        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            AddSubcontractorFileWindow addSubcontractorFileWindow = new AddSubcontractorFileWindow();
            addSubcontractorFileWindow.ShowDialog();
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            AddContractorFileWindow addContractorFileWindow = new AddContractorFileWindow();
            addContractorFileWindow.ShowDialog();
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            GenerateReportWindow generateReportWindow = new GenerateReportWindow();
            generateReportWindow.ShowDialog();
        }

        private void Marks_Click(object sender, RoutedEventArgs e)
        {
            Marks marks = new Marks();
            marks.Show();
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow helpWindow = new HelpWindow();
            helpWindow.ShowDialog();
        }
    }
}
