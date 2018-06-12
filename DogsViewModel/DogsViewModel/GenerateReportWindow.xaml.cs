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
using Microsoft.Win32;

namespace DogsViewModel
{
    /// <summary>
    /// Логика взаимодействия для GenerateReportWindow.xaml
    /// </summary>
    public partial class GenerateReportWindow : Window
    {
        public GenerateReportWindow()
        {
            InitializeComponent();
        }
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog SelectGenerateReportPath = new OpenFileDialog();
            SelectGenerateReportPath.ShowDialog();
        }
        private void TextBox2_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            TextBox2.Clear();
        }
    }
}
