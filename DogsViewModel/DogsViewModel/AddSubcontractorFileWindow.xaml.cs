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
    /// Логика взаимодействия для AddSubcontractorFileWindow.xaml
    /// </summary>
    public partial class AddSubcontractorFileWindow : Window
    {
        public AddSubcontractorFileWindow()
        {
            InitializeComponent();
        }
        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog SelectSubcontractorFilePath = new OpenFileDialog();
            SelectSubcontractorFilePath.ShowDialog();
        }
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
