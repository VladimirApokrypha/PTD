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
using FuncionalPTD.BusinessLogic;

namespace DogsViewModel
{
    /// <summary>
    /// Логика взаимодействия для AddContractorFileWindow.xaml
    /// </summary>
    public partial class AddContractorFileWindow : Window
    {
        public AddContractorFileWindow()
        {
            InitializeComponent();
            DataContext = new AddContractorFileBL();
        }
        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog SelectContractorFilePath = new OpenFileDialog();
            SelectContractorFilePath.ShowDialog();
            TextBox1.Text = SelectContractorFilePath.FileName;
        }
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TextBox2_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TextBox2.Clear();
        }
    }
}
