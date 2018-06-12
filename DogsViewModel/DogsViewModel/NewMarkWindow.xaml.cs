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
    /// Логика взаимодействия для NewMarkWindow.xaml
    /// </summary>
    public partial class NewMarkWindow : Window
    {
        public NewMarkWindow()
        {
            InitializeComponent();
        }
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog SelectCellPath = new OpenFileDialog();
            SelectCellPath.ShowDialog();
        }
        private void TextBox2_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            TextBox2.Clear();
        }
        private void TextBox3_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            TextBox3.Clear();
        }

        private void TextBox1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
