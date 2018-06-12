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

namespace DogsViewModel
{
    /// <summary>
    /// Логика взаимодействия для Marks.xaml
    /// </summary>
    public partial class Marks : Window
    {
        public Marks()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            NewMarkWindow newMarkWindow = new NewMarkWindow();
            newMarkWindow.ShowDialog();
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
