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
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Office.Interop.Excel;

namespace Shift
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            // create open file dialog
            Microsoft.Win32.OpenFileDialog fd = new Microsoft.Win32.OpenFileDialog();

            // set filter for file extension
            fd.DefaultExt = ".xlsx";

            // Display fd 
            Nullable<bool> result = fd.ShowDialog();

            // Get selected file name
            if (result == true)
            {
                App.Start(fd.FileName);
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            SheetProcessor sp = new SheetProcessor();

            // create open file dialog
            Microsoft.Win32.OpenFileDialog fd = new Microsoft.Win32.OpenFileDialog();

            // set filter for file extension
            fd.DefaultExt = ".xlsx";

            // Display fd 
            Nullable<bool> result = fd.ShowDialog();

            // Get selected file name
            if (result == true)
            {
                sp.Reformat(fd.FileName);
            }
        }
    }
}
