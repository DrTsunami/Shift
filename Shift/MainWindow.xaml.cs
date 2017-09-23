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
using System.Runtime.InteropServices;

namespace Shift
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        ////////////////////////////////////////////////////////////////////////////////////////
        // vars
        ////////////////////////////////////////////////////////////////////////////////////////
        
        int sheetVerifyCol = 1;
        int sheetVerifyRow = 40;


        ////////////////////////////////////////////////////////////////////////////////////////
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            // create open file dialog
            Microsoft.Win32.OpenFileDialog fd = new Microsoft.Win32.OpenFileDialog();

            // set filter for file extension
            fd.Filter = "Excel Files (*.xlsx)| *.xlsx";

            // Display fd 
            Nullable<bool> result = fd.ShowDialog();

            // Get selected file name
            if (result == true)
            {
                if (VerifyFile(fd.FileName.ToString()))
                {
                    App.Start(fd.FileName.ToString());
                } else
                {
                    Console.WriteLine("ERROR: file not verified. Please verify file");
                    MessageBox.Show("File not verified, please verify file by inputting '[VERIFIED]' in the cell [1, 40]");
                }
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

        private bool VerifyFile(String path)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            String verifyCell = (xlWorksheet.Cells[sheetVerifyRow, sheetVerifyCol] as
                Microsoft.Office.Interop.Excel.Range).Value;

            if (verifyCell != null)
            {
                if (verifyCell.Equals("[VERIFIED]"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            } else
            {
                return false;
            }
            
        }

        private void XlCleanup(Microsoft.Office.Interop.Excel.Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook,
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet,
            Microsoft.Office.Interop.Excel.Range xlRange)
        {
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("objects released");
        }
    }
}
