using System;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using MessageBox = System.Windows.MessageBox;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelComparison
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool IsCheck = false;
        public MainWindow()
        {
            InitializeComponent();
        }

        public bool ReadExcel(string path)
        {
            try
            {
                //check exception handling is working or not.
                //throw new Exception();

                // define sheet name
                string firstSheet = "TRUE";
                string toBeCheckedSheet = "ToBeCheck";

                FileInfo fileInfo = new FileInfo(path);

                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorksheet trueSheet = package.Workbook.Worksheets[firstSheet];
                ExcelWorksheet toBeCheckSheet = package.Workbook.Worksheets[toBeCheckedSheet];
                
                // get number of rows and columns in the sheet
                int rows = trueSheet.Dimension.Rows;
                int columns = trueSheet.Dimension.Columns;

                for (int i = 2; i <= rows; i++)
                {
                    for (int j = 2; j <= columns; j++)
                    {
                        //checked for empty cell in ToBeCheck sheet
                        if (trueSheet.Cells[i, j].Value != null && toBeCheckSheet.Cells[i, j].Value == null)
                        {
                            toBeCheckSheet.Cells[i, j].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            toBeCheckSheet.Cells[i, j].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                        }
                        else
                        {
                            //checked for invalid value in ToBeCheck sheet

                            var trueValue = trueSheet.Cells[i, j].Value ?? string.Empty;
                            var toBeCheckValue = toBeCheckSheet.Cells[i, j].Value ?? string.Empty;
                            if (!trueValue.Equals(toBeCheckValue))
                            {
                                toBeCheckSheet.Cells[i, j].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                toBeCheckSheet.Cells[i, j].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }
                        }
                    }
                }

                package.Save();
                IsCheck = true;
            }
            catch (Exception ex)
            {
                IsCheck = false;
                throw;
            }
            return IsCheck;
        }

        public void ExecuteExcelMacro(string path, string macroName)
        {
            try
            {
                Excel.ApplicationClass oExcel = new Excel.ApplicationClass();
                oExcel.Visible = true;
                Excel.Workbooks oBooks = oExcel.Workbooks;
                Excel._Workbook oBook = null;
                oBook = oBooks.Open(path);

                // Run the macros.
                RunMacro(oExcel, new Object[] { macroName });

                // Quit Excel and clean up.
                oBook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                oBook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                oBooks = null;
                oExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                oExcel = null;
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }

        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            string path = txtExcelPath.Text.Trim().ToString();

            try
            {
                if (!string.IsNullOrEmpty(path))
                {
                    if (ReadExcel(path))
                    {
                        MessageBox.Show("Successfully Comparision", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Something worng in comparision", "Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Please type Excel path..", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtExcelPath.Focus();
                }
            }
            catch (Exception ex)
            {
                string str = ex.ToString();
                Console.WriteLine(str);
            }
            
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = txtExcelPath.Text.Trim().ToString();
                string macroName = txtMacroName.Text.Trim().ToString();
                if (string.IsNullOrEmpty(path))
                {
                    MessageBox.Show("Please type Excel path in Excel Path Text Box.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtExcelPath.Focus();
                }
                else if (string.IsNullOrEmpty(macroName))
                {
                    MessageBox.Show("Please type Macro Name.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtMacroName.Focus();
                }
                else
                {
                    ExecuteExcelMacro(path,macroName);
                }
            }
            catch (Exception ex)
            {
                string str = ex.ToString();
                Console.WriteLine(str);
            }
            
        }
    }

}
