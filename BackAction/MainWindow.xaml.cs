using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
namespace BackAction
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private double _param1 = 0;
        private double _param2 = 0;

        public MainWindow()
        {
            InitializeComponent();

            var users = new List<Param>();
            users.Add(new Param() { Name = "M", Rozmir = "*10^-27 кг" });
            users.Add(new Param() { Name = "\x03C5", Rozmir = "*10^9 Гц" }); 
            users.Add(new Param() { Name = "\x03C9с", Rozmir = "*10^9 Гц" });
            users.Add(new Param() { Name = "fq", Rozmir = "*10^9 Гц" });
            users.Add(new Param() { Name = "n", Rozmir = "" });
            users.Add(new Param() { Name = "Tm", Rozmir = "*10^-9 c" });
            users.Add(new Param() { Name = "\x03B4(\x03C9c - \x03C9q)", Rozmir = "" });
            users.Add(new Param() { Name = "c", Rozmir = "*10^8 м/с" }); 
            users.Add(new Param() { Name = "\x03B4n", Rozmir = "" });
            users.Add(new Param() { Name = "N", Rozmir = "" });

            paramsGrid.ItemsSource = users;

            rezultParam1.Content = string.Concat("SFF = ", " (Вт/Гц)");
            rezultParam2.Content = string.Concat("F = ", " (Дж)");
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Export Excel File To";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    Excel.Application excelApp = new Excel.Application();
                    if (excelApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }

                    var workBook = excelApp.Workbooks.Add();
                    var workSheet = workBook.Worksheets.get_Item(1);

                    int rowWorkSheet;
                    for (rowWorkSheet = 1; rowWorkSheet <= paramsGrid.Items.Count; rowWorkSheet++)
                    {
                        workSheet.Cells[rowWorkSheet, 1] = ((Param)paramsGrid.Items.GetItemAt(rowWorkSheet-1)).Name;
                        workSheet.Cells[rowWorkSheet, 2] = ((Param)paramsGrid.Items.GetItemAt(rowWorkSheet-1)).Value;
                        workSheet.Cells[rowWorkSheet, 3] = ((Param)paramsGrid.Items.GetItemAt(rowWorkSheet-1)).Rozmir;
                    }
                    rowWorkSheet++;
                    workSheet.Cells[rowWorkSheet, 1] = "SFF";
                    workSheet.Cells[rowWorkSheet, 2] = _param1;
                    workSheet.Cells[rowWorkSheet, 3] = "Вт/Гц";

                    rowWorkSheet++;
                    workSheet.Cells[rowWorkSheet, 1] = "F";
                    workSheet.Cells[rowWorkSheet, 2] = _param2;
                    workSheet.Cells[rowWorkSheet, 3] = "Дж";

                    workBook.SaveAs(saveFileDialog.FileName);
                    workBook.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void makeClaculations(object sender, RoutedEventArgs e)
        { 
            double obMass = ((Param)paramsGrid.Items.GetItemAt(0)).Value;
            double obFrequn = ((Param)paramsGrid.Items.GetItemAt(1)).Value;
            double cavResFrequn = ((Param)paramsGrid.Items.GetItemAt(2)).Value;
            double prTransFreq = ((Param)paramsGrid.Items.GetItemAt(3)).Value;
            double photonNum = ((Param)paramsGrid.Items.GetItemAt(4)).Value;
            double measurTime = ((Param)paramsGrid.Items.GetItemAt(5)).Value;
            double phaseNoise = ((Param)paramsGrid.Items.GetItemAt(6)).Value;
            double lightSpeed = ((Param)paramsGrid.Items.GetItemAt(7)).Value;
            double photonNumFluct = ((Param)paramsGrid.Items.GetItemAt(8)).Value;
            double measurNum = ((Param)paramsGrid.Items.GetItemAt(9)).Value;

            _param1 = Calculation.spectrum(obMass, obFrequn, cavResFrequn, photonNum,
                phaseNoise, prTransFreq, lightSpeed, measurTime, measurNum) * 10e-14;
            _param2 = Calculation.spectrum(obMass, obFrequn, cavResFrequn, photonNum,
                phaseNoise, prTransFreq, lightSpeed, measurTime, measurNum, photonNumFluct) * 10e-16;

            rezultParam1.Content = string.Concat("SFF = ", _param1.ToString(), " (Вт/Гц)");
            rezultParam2.Content = string.Concat("F = ", _param2.ToString(), " (Дж)");
        }
    }
}