using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
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

            dgUsers.ItemsSource = users;

            rezultParam1.Content = string.Concat("SFF = ", " (Вт/Гц)");
            rezultParam2.Content = string.Concat("F = ", " (Дж)");
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Execl files (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            //saveFileDialog.CreatePrompt = true;
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

                    int i;
                    for (i = 1; i < dgUsers.Items.Count; i++)
                    {
                        workSheet.Cells[i, 1] = ((Param)dgUsers.Items.GetItemAt(i)).Name;
                        workSheet.Cells[i, 2] = ((Param)dgUsers.Items.GetItemAt(i)).Value;
                        workSheet.Cells[i, 3] = ((Param)dgUsers.Items.GetItemAt(i)).Rozmir;
                    }
                    i++;
                    workSheet.Cells[i, 1] = "SFF";
                    workSheet.Cells[i, 2] = _param1;
                    workSheet.Cells[i, 3] = "Вт/Гц";

                    i++;
                    workSheet.Cells[i, 1] = "F";
                    workSheet.Cells[i, 2] = _param2;
                    workSheet.Cells[i, 3] = "Н(м/с^2)^3";

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
            double obMass = ((Param)dgUsers.Items.GetItemAt(0)).Value;
            double obFrequn = ((Param)dgUsers.Items.GetItemAt(1)).Value;
            double cavResFrequn = ((Param)dgUsers.Items.GetItemAt(2)).Value;
            double prTransFreq = ((Param)dgUsers.Items.GetItemAt(3)).Value;
            double photonNum = ((Param)dgUsers.Items.GetItemAt(4)).Value;
            double measurTime = ((Param)dgUsers.Items.GetItemAt(5)).Value;
            double phaseNoise = ((Param)dgUsers.Items.GetItemAt(6)).Value;
            double lightSpeed = ((Param)dgUsers.Items.GetItemAt(7)).Value;
            double photonNumFluct = ((Param)dgUsers.Items.GetItemAt(8)).Value;
            double measurNum = ((Param)dgUsers.Items.GetItemAt(9)).Value;

            _param1 = Calculation.spectrum(obMass, obFrequn, cavResFrequn, photonNum,
                phaseNoise, prTransFreq, lightSpeed, measurTime, measurNum) * 10e-14;
            _param2 = Calculation.spectrum(obMass, obFrequn, cavResFrequn, photonNum,
                phaseNoise, prTransFreq, lightSpeed, measurTime, measurNum, photonNumFluct) * 10e-16;

            rezultParam1.Content = string.Concat("SFF = ", _param1.ToString(), " (Вт/Гц)");
            rezultParam2.Content = string.Concat("F = ", _param2.ToString(), " (Н(м/с^2)^3)");
        }
    }
}