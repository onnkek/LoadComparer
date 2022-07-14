using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoadComparer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Title = "XLSX Comparer";
            textOutput.Visibility = Visibility.Hidden;
        }
        public static string pathInput1;
        public static string pathInput2;
        public static string pathOutput;
        public static string nameFile1;
        public static string nameFile2;
        public static Excel.Application excel = new Excel.Application();
        public static Excel.Workbook book;
        public static Excel.Worksheet sheet;

        private void input1_Click(object sender, RoutedEventArgs e)
        {
            selectUid1.Items.Clear();
            selectLoad1.Items.Clear();
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv|All files|*.*";
            if (openDialog.ShowDialog() == true)
            {
                pathInput1 = openDialog.FileName;
                nameFile1 = openDialog.SafeFileName;
            }
            if (pathInput1 != null)
            {
                if (pathInput1.Substring(pathInput1.Count() - 5).Contains(".csv"))
                    book = excel.Workbooks.OpenXML(pathInput1);
                else
                    book = excel.Workbooks.Open(pathInput1);
                sheet = book.ActiveSheet;
                List<string> head = new List<string>();
                for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                {
                    selectUid1.Items.Add(sheet.Cells[1, i].Value);
                    selectLoad1.Items.Add(sheet.Cells[1, i].Value);
                    if (sheet.Cells[1, i].Value.Contains("Uid"))
                        selectUid1.SelectedIndex = i - 1;
                    if (sheet.Cells[1, i].Value.Contains("Переток"))
                        selectLoad1.SelectedIndex = i - 1;
                }
                excel.Quit();
            }
            textOutput.Visibility = Visibility.Hidden;
        }

        private void input2_Click(object sender, RoutedEventArgs e)
        {
            selectUid2.Items.Clear();
            selectLoad2.Items.Clear();
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv|All files|*.*";
            if (openDialog.ShowDialog() == true)
            {
                pathInput2 = openDialog.FileName;
                nameFile2 = openDialog.SafeFileName;
            }
            if (pathInput2 != null)
            {
                if (pathInput2.Substring(pathInput2.Count() - 5).Contains(".csv"))
                    book = excel.Workbooks.OpenXML(pathInput2);
                else
                    book = excel.Workbooks.Open(pathInput2);
                sheet = book.ActiveSheet;
                List<string> head = new List<string>();
                for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                {
                    selectUid2.Items.Add(sheet.Cells[1, i].Value);
                    selectLoad2.Items.Add(sheet.Cells[1, i].Value);
                    if (sheet.Cells[1, i].Value.Contains("Uid"))
                        selectUid2.SelectedIndex = i - 1;
                    if (sheet.Cells[1, i].Value.Contains("Переток"))
                        selectLoad2.SelectedIndex = i - 1;
                }
                excel.Quit();
            }
            textOutput.Visibility = Visibility.Hidden;
        }

        private void compare_Click(object sender, RoutedEventArgs e)
        {
            if (pathInput1.Substring(pathInput1.Count() - 5).Contains(".csv"))
                book = excel.Workbooks.OpenXML(pathInput1);
            else
                book = excel.Workbooks.Open(pathInput1);
            sheet = book.ActiveSheet;

            Excel.Range selectColumn;
            Array array;
            List<string> uid1 = new List<string>();
            List<string> uid2 = new List<string>();
            List<string> load1 = new List<string>();
            List<string> load2 = new List<string>();


            int numberOfUid;
            int numberOfLoad;

            numberOfUid = selectUid1.SelectedIndex + 1;
            numberOfLoad = selectLoad1.SelectedIndex + 1;
            selectColumn = sheet.UsedRange.Columns[numberOfUid];
            array = (Array)selectColumn.Cells.Value2;
            uid1 = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfLoad];
            array = (Array)selectColumn.Cells.Value2;
            load1 = array.OfType<object>().Select(o => o.ToString()).ToList();

            excel.Quit();
            if (pathInput2.Substring(pathInput2.Count() - 5).Contains(".csv"))
                book = excel.Workbooks.OpenXML(pathInput2);
            else
                book = excel.Workbooks.Open(pathInput2);
            
            sheet = book.ActiveSheet;

            numberOfUid = selectUid2.SelectedIndex + 1;
            numberOfLoad = selectLoad2.SelectedIndex + 1;
            selectColumn = sheet.UsedRange.Columns[numberOfUid];
            array = (Array)selectColumn.Cells.Value2;
            uid2 = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfLoad];
            array = (Array)selectColumn.Cells.Value2;
            load2 = array.OfType<object>().Select(o => o.ToString()).ToList();
            excel.Quit();



            List<string> exceptUid1 = uid1.Except(uid2).ToList();
            List<string> exceptUid2 = uid2.Except(uid1).ToList();
            List<string> intersectUid = uid1.Intersect(uid2).ToList();
            List<string> notEqualLoad = new List<string>();
            foreach (var u in intersectUid)
            {
                var res = uid2.FirstOrDefault(x => x == u);
                if (res != null)
                {
                    var ind1 = uid1.IndexOf(res);
                    var ind2 = uid2.IndexOf(res);
                    if (load1[ind1] != load2[ind2])
                        notEqualLoad.Add(u);
                }
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = $"Сравнение {nameFile1} и {nameFile2}.xlsx";
            saveFileDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                pathOutput = saveFileDialog.FileName;
                book = excel.Workbooks.Add(Type.Missing);
                sheet = book.ActiveSheet;
                var range = sheet.get_Range("A1", "C1");
                range.ColumnWidth = 35;
                range.Interior.Color = Excel.XlRgbColor.rgbLightSkyBlue;
                range.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                sheet.Cells[1, 1].Value = "Uid: нет во втором файле";
                for (int i = 0; i < exceptUid1.Count; i++)
                {
                    sheet.Cells[i + 2, 1].Value = exceptUid1[i];
                }
                sheet.Cells[1, 2].Value = "Uid: нет в первом файле";
                for (int i = 0; i < exceptUid2.Count; i++)
                {
                    sheet.Cells[i + 2, 2].Value = exceptUid2[i];
                }
                sheet.Cells[1, 3].Value = "Uid: не совпадают значения";
                for (int i = 0; i < notEqualLoad.Count; i++)
                {
                    sheet.Cells[i + 2, 3].Value = notEqualLoad[i];
                }
                try
                {
                    excel.Application.ActiveWorkbook.SaveAs(pathOutput);
                }
                catch
                {
                    MessageBox.Show("Нет доступа для записи в файл.");
                }
                textOutput.Visibility = Visibility.Visible;
            }
            excel.Quit();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            excel.Quit();
        }
    }
}
