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
            excel.Quit();
            List<Row> rows1 = new List<Row>();
            List<Row> rows2 = new List<Row>();

            if (pathInput1.Substring(pathInput1.Count() - 5).Contains(".csv"))
                book = excel.Workbooks.OpenXML(pathInput1);
            else
                book = excel.Workbooks.Open(pathInput1);
            sheet = book.ActiveSheet;

            Excel.Range selectColumn;
            Array array;
            List<string> uid = new List<string>();
            List<string> load = new List<string>();
            List<string> name = new List<string>();
            int numberOfUid = selectUid1.SelectedIndex + 1;
            int numberOfLoad = selectLoad1.SelectedIndex + 1;
            int numberOfName = selectLoad1.SelectedIndex;

            selectColumn = sheet.UsedRange.Columns[numberOfUid];
            array = (Array)selectColumn.Cells.Value2;
            uid = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfLoad];
            array = (Array)selectColumn.Cells.Value2;
            load = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfName];
            array = (Array)selectColumn.Cells.Value2;
            name = array.OfType<object>().Select(o => o.ToString()).ToList();
            excel.Quit();

            for (int i = 1; i < uid.Count; i++)
                rows1.Add(new Row(name[i], uid[i], Convert.ToDouble(load[i])));


            if (pathInput2.Substring(pathInput2.Count() - 5).Contains(".csv"))
                book = excel.Workbooks.OpenXML(pathInput2);
            else
                book = excel.Workbooks.Open(pathInput2);

            sheet = book.ActiveSheet;

            numberOfUid = selectUid2.SelectedIndex + 1;
            numberOfLoad = selectLoad2.SelectedIndex + 1;
            numberOfName = selectLoad2.SelectedIndex;

            selectColumn = sheet.UsedRange.Columns[numberOfUid];
            array = (Array)selectColumn.Cells.Value2;
            uid = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfLoad];
            array = (Array)selectColumn.Cells.Value2;
            load = array.OfType<object>().Select(o => o.ToString()).ToList();
            selectColumn = sheet.UsedRange.Columns[numberOfName];
            array = (Array)selectColumn.Cells.Value2;
            name = array.OfType<object>().Select(o => o.ToString()).ToList();
            excel.Quit();

            for (int i = 1; i < uid.Count; i++)
                rows2.Add(new Row(name[i], uid[i], Convert.ToDouble(load[i])));

            List<Row> exceptUid1 = rows1.Where(x => !rows2.Any(y => y.Uid.Equals(x.Uid))).ToList();
            List<Row> exceptUid2 = rows2.Where(x => !rows1.Any(y => y.Uid.Equals(x.Uid))).ToList();
            List<Row> notEqualLoad = new List<Row>();
            List<Row> intersectUid = rows1.Where(x => rows2.Any(y => y.Uid.Equals(x.Uid))).ToList();

            foreach (var r in intersectUid)
            {
                var res1 = rows1.FirstOrDefault(x => x.Uid == r.Uid);
                var res2 = rows2.FirstOrDefault(x => x.Uid == r.Uid);
                if (rows1[rows1.IndexOf(res1)].Value != rows2[rows2.IndexOf(res2)].Value)
                    notEqualLoad.Add(r);
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = $"Сравнение {nameFile1} и {nameFile2}.xlsx";
            saveFileDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                pathOutput = saveFileDialog.FileName;
                book = excel.Workbooks.Add(Type.Missing);
                sheet = book.ActiveSheet;
                var range = sheet.get_Range("A1", "J2");
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.Interior.Color = Excel.XlRgbColor.rgbLightSkyBlue;
                range.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.get_Range("A1", "C1").Merge();
                sheet.get_Range("D1", "F1").Merge();
                sheet.get_Range("G1", "J1").Merge();
                sheet.Cells[1, 1].Value = "Нет во втором файле";
                sheet.Cells[2, 1].Value = "Uid";
                sheet.Cells[2, 2].Value = "Name";
                sheet.Cells[2, 3].Value = "Value";
                for (int i = 0; i < exceptUid1.Count; i++)
                {
                    sheet.Cells[i + 3, 1].Value = exceptUid1[i].Uid;
                    sheet.Cells[i + 3, 2].Value = exceptUid1[i].Name;
                    sheet.Cells[i + 3, 3].Value = exceptUid1[i].Value;
                }
                sheet.Cells[1, 4].Value = "Нет в первом файле";
                sheet.Cells[2, 4].Value = "Uid";
                sheet.Cells[2, 5].Value = "Name";
                sheet.Cells[2, 6].Value = "Value";
                for (int i = 0; i < exceptUid2.Count; i++)
                {
                    sheet.Cells[i + 3, 4].Value = exceptUid2[i].Uid;
                    sheet.Cells[i + 3, 5].Value = exceptUid2[i].Name;
                    sheet.Cells[i + 3, 6].Value = exceptUid2[i].Value;
                }
                
                
                sheet.Cells[1, 7].Value = "Не совпадают значения";
                sheet.Cells[2, 7].Value = "Uid";
                sheet.Cells[2, 8].Value = "Name";
                sheet.Cells[2, 9].Value = "Value1";
                sheet.Cells[2, 10].Value = "Value2";
                for (int i = 0; i < notEqualLoad.Count; i++)
                {
                    sheet.Cells[i + 3, 7].Value = exceptUid2[i].Uid;
                    sheet.Cells[i + 3, 8].Value = exceptUid2[i].Name;
                    sheet.Cells[i + 3, 9].Value = rows1.FirstOrDefault(x => x.Uid == notEqualLoad[i].Uid).Value;
                    sheet.Cells[i + 3, 10].Value = rows2.FirstOrDefault(x => x.Uid == notEqualLoad[i].Uid).Value;


                }
                sheet.Columns.AutoFit();
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
        public class Row
        {
            public Row(string name, string uid, double value)
            {
                Name = name;
                Uid = uid;
                Value = value;
            }
            public string Name { get; set; }
            public double Value { get; set; }
            public string Uid { get; set; }
        }
    }
}
