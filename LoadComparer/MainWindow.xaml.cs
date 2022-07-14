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
        }
        public static string pathInput1;
        public static string pathInput2;
        public static string pathOutput;
        public static int readySave;

        private void input1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
            if (openDialog.ShowDialog() == true)
            {
                pathInput1 = openDialog.FileName;
            }
            if (pathInput1 != null)
            {
                textInput1.Text = "Файл загружен";
                textInput1.Foreground = new SolidColorBrush(Colors.Green);
            }
        }

        private void input2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
            if (openDialog.ShowDialog() == true)
            {
                pathInput2 = openDialog.FileName;
            }
            if (pathInput2 != null)
            {
                textInput2.Text = "Файл загружен";
                textInput2.Foreground = new SolidColorBrush(Colors.Green);
            }
        }

        private void compare_Click(object sender, RoutedEventArgs e)
        {
            readySave = 0;
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(pathInput1);
            Excel.Worksheet sheet = workbook.ActiveSheet;

            List<string> head = new List<string>();
            for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                head.Add(sheet.Cells[1, i].Value);

            Excel.Range selectColumn;
            Array array;
            List<string> uid1 = new List<string>();
            List<string> uid2 = new List<string>();
            List<string> load1 = new List<string>();
            List<string> load2 = new List<string>();


            int numberOfUid;
            int numberOfLoad;
            if (head.FirstOrDefault(x => x.Contains("Uid")) != null && head.FirstOrDefault(x => x.Contains("Переток")) != null)
            {
                readySave++;
                numberOfUid = head.IndexOf(head.FirstOrDefault(x => x.Contains("Uid"))) + 1;
                numberOfLoad = head.IndexOf(head.FirstOrDefault(x => x.Contains("Переток"))) + 1;
                selectColumn = sheet.UsedRange.Columns[numberOfUid];
                array = (Array)selectColumn.Cells.Value2;
                uid1 = array.OfType<object>().Select(o => o.ToString()).ToList();
                selectColumn = sheet.UsedRange.Columns[numberOfLoad];
                array = (Array)selectColumn.Cells.Value2;
                load1 = array.OfType<object>().Select(o => o.ToString()).ToList();
            }
            else
                MessageBox.Show("В 1 файле отсутствует столбец с названием содержащим \"Uid\" или \"Переток\"");

            excel.Quit();
            workbook = excel.Workbooks.Open(pathInput2);
            sheet = workbook.ActiveSheet;
            head.Clear();
            for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                head.Add(sheet.Cells[1, i].Value);

            if (head.FirstOrDefault(x => x.Contains("Uid")) != null && head.FirstOrDefault(x => x.Contains("Переток")) != null)
            {
                readySave++;
                numberOfUid = head.IndexOf(head.FirstOrDefault(x => x.Contains("Uid"))) + 1;
                numberOfLoad = head.IndexOf(head.FirstOrDefault(x => x.Contains("Переток"))) + 1;
                selectColumn = sheet.UsedRange.Columns[numberOfUid];
                array = (Array)selectColumn.Cells.Value2;
                uid2 = array.OfType<object>().Select(o => o.ToString()).ToList();
                selectColumn = sheet.UsedRange.Columns[numberOfLoad];
                array = (Array)selectColumn.Cells.Value2;
                load2 = array.OfType<object>().Select(o => o.ToString()).ToList();
                excel.Quit();
            }
            else
                MessageBox.Show("Во 2 файле отсутствует столбец с названием содержащим \"Uid\" или \"Переток\"");
            
            if(readySave == 2)
            {
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
                saveFileDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
                if (saveFileDialog.ShowDialog() == true)
                {
                    pathOutput = saveFileDialog.FileName;
                    excel = new Excel.Application();
                    workbook = excel.Workbooks.Add(Type.Missing);
                    sheet = workbook.ActiveSheet;
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
                        textOutput.Text = "Сохранение успешно";
                    }
                    catch
                    {
                        MessageBox.Show("Нет доступа для записи в файл.");
                    }
                }
                excel.Quit();
            }
        }
    }
}
