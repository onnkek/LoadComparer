﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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
            textOutput.Visibility = Visibility.Hidden;
        }
        public delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);
        public static string pathInput1;
        public static string pathInput2;
        public static string pathOutput;
        public static string nameFile1;
        public static string nameFile2;
        public static Excel.Application excel = new Excel.Application();
        public static Excel.Workbook book;
        public static Excel.Worksheet sheet;
        public static int rowCount1;
        public static int rowCount2;

        private void input1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                textOutput.Visibility = Visibility.Hidden;
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
                    rowCount1 = sheet.Rows.CurrentRegion.EntireRow.Count;
                    List<string> head = new List<string>();
                    for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                    {
                        selectUid1.Items.Add(sheet.Cells[1, i].Value);
                        selectLoad1.Items.Add(sheet.Cells[1, i].Value);
                        selectColumn1.Children.Add(new CheckBox { Content = sheet.Cells[1, i].Value, LayoutTransform = new ScaleTransform(1.5, 1.5), FontFamily = new FontFamily("Calibri") });
                        // Автоматическая установка флажка Uid (если есть)
                        foreach (var cb in selectColumn1.Children)
                            if (cb is CheckBox)
                                if (((CheckBox)cb).Content.ToString().Contains("Uid"))
                                    ((CheckBox)cb).IsChecked = true;
                        // Автоматический выбор Uid и Переток (если есть)
                        if (sheet.Cells[1, i].Value.ToString().Contains("Uid"))
                            selectUid1.SelectedIndex = i - 1;
                        if (sheet.Cells[1, i].Value.ToString().Contains("Переток"))
                            selectLoad1.SelectedIndex = i - 1;
                    }
                    excel.Quit();
                }
            }
            catch
            {
                excel.Quit();
                selectUid1.Items.Clear();
                selectLoad1.Items.Clear();
                selectColumn1.Children.Clear();
                MessageBox.Show("Входной файл имел неверный формат или недопустымые названиия столбцов");
            }
            finally
            {
                excel.Quit();
            }
        }

        private void input2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                textOutput.Visibility = Visibility.Hidden;
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
                    rowCount2 = sheet.Rows.CurrentRegion.EntireRow.Count;
                    List<string> head = new List<string>();
                    for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                    {
                        selectUid2.Items.Add(sheet.Cells[1, i].Value);
                        selectLoad2.Items.Add(sheet.Cells[1, i].Value);
                        selectColumn2.Children.Add(new CheckBox { Content = sheet.Cells[1, i].Value, LayoutTransform = new ScaleTransform(1.5, 1.5), FontFamily = new FontFamily("Calibri") });
                        // Автоматическая установка флажка Uid (если есть)
                        foreach (var cb in selectColumn2.Children)
                            if (cb is CheckBox)
                                if (((CheckBox)cb).Content.ToString().Contains("Uid"))
                                    ((CheckBox)cb).IsChecked = true;
                        // Автоматический выбор Uid и Переток (если есть)
                        if (sheet.Cells[1, i].Value.Contains("Uid"))
                            selectUid2.SelectedIndex = i - 1;
                        if (sheet.Cells[1, i].Value.Contains("Переток"))
                            selectLoad2.SelectedIndex = i - 1;
                    }
                    excel.Quit();
                }
            }
            catch
            {
                excel.Quit();
                selectUid2.Items.Clear();
                selectLoad2.Items.Clear();
                selectColumn2.Children.Clear();
                MessageBox.Show("Входной файл имел неверный формат или недопустымые названиия столбцов");
            }
            finally
            {
                excel.Quit();
            }
        }

        private void compare_Click(object sender, RoutedEventArgs e)
        {
            int errors = 0;
            try
            {
                if (pathInput1 == null)
                {
                    errors++;
                    MessageBox.Show("Первый файл не загружен!");
                }
                else if (pathInput2 == null)
                {
                    errors++;
                    MessageBox.Show("Второй файл не загружен!");
                }
                else if (selectLoad1.SelectedItem == null)
                {
                    errors++;
                    MessageBox.Show("Не выбрано сравниваемое значение для файла 1!");
                }
                else if (selectLoad2.SelectedItem == null)
                {
                    errors++;
                    MessageBox.Show("Не выбрано сравниваемое значение для файла 2!");
                }
                else if(selectColumn1.Children.Count == 0)
                {
                    errors++;
                    MessageBox.Show("Количество выводимых колонок для файла 1 должно быть больше 0!");
                }
                else if (selectColumn2.Children.Count == 0)
                {
                    errors++;
                    MessageBox.Show("Количество выводимых колонок для файла 2 должно быть больше 0!");
                }
                if(errors == 0)
                {
                    UpdateProgressBarDelegate updProgress = new UpdateProgressBarDelegate(progressBar.SetValue);
                    double pbValue = 0;
                    progressBar.Maximum = 7 + selectUid1.Items.Count * rowCount1 + selectUid2.Items.Count * rowCount2;
                    progressBar.Value = pbValue;
                    List<string> outColumn1 = new List<string>();
                    List<string> outColumn2 = new List<string>();

                    // Получение данных с формы для использования в новом потоке
                    List<string> selectUidItems1 = new List<string>();
                    List<string> selectUidItems2 = new List<string>();
                    List<string> selectUidSelectedItems1 = new List<string>();
                    List<string> selectUidSelectedItems2 = new List<string>();
                    List<string> selectLoadItems1 = new List<string>();
                    List<string> selectLoadItems2 = new List<string>();
                    int selectedLoad1Index = selectLoad1.Items.IndexOf(selectLoad1.SelectedItem);
                    int selectedLoad2Index = selectLoad2.Items.IndexOf(selectLoad2.SelectedItem);
                    List<string> selectLoadSelectedItems1 = new List<string>();
                    List<string> selectLoadSelectedItems2 = new List<string>();
                    List<CheckBoxModel> selectColumnItems1 = new List<CheckBoxModel>();
                    List<CheckBoxModel> selectColumnItems2 = new List<CheckBoxModel>();

                    foreach (var cb in selectColumn1.Children)
                        selectColumnItems1.Add(new CheckBoxModel(cb as CheckBox));
                    foreach (var cb in selectColumn2.Children)
                        selectColumnItems2.Add(new CheckBoxModel(cb as CheckBox));
                    foreach (var item in selectUid1.Items)
                        selectUidItems1.Add(item.ToString());
                    foreach (var item in selectUid2.Items)
                        selectUidItems2.Add(item.ToString());
                    foreach (var item in selectUid1.SelectedItems)
                        selectUidSelectedItems1.Add(item.ToString());
                    foreach (var item in selectUid2.SelectedItems)
                        selectUidSelectedItems2.Add(item.ToString());
                    foreach (var item in selectLoad1.Items)
                        selectLoadItems1.Add(item.ToString());
                    foreach (var item in selectLoad2.Items)
                        selectLoadItems2.Add(item.ToString());
                    foreach (var item in selectLoad1.SelectedItems)
                        selectLoadSelectedItems1.Add(item.ToString());
                    foreach (var item in selectLoad2.SelectedItems)
                        selectLoadSelectedItems2.Add(item.ToString());

                    Task.Run(() =>
                    {


                        foreach (var cb in selectColumnItems1)
                            if (cb.IsChecked == true)
                                outColumn1.Add(cb.Content.ToString());
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        foreach (var cb in selectColumnItems2)
                            if (cb.IsChecked == true)
                                outColumn2.Add(cb.Content.ToString());
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        if (pathInput1.Substring(pathInput1.Count() - 5).Contains(".csv"))
                            book = excel.Workbooks.OpenXML(pathInput1);
                        else
                            book = excel.Workbooks.Open(pathInput1);
                        sheet = book.ActiveSheet;

                        //Импорт данных файла 1
                        string[,] dataFile1 = new string[selectUidItems1.Count, sheet.Rows.CurrentRegion.EntireRow.Count];
                        for (int i = 0; i < selectUidItems1.Count; i++)
                        {
                            var column = sheet.UsedRange.Columns[i + 1];
                            var array = (Array)column.Cells.Value2;
                            for (int j = 0; j < dataFile1.GetLength(1); j++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                if (array.GetValue(j + 1, 1).ToString() != "")
                                    dataFile1[i, j] = array.OfType<object>().Select(o => o.ToString()).ToList()[j];
                                else
                                    dataFile1[i, j] = "";
                            }
                        }
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        // Создание ключа сравнения для файла 1
                        string[] key1 = new string[sheet.Rows.CurrentRegion.EntireRow.Count];
                        for (int i = 0; i < dataFile1.GetLength(1); i++)
                            foreach (var item in selectUidSelectedItems1)
                                key1[i] += dataFile1[selectUidItems1.IndexOf(item), i];
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        if (pathInput2.Substring(pathInput2.Count() - 5).Contains(".csv"))
                            book = excel.Workbooks.OpenXML(pathInput2);
                        else
                            book = excel.Workbooks.Open(pathInput2);
                        sheet = book.ActiveSheet;

                        //Импорт данных файла 2
                        string[,] dataFile2 = new string[selectUidItems2.Count, sheet.Rows.CurrentRegion.EntireRow.Count];
                        for (int i = 0; i < selectUidItems2.Count; i++)
                        {
                            var column = sheet.UsedRange.Columns[i + 1];
                            var array = (Array)column.Cells.Value2;
                            for (int j = 0; j < dataFile2.GetLength(1); j++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                if (array.GetValue(j + 1, 1).ToString() != "")
                                    dataFile2[i, j] = array.OfType<object>().Select(o => o.ToString()).ToList()[j];
                                else
                                    dataFile2[i, j] = "";
                            }
                        }
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        // Создание ключа сравнения для файла 2
                        string[] key2 = new string[sheet.Rows.CurrentRegion.EntireRow.Count];
                        for (int i = 0; i < dataFile2.GetLength(1); i++)
                            foreach (var item in selectUidSelectedItems2)
                                key2[i] += dataFile2[selectUidItems2.IndexOf(item), i];
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        var exceptKey1 = key1.Where(x => !key2.Any(y => y.Equals(x))).ToList();
                        var exceptKey2 = key2.Where(x => !key1.Any(y => y.Equals(x))).ToList();
                        var intersectKey = key1.Where(x => key2.Any(y => y.Equals(x))).ToList();

                        // Поиск различных по параметру
                        List<string> notEqualParametrKey = new List<string>();
                        foreach (var key in intersectKey)
                        {
                            var res1 = key1.FirstOrDefault(x => x == key);
                            var res2 = key2.FirstOrDefault(x => x == key);
                            var parametr1 = dataFile1[selectedLoad1Index, Array.IndexOf(key1, key)];
                            var parametr2 = dataFile2[selectedLoad2Index, Array.IndexOf(key2, key)];
                            if (parametr1 != parametr2)
                                notEqualParametrKey.Add(key);
                        }
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });

                        pbValue = 0;
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, pbValue });
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.FileName = $"Сравнение {nameFile1} и {nameFile2}.xlsx";
                        saveFileDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
                        if (saveFileDialog.ShowDialog() == true)
                        {
                            Dispatcher.BeginInvoke(new Action(delegate () { progressBar.Maximum = 13; }));

                            pathOutput = saveFileDialog.FileName;
                            book = excel.Workbooks.Add(Type.Missing);
                            sheet = book.ActiveSheet;

                            // Определение столбцов 1 файла
                            sheet.Cells[1, 1].Value = "Нет во втором файле";
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, outColumn1.Count]].Merge();
                            foreach (var column in outColumn1)
                                sheet.Cells[2, outColumn1.IndexOf(column) + 1].Value = dataFile1[selectUidItems1.IndexOf(column), 0];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Определение столбцов 2 файла
                            sheet.Cells[1, outColumn1.Count + 1].Value = "Нет в первом файле";
                            sheet.Range[sheet.Cells[1, outColumn1.Count + 1], sheet.Cells[1, outColumn1.Count + outColumn2.Count]].Merge();
                            foreach (var column in outColumn2)
                                sheet.Cells[2, outColumn2.IndexOf(column) + 1 + outColumn1.Count].Value = dataFile2[selectUidItems2.IndexOf(column), 0];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Определение столбцов различия
                            sheet.Cells[1, outColumn1.Count + outColumn2.Count + 1].Value = "Различный параметр";
                            sheet.Range[sheet.Cells[1, outColumn1.Count + outColumn2.Count + 1], sheet.Cells[1, 2 * outColumn1.Count - selectUidSelectedItems1.Count + 2 * outColumn2.Count]].Merge();
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Столбцы ключа для различия
                            foreach (var column in selectUidSelectedItems1)
                                sheet.Cells[2, selectUidSelectedItems1.IndexOf(column) + outColumn1.Count + outColumn2.Count + 1].Value = dataFile1[selectUidItems1.IndexOf(column), 0];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            List<string> noKey1 = new List<string>();
                            foreach (var column in outColumn1)
                                if (!selectUidSelectedItems1.Contains(column))
                                    noKey1.Add(column);
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            List<string> noKey2 = new List<string>();
                            foreach (var column in outColumn2)
                                if (!selectUidSelectedItems2.Contains(column))
                                    noKey2.Add(column);
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Столбцы файлов для различия (без ключа, ключ - общий)
                            foreach (var column in noKey1)
                                if (!selectUidSelectedItems1.Contains(column))
                                    sheet.Cells[2, noKey1.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count + 1].Value = dataFile1[selectUidItems1.IndexOf(column), 0] + "_1";
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            foreach (var column in noKey2)
                                if (!selectUidSelectedItems2.Contains(column))
                                    sheet.Cells[2, noKey2.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + noKey1.Count + outColumn2.Count + 1].Value = dataFile2[selectUidItems2.IndexOf(column), 0] + "_2";
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });


                            // Значения разности 1
                            for (int i = 0; i < exceptKey1.Count; i++)
                                foreach (var column in outColumn1)
                                    if (Double.TryParse(dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, exceptKey1[i])], out double value))
                                        sheet.Cells[i + 3, outColumn1.IndexOf(column) + 1].Value = value;
                                    else
                                        sheet.Cells[i + 3, outColumn1.IndexOf(column) + 1].Value = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, exceptKey1[i])];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Значения разности 2
                            for (int i = 0; i < exceptKey2.Count; i++)
                                foreach (var column in outColumn2)
                                    if (Double.TryParse(dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, exceptKey2[i])], out double value))
                                        sheet.Cells[i + 3, outColumn2.IndexOf(column) + 1 + outColumn1.Count].Value = value;
                                    else
                                        sheet.Cells[i + 3, outColumn2.IndexOf(column) + 1 + outColumn1.Count].Value = dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, exceptKey2[i])];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Значения ключа для различия
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                                foreach (var column in selectUidSelectedItems1)
                                    if (Double.TryParse(dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])], out double value))
                                        sheet.Cells[i + 3, selectUidSelectedItems1.IndexOf(column) + outColumn1.Count + outColumn2.Count + 1].Value = value;
                                    else
                                        sheet.Cells[i + 3, selectUidSelectedItems1.IndexOf(column) + outColumn1.Count + outColumn2.Count + 1].Value = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            // Значения файлов для различия (без ключа, ключ - общий)
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                                foreach (var column in noKey1)
                                    if (!selectUidSelectedItems1.Contains(column))
                                        if (Double.TryParse(dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])], out double value))
                                            sheet.Cells[i + 3, noKey1.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count + 1].Value = value;
                                        else
                                            sheet.Cells[i + 3, noKey1.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count + 1].Value = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])];


                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                                foreach (var column in noKey2)
                                    if (!selectUidSelectedItems1.Contains(column))
                                        if (Double.TryParse(dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, notEqualParametrKey[i])], out double value))
                                            sheet.Cells[i + 3, noKey2.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + noKey1.Count + outColumn2.Count + 1].Value = value;
                                        else
                                            sheet.Cells[i + 3, noKey2.IndexOf(column) + selectUidSelectedItems1.Count + outColumn1.Count + noKey1.Count + outColumn2.Count + 1].Value = dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, notEqualParametrKey[i])];
                            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });


                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 2 * outColumn1.Count - selectUidSelectedItems1.Count + 2 * outColumn2.Count]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; ;
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 2 * outColumn1.Count - selectUidSelectedItems1.Count + 2 * outColumn2.Count]].Interior.Color = Excel.XlRgbColor.rgbLightSkyBlue;
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 2 * outColumn1.Count - selectUidSelectedItems1.Count + 2 * outColumn2.Count]].Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                            sheet.Columns.AutoFit();
                            try
                            {
                                excel.Application.ActiveWorkbook.SaveAs(pathOutput);
                            }
                            catch
                            {
                                MessageBox.Show("Нет доступа для записи в файл.");
                            }
                        }

                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        Marshal.ReleaseComObject(sheet);
                        book.Close();
                        Marshal.ReleaseComObject(book);
                        excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        pbValue = 0;
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, pbValue });
                    }).ContinueWith(delegate { textOutput.Visibility = Visibility.Visible; });
                    
                }
            }
            catch
            {
                MessageBox.Show("Произошла непредвиденная ошибка");
            }
            finally
            {
                excel.Quit();
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {

        }
        public class CheckBoxModel
        {
            public CheckBoxModel(CheckBox checkbox)
            {
                IsChecked = checkbox.IsChecked;
                Content = checkbox.Content;
            }
            public bool? IsChecked { get; set; }
            public object Content { get; set; }

        }
    }
}
