using Microsoft.Win32;
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
            log += IsOfficeInstalled();
            log += IsDotNetInstalled();
            logWindow = new Log();
            logWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            WtiteLog($"[{DateTime.Now}] Инициализация приложения завершена.\r\n");
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
        public static string log;
        public static Log logWindow;

        private static void WtiteLog(string text)
        {
            log += text;
            logWindow.UpdateInfo(log);
        }
        private static string IsDotNetInstalled()
        {
            const string subkey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";

            using (var ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey))
            {
                if (ndpKey != null && ndpKey.GetValue("Release") != null)
                {
                    return $"[{DateTime.Now}] .NET Framework Version: {CheckFor45PlusVersion((int)ndpKey.GetValue("Release"))}\r\n";
                }
                else
                {
                    return $"[{DateTime.Now}] .NET Framework Version 4.5 or later is not detected.\r\n";
                }
            }
        }
        // Checking the version using >= enables forward compatibility.
        static string CheckFor45PlusVersion(int releaseKey)
        {
            if (releaseKey >= 533320)
                return "4.8.1 or later";
            if (releaseKey >= 528040)
                return "4.8";
            if (releaseKey >= 461808)
                return "4.7.2";
            if (releaseKey >= 461308)
                return "4.7.1";
            if (releaseKey >= 460798)
                return "4.7";
            if (releaseKey >= 394802)
                return "4.6.2";
            if (releaseKey >= 394254)
                return "4.6.1";
            if (releaseKey >= 393295)
                return "4.6";
            if (releaseKey >= 379893)
                return "4.5.2";
            if (releaseKey >= 378675)
                return "4.5.1";
            if (releaseKey >= 378389)
                return "4.5";
            // This code should never execute. A non-null release key should mean
            // that 4.5 or later is installed.
            return $"No 4.5 or later version detected";
        }
        private static string IsOfficeInstalled()
        {
            string res;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe");
            if (key != null)
            {
                res = $"[{DateTime.Now}] Найден установленный Excel по пути {key.GetValue("Path")}\r\n";
                key.Close();
            }
            else
            {
                res = $"[{DateTime.Now}] Не найден установленный Excel\r\n";
                MessageBox.Show("Для работы программы необходим пакет MS Office Excel");
            }

            return res;
        }
        private void input1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WtiteLog($"[{DateTime.Now}] Запрос выбора файла 1\r\n");
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
                    WtiteLog($"[{DateTime.Now}] Выбран файл {pathInput1}\r\n");
                    if (pathInput1.Substring(pathInput1.Count() - 5).Contains(".csv"))
                        book = excel.Workbooks.OpenXML(pathInput1);
                    else
                        book = excel.Workbooks.Open(pathInput1);
                    sheet = book.ActiveSheet;
                    rowCount1 = sheet.Rows.CurrentRegion.EntireRow.Count;
                    WtiteLog($"[{DateTime.Now}] Число строк книги: {rowCount1}\r\n");
                    WtiteLog($"[{DateTime.Now}] Число колонок книги: {sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column}\r\n");
                    List<string> head = new List<string>();
                    for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                    {
                        WtiteLog($"[{DateTime.Now}] Добавлена колонка \"{sheet.Cells[1, i].Value}\"\r\n");
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
                    if (excel != null)
                        excel.Quit();
                    WtiteLog($"[{DateTime.Now}] Колонки успешно загружены\r\n");
                }
            }
            catch(Exception exc)
            {
                WtiteLog($"[{DateTime.Now}][ERROR] {exc.Message}\r\n");
                WtiteLog($"{exc.StackTrace}\r\n");
                if (excel != null)
                    excel.Quit();
                selectUid1.Items.Clear();
                selectLoad1.Items.Clear();
                selectColumn1.Children.Clear();
                MessageBox.Show("Входной файл имел неверный формат или недопустымые названиия столбцов");
                WtiteLog($"[{DateTime.Now}][ERROR] Входной файл имел неверный формат или недопустымые названиия столбцов\r\n");
            }
            finally
            {
                if (excel != null)
                    excel.Quit();
            }
        }

        private void input2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WtiteLog($"[{DateTime.Now}] Запрос выбора файла 2\r\n");
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
                    WtiteLog($"[{DateTime.Now}] Выбран файл {pathInput2}\r\n");
                    if (pathInput2.Substring(pathInput2.Count() - 5).Contains(".csv"))
                        book = excel.Workbooks.OpenXML(pathInput2);
                    else
                        book = excel.Workbooks.Open(pathInput2);
                    sheet = book.ActiveSheet;
                    rowCount2 = sheet.Rows.CurrentRegion.EntireRow.Count;
                    WtiteLog($"[{DateTime.Now}] Число строк книги: {rowCount2}\r\n");
                    WtiteLog($"[{DateTime.Now}] Число колонок книги: {sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column}\r\n");
                    List<string> head = new List<string>();
                    for (int i = 1; i <= sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; i++)
                    {
                        WtiteLog($"[{DateTime.Now}] Добавлена колонка \"{sheet.Cells[1, i].Value}\"\r\n");
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
                    if (excel != null)
                        excel.Quit();
                }
            }
            catch(Exception exc)
            {
                WtiteLog($"[{DateTime.Now}][ERROR] {exc.Message}\r\n");
                WtiteLog($"{exc.StackTrace}\r\n");
                if (excel != null)
                    excel.Quit();
                selectUid2.Items.Clear();
                selectLoad2.Items.Clear();
                selectColumn2.Children.Clear();
                MessageBox.Show("Входной файл имел неверный формат или недопустымые названиия столбцов");
                WtiteLog($"[{DateTime.Now}][ERROR] Входной файл имел неверный формат или недопустымые названиия столбцов\r\n");
            }
            finally
            {
                if (excel != null)
                    excel.Quit();
            }
            WtiteLog($"[{DateTime.Now}] Колонки успешно загружены\r\n");
        }

        private void compare_Click(object sender, RoutedEventArgs e)
        {
            WtiteLog($"[{DateTime.Now}] Натажата кнопка сравнения\r\n");
            int errors = 0;
            try
            {
                if (pathInput1 == null)
                {
                    errors++;
                    MessageBox.Show("Первый файл не загружен!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Первый файл не загружен!\r\n");
                }
                else if (pathInput2 == null)
                {
                    errors++;
                    MessageBox.Show("Второй файл не загружен!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Второй файл не загружен!\r\n");
                }
                else if (selectLoad1.SelectedItem == null)
                {
                    errors++;
                    MessageBox.Show("Не выбрано сравниваемое значение для файла 1!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Не выбрано сравниваемое значение для файла 1!\r\n");
                }
                else if (selectLoad2.SelectedItem == null)
                {
                    errors++;
                    MessageBox.Show("Не выбрано сравниваемое значение для файла 2!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Не выбрано сравниваемое значение для файла 2!\r\n");
                }
                else if (selectColumn1.Children.Count == 0)
                {
                    errors++;
                    MessageBox.Show("Количество выводимых колонок для файла 1 должно быть больше 0!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Количество выводимых колонок для файла 1 должно быть больше 0!\r\n");
                }
                else if (selectColumn2.Children.Count == 0)
                {
                    errors++;
                    MessageBox.Show("Количество выводимых колонок для файла 2 должно быть больше 0!");
                    WtiteLog($"[{DateTime.Now}][ERROR] Количество выводимых колонок для файла 2 должно быть больше 0!\r\n");
                }
                if (errors == 0)
                {
                    WtiteLog($"[{DateTime.Now}] Ошибки отсутствуют\r\n");
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
                        log += $"[{DateTime.Now}] Старт сравнения файлов\r\n";

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
                                if (array.GetValue(j + 1, 1) != null)
                                    if (array.GetValue(j + 1, 1).ToString() != "")
                                        dataFile1[i, j] = array.GetValue(j + 1, 1).ToString();
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
                                if (array.GetValue(j + 1, 1) != null)
                                    if (array.GetValue(j + 1, 1).ToString() != "")
                                        dataFile2[i, j] = array.GetValue(j + 1, 1).ToString();
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
                        log += $"[{DateTime.Now}] Сравнение файлов завершено успешно\r\n";
                        pbValue = 0;
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, pbValue });
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.FileName = $"Сравнение {nameFile1} и {nameFile2}.xlsx";
                        saveFileDialog.Filter = "Книга Excel (.xlsx) | *.xlsx|All files| *.*";
                        log += $"[{DateTime.Now}] Запрос сохранения файла\r\n";
                        if (saveFileDialog.ShowDialog() == true)
                        {
                            log += $"[{DateTime.Now}] Старт сохранения файла\r\n";
                            Dispatcher.BeginInvoke(new Action(delegate ()
                            {
                                progressBar.Maximum = 13 + exceptKey1.Count + exceptKey2.Count + 3 * notEqualParametrKey.Count;
                            }));

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


                            //// Значения разности 1
                            string[,] result = new string[exceptKey1.Count, outColumn1.Count];
                            for (int i = 0; i < exceptKey1.Count; i++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                foreach (var column in outColumn1)
                                    result[i, outColumn1.IndexOf(column)] = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, exceptKey1[i])];
                            }
                            sheet.Range[sheet.Cells[3, 1], sheet.Cells[exceptKey1.Count + 2, outColumn1.Count]] = result;

                            // Значения разности 2
                            result = new string[exceptKey2.Count, outColumn2.Count];
                            for (int i = 0; i < exceptKey2.Count; i++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                foreach (var column in outColumn2)
                                    result[i, outColumn2.IndexOf(column)] = dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, exceptKey2[i])];
                            }
                            sheet.Range[sheet.Cells[3, 1 + outColumn1.Count], sheet.Cells[exceptKey2.Count + 2, outColumn2.Count + outColumn1.Count]] = result;

                            // Значения ключа для различия
                            result = new string[notEqualParametrKey.Count, selectUidSelectedItems1.Count];
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                foreach (var column in selectUidSelectedItems1)
                                    result[i, selectUidSelectedItems1.IndexOf(column)] = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])];
                            }
                            sheet.Range[sheet.Cells[3, 1 + outColumn1.Count + outColumn2.Count], sheet.Cells[notEqualParametrKey.Count + 2, outColumn1.Count + outColumn2.Count + selectUidSelectedItems1.Count]] = result;

                            // Значения файлов для различия (без ключа, ключ - общий)
                            result = new string[notEqualParametrKey.Count, noKey1.Count];
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                foreach (var column in noKey1)
                                    if (!selectUidSelectedItems1.Contains(column))
                                        result[i, noKey1.IndexOf(column)] = dataFile1[selectUidItems1.IndexOf(column), Array.IndexOf(key1, notEqualParametrKey[i])];
                            }
                            sheet.Range[sheet.Cells[3, 1 + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count], sheet.Cells[notEqualParametrKey.Count + 2, noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]] = result;


                            result = new string[notEqualParametrKey.Count, noKey2.Count];
                            for (int i = 0; i < notEqualParametrKey.Count; i++)
                            {
                                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++pbValue });
                                foreach (var column in noKey2)
                                    if (!selectUidSelectedItems1.Contains(column))
                                        result[i, noKey2.IndexOf(column)] = dataFile2[selectUidItems2.IndexOf(column), Array.IndexOf(key2, notEqualParametrKey[i])];
                            }
                            sheet.Range[sheet.Cells[3, 1 + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count], sheet.Cells[notEqualParametrKey.Count + 2, noKey2.Count + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]] = result;

                            sheet.Range[sheet.Cells[1, outColumn1.Count + outColumn2.Count + 1], sheet.Cells[1, noKey2.Count + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]].Merge();
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, noKey2.Count + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; ;
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, noKey2.Count + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]].Interior.Color = Excel.XlRgbColor.rgbLightSkyBlue;
                            sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, noKey2.Count + noKey1.Count + selectUidSelectedItems1.Count + outColumn1.Count + outColumn2.Count]].Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                            sheet.Columns.AutoFit();
                            try
                            {
                                excel.Application.ActiveWorkbook.SaveAs(pathOutput);
                                log += $"[{DateTime.Now}] Файл успешно сохранен\r\n";
                            }
                            catch (Exception exc)
                            {
                                log += $"[{DateTime.Now}][ERROR] {exc.Message}\r\n";
                                log += $"{exc.StackTrace}\r\n";
                                MessageBox.Show("Нет доступа для записи в файл.");
                            }
                        }

                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        Marshal.ReleaseComObject(sheet);
                        book.Close();
                        Marshal.ReleaseComObject(book);
                        if (excel != null)
                            excel.Quit();
                        Marshal.ReleaseComObject(excel);
                        pbValue = 0;
                        Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, pbValue });
                        log += $"[{DateTime.Now}] Очищены ресурсы\r\n";
                    }).ContinueWith(UpdateResult, TaskScheduler.FromCurrentSynchronizationContext());

                }
            }
            catch (Exception exc)
            {
                WtiteLog($"[{DateTime.Now}][ERROR] {exc.Message}\r\n");
                WtiteLog($"{exc.StackTrace}\r\n");
                MessageBox.Show("Произошла непредвиденная ошибка");
            }
            finally
            {
                if (excel != null)
                    excel.Quit();
            }
        }
        private void UpdateResult(Task obj)
        {
            textOutput.Visibility = Visibility.Visible;
            logWindow.UpdateInfo(log);
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

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            About about = new About();
            about.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            about.Topmost = true;
            about.Show();
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            logWindow.UpdateInfo(log);
            logWindow.Show();
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            Help helpWindow = new Help();
            helpWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            helpWindow.Show();
        }
    }
}
