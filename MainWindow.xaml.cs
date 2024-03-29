﻿using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace KFV
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public ObservableCollection<Prop> prop = new ObservableCollection<Prop>();

        static string title = "Вид труб;Прокат в метрах;Размер з.;Размер г.;" +
                              "Проходимость для производства;Метро проходов;Маршрут;Стан ХПТ готовый;Норма;Кол-во смен;ХПТ 1 предфинал;" +
                              "Метры 1 предфина;ХПТ 2 предфинал;Метры 2 предфина;ХПТ 3 предфинал;Метры 3 предфина;ХПТ 4 предфинал;Метры 4 предфина;";

        static string foglio = "qwa.xls";

        string path = "2.txt";

        #region avgNormaGot
        decimal got32 = 0;
        decimal got55 = 0;
        decimal got75 = 0;
        decimal got55_2 = 0;
        decimal got90 = 0;
        decimal normaGot1 = 0;
        decimal normaGot3 = 0;
        decimal normaGot2 = 0;
        decimal normaGot4 = 0;
        decimal normaGot5 = 0;
        decimal stangot1 = 0;
        decimal stangot2 = 0;
        decimal stangot3 = 0;
        decimal stangot4 = 0;
        decimal stangot5 = 0;

        decimal zagot55 = 0;
        decimal zagot75 = 0;
        decimal zagot90 = 0;
        decimal zastangot1 = 0;
        decimal zastangot2 = 0;
        decimal zastangot3 = 0;
        decimal zastangot4 = 0;
        decimal zastangot5 = 0;

        decimal m1 = 0;
        decimal m2 = 0;
        decimal m3 = 0;
        decimal m4 = 0;
        int n = 0;
        #endregion

        public MainWindow()
        {
            InitializeComponent();
            dataGrid.ItemsSource = prop;
        }

        /// <summary>
        /// Запуск потока Excel
        /// </summary>
        void OpenFile()
        {
            try
            {
                var excelApp = new Excel.Application();
                excelApp.Visible = true;
                Workbooks books = excelApp.Workbooks;
                Workbook sheets = books.Open(foglio);
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        /// <summary>
        /// Разметка Excel
        /// </summary>
        /// <param name="worksheet"></param>
        void AddCells(Worksheet worksheet, int i)
        {
            Fg(worksheet);

            #region Таблицы
            worksheet.Cells[i + 2, 5] = $"Метропроходы в сут. по вкл. {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 2, 5].Font.Bold = true;
            worksheet.Cells[i + 2, 5].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 3, 5] = $"Метропроходы заготовки по вкл. {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 3, 5].Font.Bold = true;
            worksheet.Cells[i + 3, 5].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 4, 5] = "Итого норм для проката заготовки";
            worksheet.Cells[i + 4, 5].Font.Bold = true;
            worksheet.Cells[i + 4, 5].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 5, 5] = $"Итого норм загот./сут для проката заготовки до вкл {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 5, 5].Font.Bold = true;
            worksheet.Cells[i + 5, 5].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            worksheet.Cells[i + 2, 7] = $"Итого норм для проката готового размера";
            worksheet.Cells[i + 2, 7].Font.Bold = true;
            worksheet.Cells[i + 2, 7].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 3, 7] = $"Итого норм готовых/сут до вкл {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 3, 7].Font.Bold = true;
            worksheet.Cells[i + 3, 7].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 4, 7] = $"Итого кол-во норм/сут (заг.+гот) до вкл {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 4, 7].Font.Bold = true;
            worksheet.Cells[i + 4, 7].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            worksheet.Cells[i + 7, 1] = "Тип";
            worksheet.Cells[i + 7, 1].Font.Bold = true;
            worksheet.Cells[i + 7, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 8, 1] = "ХПТ 32";
            worksheet.Cells[i + 8, 1].Font.Bold = true;
            worksheet.Cells[i + 8, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 9, 1] = "ХПТ 55";
            worksheet.Cells[i + 9, 1].Font.Bold = true;
            worksheet.Cells[i + 9, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 10, 1] = "ХПТ 55-2";
            worksheet.Cells[i + 10, 1].Font.Bold = true;
            worksheet.Cells[i + 10, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 11, 1] = "ХПТ 75";
            worksheet.Cells[i + 11, 1].Font.Bold = true;
            worksheet.Cells[i + 11, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 12, 1] = "ХПТ 90";
            worksheet.Cells[i + 12, 1].Font.Bold = true;
            worksheet.Cells[i + 12, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;


            worksheet.Cells[i + 7, 2] = "Прокат готовых, м";
            worksheet.Cells[i + 7, 2].Font.Bold = true;
            worksheet.Cells[i + 7, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 7, 3] = "Средняя норма, м";
            worksheet.Cells[i + 7, 3].Font.Bold = true;
            worksheet.Cells[i + 7, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 7, 4] = $"Необходимое количество станов в сутки до вкл {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 7, 4].Font.Bold = true;
            worksheet.Cells[i + 7, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;


            worksheet.Cells[i + 14, 1] = "Тип";
            worksheet.Cells[i + 14, 1].Font.Bold = true;
            worksheet.Cells[i + 14, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 15, 1] = "ХПТ 55";
            worksheet.Cells[i + 15, 1].Font.Bold = true;
            worksheet.Cells[i + 15, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 16, 1] = "ХПТ 75";
            worksheet.Cells[i + 16, 1].Font.Bold = true;
            worksheet.Cells[i + 16, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 17, 1] = "ХПТ 90";
            worksheet.Cells[i + 17, 1].Font.Bold = true;
            worksheet.Cells[i + 17, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;


            worksheet.Cells[i + 14, 2] = "Прокат заготовки, м";
            worksheet.Cells[i + 14, 2].Font.Bold = true;
            worksheet.Cells[i + 14, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 14, 3] = "Средняя норма, м";
            worksheet.Cells[i + 14, 3].Font.Bold = true;
            worksheet.Cells[i + 14, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 14, 4] = $"Необходимое количество станов в сутки до {datePicker.SelectedDate.Value.ToString("dd.MM")}";
            worksheet.Cells[i + 14, 4].Font.Bold = true;
            worksheet.Cells[i + 14, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;


            worksheet.Cells[i + 20, 1] = "ХПТ 32";
            worksheet.Cells[i + 20, 1].Font.Bold = true;
            worksheet.Cells[i + 20, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 21, 1] = "ХПТ 55";
            worksheet.Cells[i + 21, 1].Font.Bold = true;
            worksheet.Cells[i + 21, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 22, 1] = "ХПТ 55-2";
            worksheet.Cells[i + 22, 1].Font.Bold = true;
            worksheet.Cells[i + 22, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 23, 1] = "ХПТ 75";
            worksheet.Cells[i + 23, 1].Font.Bold = true;
            worksheet.Cells[i + 23, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 24, 1] = "ХПТ 90";
            worksheet.Cells[i + 24, 1].Font.Bold = true;
            worksheet.Cells[i + 24, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 19, 1] = "Тип";
            worksheet.Cells[i + 19, 1].Font.Bold = true;
            worksheet.Cells[i + 19, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;


            worksheet.Cells[i + 19, 2] = "Итого количество в сутки";
            worksheet.Cells[i + 19, 2].Font.Bold = true;
            worksheet.Cells[i + 19, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            #endregion
        }

        private void AllBorders(Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range tRange = worksheet.UsedRange;
            tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
        }

        /// <summary>
        /// Запись в Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application();
                Workbook workbook = app.Workbooks.Add(System.Reflection.Missing.Value);
                Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);
                AllBorders(worksheet);

                int i = 2;

                ExcelData(worksheet, i);

                AddCells(worksheet, n);

                Avg1(prop, worksheet, n);
                Avg2(prop, worksheet, n);

                AVGNORM(prop, worksheet, n);

                ProkatSum(prop, worksheet, n);

                ProkatSum1(prop, worksheet, n);

                AvgEndGot(worksheet, n);
                AvgEndGot1(worksheet, n);


                worksheet.Cells[n + 20, 2] = stangot1 + zastangot1;
                worksheet.Cells[n + 20, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 21, 2] = stangot2 + zastangot2;
                worksheet.Cells[n + 21, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 22, 2] = stangot3 + zastangot3;
                worksheet.Cells[n + 22, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 23, 2] = stangot4 + zastangot4;
                worksheet.Cells[n + 23, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 24, 2] = stangot5 + zastangot5;
                worksheet.Cells[n + 24, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                worksheet.Cells[n + 15, 3] = 700;
                worksheet.Cells[n + 15, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 16, 3] = 700;
                worksheet.Cells[n + 16, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[n + 17, 3] = 700;
                worksheet.Cells[n + 17, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                //Excel.Range range1 = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[n, 18]);
                //RowColor(range1);

                var a = System.Reflection.Missing.Value;
                workbook.SaveAs(foglio);
                MessageBoxResult result = MessageBox.Show("Открыть файл?", "Открыть?", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    OpenFile();
                }
                else
                    MessageBox.Show("Файл сохранен");

                workbook.Close(a, a, a);
                app.Quit();

                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Не выбрана дата или закройте документ" + ex.Message);
            }
        }

        void Avg1(ObservableCollection<Prop> prop, Worksheet worksheet, int i)
        {
            MainView mv = new MainView();
            m1 = mv.Avg(prop, worksheet, i + 2, 8);
            worksheet.Cells[i + 2, 8] = m1;
            worksheet.Cells[i + 2, 8].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 3, 8] = m1 / Date();
            worksheet.Cells[i + 3, 8].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        void Avg2(ObservableCollection<Prop> prop, Worksheet worksheet, int i)
        {
            MainView mv = new MainView();
            m2 = mv.Afg2(prop);
            m3 = mv.Afg1(prop);
            m4 = m2 - m3;
            decimal l = ((m2 - m3) / 800) / Date();
            worksheet.Cells[i + 2, 6] = m2 / Date();
            worksheet.Cells[i + 2, 6].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 3, 6] = m4;
            worksheet.Cells[i + 3, 6].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 4, 6] = (m2 - m3) / 800;
            worksheet.Cells[i + 4, 6].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 5, 6] = ((m2 - m3) / 800) / Date();
            worksheet.Cells[i + 5, 6].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 4, 8] = (m1 / Date()) + l;
            worksheet.Cells[i + 4, 8].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        int Date()
        {
            int v = 2;
            DateTime date1 = DateTime.Now;
            DateTime date2 = Convert.ToDateTime(datePicker.SelectedDate.Value.ToString("dd.MM.yyyy"));
            v += (date2 - date1).Days;
            return v;
        }

        /// <summary>
        /// Средня норма готовой
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="worksheet"></param>
        /// <param name="i"></param>
        void AVGNORM(ObservableCollection<Prop> prop, Worksheet worksheet, int i)
        {
            MainView mv = new MainView();
            normaGot1 = mv.AvgNORMA(prop, worksheet, i + 8, "ХПТ 32", 3);
            normaGot2 = mv.AvgNORMA(prop, worksheet, i + 9, "ХПТ 55", 3);
            normaGot3 = mv.AvgNORMA(prop, worksheet, i + 10, "ХПТ 55-2", 3);
            normaGot4 = mv.AvgNORMA(prop, worksheet, i + 11, "ХПТ 75", 3);
            normaGot5 = mv.AvgNORMA(prop, worksheet, i + 12, "ХПТ 90", 3);
        }

        /// <summary>
        /// Прокат готовый
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="worksheet"></param>
        /// <param name="i"></param>
        void ProkatSum(ObservableCollection<Prop> prop, Worksheet worksheet, int i)
        {
            MainView mv = new MainView();
            got32 = mv.Prokat(prop, worksheet, i + 8, "ХПТ 32", 2);
            got55 = mv.Prokat(prop, worksheet, i + 9, "ХПТ 55", 2);
            got55_2 = mv.Prokat(prop, worksheet, i + 10, "ХПТ 55-2", 2);
            got75 = mv.Prokat(prop, worksheet, i + 11, "ХПТ 75", 2);
            got90 = mv.Prokat(prop, worksheet, i + 12, "ХПТ 90", 2);
        }

        void AvgEndGot(Worksheet worksheet, int i)
        {
            try
            {
                stangot1 = worksheet.Cells[i + 8, 4] = got32 / normaGot1 / Date();
                worksheet.Cells[i + 8, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                stangot2 = worksheet.Cells[i + 9, 4] = got55 / normaGot2 / Date();
                worksheet.Cells[i + 9, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                stangot3 = worksheet.Cells[i + 10, 4] = got55_2 / normaGot3 / Date();
                worksheet.Cells[i + 10, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                stangot4 = worksheet.Cells[i + 11, 4] = got75 / normaGot4 / Date();
                worksheet.Cells[i + 11, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                stangot5 = worksheet.Cells[i + 12, 4] = got90 / normaGot5 / Date();
                worksheet.Cells[i + 12, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
        }

        /// <summary>
        /// Прокат заг
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="worksheet"></param>
        /// <param name="i"></param>
        void ProkatSum1(ObservableCollection<Prop> prop, Worksheet worksheet, int i)
        {
            MainView mv = new MainView();
            zagot55 = mv.Prokat1(prop, worksheet, i, "ХПТ 55");
            zagot75 = mv.Prokat1(prop, worksheet, i, "ХПТ 75");
            zagot90 = mv.Prokat1(prop, worksheet, i, "ХПТ 90");

            worksheet.Cells[i + 15, 2] = zagot55;
            worksheet.Cells[i + 15, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 16, 2] = zagot75;
            worksheet.Cells[i + 16, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Cells[i + 17, 2] = zagot90;
            worksheet.Cells[i + 17, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="i"></param>
        void AvgEndGot1(Worksheet worksheet, int i)
        {
            try
            {
                if (zagot55 != 0)
                    zastangot2 = worksheet.Cells[i + 15, 4] = zagot55 / 700 / Date();
                worksheet.Cells[i + 15, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                if (zagot75 != 0)
                    zastangot4 = worksheet.Cells[i + 16, 4] = zagot75 / 700 / Date();
                worksheet.Cells[i + 16, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
            try
            {
                if (zagot90 != 0)
                    zastangot5 = worksheet.Cells[i + 17, 4] = zagot90 / 700 / Date();
                worksheet.Cells[i + 17, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            catch { }
        }

        /// <summary>
        /// Заполнение Excel datagrid
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="i"></param>
        /// <param name="n"></param>
        void ExcelData(Worksheet worksheet, int i)
        {
            foreach (Prop p in prop)
            {
                worksheet.Cells[i, 1] = p.SelectedString;
                worksheet.Cells[i, 1].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 2] = p.V2;
                worksheet.Cells[i, 2].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 3] = p.V4;
                worksheet.Cells[i, 3].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 4] = p.V5;
                worksheet.Cells[i, 4].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 5] = p.Pas;
                worksheet.Cells[i, 5].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 6] = p.V1;
                worksheet.Cells[i, 6].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 7] = p.V3;
                worksheet.Cells[i, 7].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 8] = p.SelectedString1;
                worksheet.Cells[i, 8].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 9] = p.SelectedString2;
                worksheet.Cells[i, 9].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 10] = p.V6;
                worksheet.Cells[i, 10].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 11] = p.SelectedString6;
                worksheet.Cells[i, 11].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 12] = p.Metri4;
                worksheet.Cells[i, 12].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 13] = p.SelectedString3;
                worksheet.Cells[i, 13].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 14] = p.Metri3;
                worksheet.Cells[i, 14].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 15] = p.SelectedString4;
                worksheet.Cells[i, 15].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 16] = p.Metri2;
                worksheet.Cells[i, 16].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 17] = p.SelectedString5;
                worksheet.Cells[i, 17].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                worksheet.Cells[i, 18] = p.Metri1;
                worksheet.Cells[i, 18].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                i++;
                n = i;
            }
        }

        /// <summary>
        /// Оглавление
        /// </summary>
        /// <param name="worksheet"></param>
        void Fg(Worksheet worksheet)
        {
            string[] b = title.Split(';');
            for (int i = 1; i <= 18; i++)
            {
                worksheet.Cells[1, i] = b[i - 1];
                worksheet.Cells[1, i].Font.Bold = true;
                worksheet.Cells[1, i].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
        }

        /// <summary>
        /// Добавление в коллекцию
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Add_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] b = five.Text.Split(' '); ///102(0) 65(1) 45(2) 21(3)
                string[] n = six.Text.Split(' ');

                float m = (float)Convert.ToDouble(seven.Text);
                float s1, stanka2, s2, stanka4, stanka3, stanka5;
                float d1, marsh2, d2, marsh4, marsh3, marsh5;

                if (b.Length == 2 && n.Length == 2)
                {
                    d1 = (float)Convert.ToDouble(b[0]);
                    d2 = (float)Convert.ToDouble(b[1]);

                    s1 = (float)Convert.ToDouble(n[0]);
                    s2 = (float)Convert.ToDouble(n[1]);

                    prop.Add(new Prop(d1, s1, d2, s2, m) { V3 = $"{b[0]}x{n[0]}|{b[1]}x{n[1]}", V2 = seven.Text });
                }
                else if (b.Length == 3 && n.Length == 3)
                {
                    d1 = (float)Convert.ToDouble(b[0]);
                    marsh2 = (float)Convert.ToDouble(b[1]);
                    d2 = (float)Convert.ToDouble(b[2]);

                    s1 = (float)Convert.ToDouble(n[0]);
                    stanka2 = (float)Convert.ToDouble(n[1]);
                    s2 = (float)Convert.ToDouble(n[2]);

                    prop.Add(new Prop(d1, s1, d2, s2, marsh2, stanka2, m) { V3 = $"{b[0]}x{n[0]}|{b[1]}x{n[1]}|{b[2]}x{n[2]}", V2 = seven.Text });
                }
                else if (b.Length == 4 && n.Length == 4)
                {
                    d1 = (float)Convert.ToDouble(b[0]);
                    marsh2 = (float)Convert.ToDouble(b[1]);
                    marsh3 = (float)Convert.ToDouble(b[2]);
                    d2 = (float)Convert.ToDouble(b[3]);

                    s1 = (float)Convert.ToDouble(n[0]);
                    stanka2 = (float)Convert.ToDouble(n[1]);
                    stanka3 = (float)Convert.ToDouble(n[2]);
                    s2 = (float)Convert.ToDouble(n[3]);

                    prop.Add(new Prop(d1, s1, d2, s2, marsh2, stanka2, marsh3, stanka3, m) { V3 = $"{b[0]}x{n[0]}|{b[1]}x{n[1]}|{b[2]}x{n[2]}|{b[3]}x{n[3]}", V2 = seven.Text });
                }
                else if (b.Length == 5 && n.Length == 5)
                {
                    d1 = (float)Convert.ToDouble(b[0]);
                    marsh2 = (float)Convert.ToDouble(b[1]);
                    marsh3 = (float)Convert.ToDouble(b[2]);
                    marsh4 = (float)Convert.ToDouble(b[3]);
                    d2 = (float)Convert.ToDouble(b[4]);

                    s1 = (float)Convert.ToDouble(n[0]);
                    stanka2 = (float)Convert.ToDouble(n[1]);
                    stanka3 = (float)Convert.ToDouble(n[2]);
                    stanka4 = (float)Convert.ToDouble(n[3]);
                    s2 = (float)Convert.ToDouble(n[4]);

                    prop.Add(new Prop(d1, s1, d2, s2, marsh2, stanka2, marsh3, stanka3, marsh4, stanka4, m) { V3 = $"{b[0]}x{n[0]}|{b[1]}x{n[1]}|{b[2]}x{n[2]}|{b[3]}x{n[3]}|{b[4]}x{n[4]}", V2 = seven.Text });
                }
                else if (b.Length == 6 && n.Length == 6)
                {
                    d1 = (float)Convert.ToDouble(b[0]);
                    marsh2 = (float)Convert.ToDouble(b[1]);
                    marsh3 = (float)Convert.ToDouble(b[2]);
                    marsh4 = (float)Convert.ToDouble(b[3]);
                    marsh5 = (float)Convert.ToDouble(b[4]);
                    d2 = (float)Convert.ToDouble(b[5]);

                    s1 = (float)Convert.ToDouble(n[0]);
                    stanka2 = (float)Convert.ToDouble(n[1]);
                    stanka3 = (float)Convert.ToDouble(n[2]);
                    stanka4 = (float)Convert.ToDouble(n[3]);
                    stanka5 = (float)Convert.ToDouble(n[4]);
                    s2 = (float)Convert.ToDouble(n[5]);

                    prop.Add(new Prop(d1, s1, d2, s2, marsh2, stanka2, marsh3, stanka3, marsh4, stanka4, marsh5, stanka5, m) { V3 = $"{b[0]}x{n[0]}|{b[1]}x{n[1]}|{b[2]}x{n[2]}|{b[3]}x{n[3]}|{b[4]}x{n[4]}|{b[5]}x{n[5]}", V2 = seven.Text });
                }
                else
                {
                    MessageBox.Show("Не введен маршрут");
                }
            }
            catch
            {
                MessageBox.Show("Ошибка данных");
            }

            five.Clear();
            six.Clear();
            seven.Clear();

            five.Text = "Маршрут диаметр";
            six.Text = "Маршрут толщина";
            seven.Text = "Прокат";
        }

        /// <summary>
        /// Открыть файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlsxApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = xlsxApp.Workbooks.Open(foglio);
            xlsxApp.Visible = true;
        }

        #region OpenFileDialoge
        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFileDialog = new OpenFileDialog();
            oFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";

            if (oFileDialog.ShowDialog() == false)
            {
                return;
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;
            string sCellData = "";
            double dCellData;

            xlApp = new Excel.Application();

            try
            {
                xlWorkBook = xlApp.Workbooks.Open(oFileDialog.FileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;

                System.Data.DataTable dt = new System.Data.DataTable();

                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    str = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;
                    dt.Columns.Add(str, typeof(string));
                }

                for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string sData = "";
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        try
                        {
                            sCellData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                            sData += sCellData + "/";
                        }
                        catch (Exception ex)
                        {
                            dCellData = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                            sData += dCellData.ToString() + "/";
                        }
                    }
                    sData = sData.Remove(sData.Length - 1, 1);
                    dt.Rows.Add(sData.Split('/'));
                }

                dataGrid.ItemsSource = dt.DefaultView;

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 열기 실패! : " + ex.Message);
                return;
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        #region FocusButtom
        private void five_GotFocus(object sender, RoutedEventArgs e)
        {
            five.Text = string.Empty;
        }

        private void six_GotFocus(object sender, RoutedEventArgs e)
        {
            six.Text = string.Empty;
        }

        private void seven_GotFocus(object sender, RoutedEventArgs e)
        {
            seven.Text = String.Empty;
        }
        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var avg55 = from xp in prop
                        where xp.SelectedString1 == "ХПТ 32"
                        select xp.V1;

            decimal avg5 = 0;

            foreach (var p in avg55)
            {
                avg5 += Convert.ToDecimal(p);
            }

            MessageBox.Show($"{avg5}");
        }

        void WriteBase()
        {
            using (StreamWriter sw = new StreamWriter(path, false, System.Text.Encoding.UTF8))
            {
                int i = 0;
                foreach (var p in prop)
                {
                    sw.WriteLine(prop[i].ToString());
                    i++;
                }
            }
        }

        void AddBase(ObservableCollection<Prop> prop)
        {
            string[] b = File.ReadAllLines(path);

            foreach (var a in b)
            {
                string[] s = a.Split(';');
                string[] i = s[0].Split('|');
                switch (i.Length)
                {
                    case 2:
                        string[] k1 = i[0].Split('x');
                        string[] k2 = i[1].Split('x');
                        prop.Add(new Prop((float)Convert.ToDouble(k1[0]), (float)Convert.ToDouble(k1[1]), (float)Convert.ToDouble(k2[0]),
                                          (float)Convert.ToDouble(k2[1]), (float)Convert.ToDouble(s[1]))
                        { V2 = s[1], V3 = s[0], SelectedString1 = s[8], SelectedString2 = s[2], Pas = Massiv(s, 13), SelectedString = s[7], V1 = Massiv(s,14) });

                        break;
                    case 3:
                        string[] v = s[0].Split('|');
                        string[] t1 = v[0].Split('x');
                        string[] t2 = v[1].Split('x');
                        string[] t3 = v[2].Split('x');
                        prop.Add(new Prop(Massiv(t1, 0), Massiv(t1, 1), Massiv(t3, 0), Massiv(t3, 1),
                                          Massiv(t2, 0), Massiv(t2, 1), Massiv(s, 1))
                        {
                            V2 = s[1],
                            V3 = s[0],
                            Metri4 = Massiv(s, 3),
                            SelectedString1 = s[8],
                            SelectedString2 = s[2],
                            SelectedString6 = s[12],
                            Pas = Massiv(s, 13),
                            SelectedString = s[7],
                            V1 = Massiv(s, 14)
                        });
                        break;
                    case 4:
                        string[] v1 = s[0].Split('|');
                        string[] t11 = v1[0].Split('x');
                        string[] t22 = v1[1].Split('x');
                        string[] t33 = v1[2].Split('x');
                        string[] t44 = v1[3].Split('x');
                        prop.Add(new Prop(Massiv(t11, 0), Massiv(t11, 1), Massiv(t44, 0), Massiv(t44, 1),
                                          Massiv(t33, 0), Massiv(t33, 1), Massiv(t22, 0), Massiv(t22, 1),
                                          Massiv(s, 1))
                        {
                            V2 = s[1],
                            V3 = s[0],
                            Metri4 = Massiv(s, 3),
                            Metri3 = Massiv(s, 4),
                            SelectedString1 = s[8],
                            SelectedString2 = s[2],
                            SelectedString6 = s[12],
                            SelectedString3 = s[9],
                            Pas = Massiv(s, 13),
                            SelectedString = s[7],
                            V1 = Massiv(s, 14)
                        });
                        break;
                    case 5:
                        string[] v11 = s[0].Split('|');
                        string[] t111 = v11[0].Split('x');
                        string[] t222 = v11[1].Split('x');
                        string[] t333 = v11[2].Split('x');
                        string[] t444 = v11[3].Split('x');
                        string[] t5 = v11[3].Split('x');
                        prop.Add(new Prop(Massiv(t111, 0), Massiv(t111, 1), Massiv(t5, 0), Massiv(t5, 1),
                                          Massiv(t444, 0), Massiv(t444, 1), Massiv(t333, 0), Massiv(t333, 1),
                                          Massiv(t222, 0), Massiv(t222, 1), Massiv(s, 1))
                        {
                            V2 = s[1],
                            V3 = s[0],
                            Metri4 = Massiv(s, 3),
                            Metri3 = Massiv(s, 4),
                            Metri2 = Massiv(s, 5),
                            SelectedString1 = s[8],
                            SelectedString2 = s[2],
                            SelectedString6 = s[12],
                            SelectedString3 = s[9],
                            SelectedString4 = s[10],
                            Pas = Massiv(s, 13),
                            SelectedString = s[7],
                            V1 = Massiv(s, 14)
                        });
                        break;
                    case 6:
                        string[] v2 = s[0].Split('|');
                        string[] t12 = v2[0].Split('x');
                        string[] t23 = v2[1].Split('x');
                        string[] t34 = v2[2].Split('x');
                        string[] t45 = v2[3].Split('x');
                        string[] t55 = v2[4].Split('x');
                        string[] t6 = v2[5].Split('x');
                        prop.Add(new Prop(Massiv(t12, 0), Massiv(t12, 1), Massiv(t6, 0), Massiv(t6, 1),
                                          Massiv(t55, 0), Massiv(t55, 1), Massiv(t45, 0), Massiv(t45, 1),
                                          Massiv(t34, 0), Massiv(t34, 1), Massiv(t23, 0), Massiv(t23, 1),
                                          Massiv(s, 1))
                        {
                            V2 = s[1],
                            V3 = s[0],
                            Metri4 = Massiv(s, 3),
                            Metri3 = Massiv(s, 4),
                            Metri2 = Massiv(s, 5),
                            Metri1 = Massiv(s, 6),
                            SelectedString1 = s[8],
                            SelectedString2 = s[2],
                            SelectedString6 = s[12],
                            SelectedString3 = s[9],
                            SelectedString4 = s[10],
                            SelectedString5 = s[11],
                            Pas = Massiv(s, 13),
                            SelectedString = s[7],
                            V1 = Massiv(s, 14)
                        });
                        break;
                    default:
                        break;
                }
            }
        }

        float Massiv(string[] s, int i)
        {
            float d = (float)Convert.ToDouble(s[i]);
            return d;
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                AddBase(prop);
            }
            catch
            {
                MessageBox.Show("Нет данных");
            }
            dataGrid.Items.Refresh();
        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            WriteBase();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            prop.Clear();
            dataGrid.Items.Refresh();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            prop.Clear();
            dataGrid.Items.Refresh();
            try
            {
                AddBase(prop);
            }
            catch
            {
                MessageBox.Show("Нет данных");
            }
            dataGrid.Items.Refresh();
        }
    }
}
