using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace KFV
{
    public class MainView
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prop">Коллекция</param>
        /// <param name="worksheet"></param>
        /// <param name="i">Строка</param>
        /// <param name="g">Номер ХПТ</param>
        /// <param name="a">Колонка</param>
        public decimal AvgNORMA(ObservableCollection<Prop> prop, Worksheet worksheet, int i, string g, int a)
        {
            int z = 1;
            var avg55 = from xp in prop
                        where xp.SelectedString1 == g
                        select xp.SelectedString2;

            decimal avg5 = 0;

            foreach (var p in avg55)
            {
                avg5 += Convert.ToDecimal(p);
                z++;
            }

            z -= 1;

            if (z != 0)
            {
                worksheet.Cells[i, a] = avg5 / z;
                worksheet.Cells[i, a].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                return avg5 / z;
            }
            else
                return avg5;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="prop">Коллекция</param>
        /// <param name="worksheet"></param>
        /// <param name="i">Строка</param>
        /// <param name="g">Номер ХПТ</param>
        /// <param name="a">Колонка</param>
        public decimal Prokat(ObservableCollection<Prop> prop, Worksheet worksheet, int i, string g, int a)
        {
            var avg55 = from xp in prop
                        where xp.SelectedString1 == g
                        select xp.V2;

            decimal avg5 = 0;

            foreach (var p in avg55)
            {
                avg5 += Convert.ToDecimal(p);
            }

            worksheet.Cells[i, a] = avg5;
            worksheet.Cells[i, a].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            return avg5;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="prop">Коллекция</param>
        /// <param name="worksheet"></param>
        /// <param name="i">Строка</param>
        /// <param name="g">Номер ХПТ</param>
        /// <param name="a">Колонка</param>
        public decimal Prokat1(ObservableCollection<Prop> prop, Worksheet worksheet, int i, string g)
        {
            var avg55 = from xp in prop
                        where xp.SelectedString3 == g
                        select xp.Metri3;

            var avg551 = from xp in prop
                        where xp.SelectedString4 == g
                        select xp.Metri2;

            var avg552 = from xp in prop
                         where xp.SelectedString5 == g
                         select xp.Metri1;

            var avg553 = from xp in prop
                         where xp.SelectedString6 == g
                         select xp.Metri4;

            decimal avg5 = 0;

            foreach (var p in avg55)
            {
                avg5 += Convert.ToDecimal(p);
            }

            foreach (var p in avg551)
            {
                avg5 += Convert.ToDecimal(p);
            }

            foreach (var p in avg552)
            {
                avg5 += Convert.ToDecimal(p);
            }

            foreach (var p in avg553)
            {
                avg5 += Convert.ToDecimal(p);
            }

            return avg5;
        }

        public decimal Avg(ObservableCollection<Prop> prop, Worksheet worksheet, int i, int a)
        {
            try
            {
                var avg55 = from xp in prop
                            select Convert.ToDecimal(xp.V2) / Convert.ToDecimal(xp.SelectedString2);

                decimal avg5 = 0;

                foreach (var p in avg55)
                {
                    avg5 += Convert.ToDecimal(p);
                }

                for (int h = 0; h < prop.Count; h++)
                {
                    worksheet.Cells[h + 2, 10] = Convert.ToDecimal(prop[h].V2) / Convert.ToDecimal(prop[h].SelectedString2);
                    worksheet.Cells[h + 2, 10].Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                }

                return avg5;
            }
            catch
            {
                return 0;
            }
        }

        public decimal Afg1(ObservableCollection<Prop> prop)
        {
            var avg = from xp in prop
                      select xp.V2;

            decimal a = 0;
            foreach(var b in avg)
            {
                a += Convert.ToDecimal(b);
            }

            return a;
        }

        public decimal Afg2(ObservableCollection<Prop> prop)
        {
            var avg = from xp in prop
                      select xp.V1;

            decimal a = 0;
            foreach (var b in avg)
            {
                a += Convert.ToDecimal(b);
            }

            return a;
        }
    }
}
