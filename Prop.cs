using System;
using System.Collections.ObjectModel;

namespace KFV
{
    public class Prop : MainWindow
    {
        #region Поля и свойства
        public float pasability = 0;
        public float d1;
        public float s1;
        public float d2;
        public float s2;
        public float m1;             //КфВ
        public float m2;             //КфВ
        public float m3;            //КфВ
        public float m4;             //КфВ
        public float m5;             //КфВ

        public float marsh1;
        public float marsh2;
        public float marsh3;
        public float marsh4;
        public float marsh5;

        public float stenk1;
        public float stenk2;
        public float stenk3;
        public float stenk4;
        public float stenk5;

        public float pr1;           //прогон 1
        public float pr2;           //прогон 2
        public float pr3;           //прогон 3
        public float pr4;           //прогон 3
        public float pr5;           //прогон 3

        private float pm1;            //Прогонный метр 1
        private float pm2;            //Прогонный метр 2
        private float pm3;            //Прогонный метр 3
        private float pm4;            //Прогонный метр 3
        private float pm5;            //Прогонный метр 3

        public float dubl;
        public int a;

        public float Metri1 { get; set; }
        public float Metri2 { get; set; }
        public float Metri3 { get; set; }
        public float Metri4 { get; set; }

        public float Marsh1 { get; set; }
        public float Marsh2 { get; set; }
        public float Marsh3 { get; set; }
        public float Marsh4 { get; set; }
        public float Marsh5 { get; set; }

        public float Stenk1 { get; set; }
        public float Stenk2 { get; set; }
        public float Stenk3 { get; set; }
        public float Stenk4 { get; set; }
        public float Stenk5 { get; set; }

        public float Dia
        {
            get { return d1; }
            set { d1 = value; }
        }

        public float Sten
        {
            get { return s1; }
            set { s1 = value; }
        }

        public float Diam
        {
            get { return d2; }
            set { d2 = value; }
        }

        public float Stenk
        {
            get { return s2; }
            set { s2 = value; }
        }
        public float Pas { get; set; }
        public float V1 { get; set; }
        public string V2 { get; set; }
        public string V3 { get; set; }
        public string V4 { get; set; }
        public string V5 { get; set; }
        public string V6 { get; set; }
        public string V8 { get; set; }
        #endregion

        #region ComboBox
        private ObservableCollection<string> _testList1 = new ObservableCollection<string> { "ХПТ 32", "ХПТ 55", "ХПТ 55-2", "ХПТ 75", "ХПТ 90" };
        public ObservableCollection<string> TestList1
        {
            get
            {
                return _testList1;
            }
        }

        private string _ss1;
        public string SelectedString1
        {
            get
            {
                return _ss1;
            }
            set
            {
                _ss1 = value;
            }
        }

        private ObservableCollection<string> _testList = new ObservableCollection<string> { "КОТ", "ОН гот",
                                                                                            "ОН заг", "ОТ гот",
                                                                                            "ОТ заг","ЭХП гот", "ЭХП заг", "ОТ заг", "БР заг", "БР гот" };
        public ObservableCollection<string> TestList
        {
            get
            {
                return _testList;
            }
        }

        private string _ss;
        public string SelectedString
        {
            get
            {
                return _ss;
            }
            set
            {
                _ss = value;
            }
        }

        private ObservableCollection<string> _testList2 = new ObservableCollection<string> { "400", "500",
                                                                                            "600", "700",
                                                                                            "800", "900", "1000", "1100", "1200", "1300", "1400", "1500", "1600", "1700", "1800", "1900", "2000" };
        public ObservableCollection<string> TestList2
        {
            get
            {
                return _testList2;
            }
        }

        private string _ss2;
        public string SelectedString2
        {
            get
            {
                return _ss2;
            }
            set
            {
                _ss2 = value;
            }
        }

        private ObservableCollection<string> _testList3 = new ObservableCollection<string> { "ХПТ 55", "ХПТ 75", "ХПТ 90" };
        public ObservableCollection<string> TestList3
        {
            get
            {
                return _testList3;
            }
        }

        private string _ss3;
        public string SelectedString3
        {
            get
            {
                return _ss3;
            }
            set
            {
                _ss3 = value;
            }
        }
        private ObservableCollection<string> _testList4 = new ObservableCollection<string> {  "ХПТ 55", "ХПТ 75", "ХПТ 90" };
        public ObservableCollection<string> TestList4
        {
            get
            {
                return _testList4;
            }
        }

        private string _ss4;
        public string SelectedString4
        {
            get
            {
                return _ss4;
            }
            set
            {
                _ss4 = value;
            }
        }

        private ObservableCollection<string> _testList5 = new ObservableCollection<string> {"ХПТ 55", "ХПТ 75", "ХПТ 90" };
        public ObservableCollection<string> TestList5
        {
            get
            {
                return _testList5;
            }
        }

        private string _ss5;
        public string SelectedString5
        {
            get
            {
                return _ss5;
            }
            set
            {
                _ss5 = value;
            }
        }
        private ObservableCollection<string> _testList6 = new ObservableCollection<string> { "ХПТ 55", "ХПТ 75", "ХПТ 90" };
        public ObservableCollection<string> TestList6
        {
            get
            {
                return _testList6;
            }
        }

        private string _ss6;
        public string SelectedString6
        {
            get
            {
                return _ss6;
            }
            set
            {
                _ss6 = value;
            }
        }
        #endregion

        #region Перегрузки конструкторов
        public Prop() {
            Dia = default;
            Sten = default;
            Diam = default;
            Stenk = default;
            Pas = default;
            V1 = default;
            V2 = default;
            V3 = default;
            V4 = default;
            V5 = default;
            V6 = default;
            SelectedString2 = default;
            SelectedString = default;
            SelectedString1 = default;
            SelectedString3 = default;
            SelectedString4 = default;
            Metri1 = default;
            Metri2 = default;
            Metri3 = default;
            Metri4 = default;
        }

        public Prop(float d1, float s1, float d2, float s2, float m)
        {
            Dia = d1;
            Sten = s1;
            Diam = d2;
            Stenk = s2;
            Pas = Meybe();
            V1 = GetChill(m);
            V2 = "";
            V3 = "";
            V4 = $"{d1}x{s1}";
            V5 = $"{d2}x{s2}";
            V6 = "";
            SelectedString2 = "";
            SelectedString = "";
            SelectedString1 = "";
            SelectedString3 = "";
            SelectedString4 = "";
            Metri1 = default;
            Metri2 = default;
            Metri3 = default;
            Metri4 = default;
        }

        public float GetChill(float a)
        {
            float f = pasability;
            float n = f * a;
            return n;
        }

        /// <summary> ++++++++++++
        /// 
        /// </summary>
        /// <param name="d1">76</param>
        /// <param name="s1">8,1</param>
        /// <param name="d2">20</param>
        /// <param name="s2">2</param>
        /// <param name="marsh1">45</param>
        /// <param name="stenk1">4,1</param>
        public Prop(float d1, float s1, float d2, float s2,
            float marsh1, float stenk1, float m)
        {
            Dia = d1;
            Sten = s1;
            Diam = d2;
            Stenk = s2;
            Pas = Meybe(d2, s2, marsh1, stenk1,m);
            V1 = GetChill(m);
            V2 = "";
            V3 = "";
            V4 = $"{d1}x{s1}";
            V5 = $"{d2}x{s2}";
            V6 = "";
            SelectedString2 = "";
            SelectedString = "";
            SelectedString1 = "";
            SelectedString3 = "";
            SelectedString4 = "";
            Metri4 = m / pr2;
            Metri2 = default;
            Metri3 = default;
            Metri1 = default;
        }

        /// <summary> ++++++++++
        /// 
        /// </summary>
        /// <param name="marsh1">20</param>
        /// <param name="stenk1">2</param>
        /// <param name="marsh2">45</param>
        /// <param name="stenk2">4,1</param>
        /// <returns></returns>
        public float Meybe(float marsh1, float stenk1, float marsh2, float stenk2, float m)
        {
            Big2(marsh2, stenk2, marsh1, stenk1);
            pm1 = 1000 / pr2;
            return pasability = (pm1 + 1000) / 1000;
        }

        /// <summary> +++++++++
        /// 
        /// </summary>
        /// <param name="d1">102</param>
        /// <param name="s1">11</param>
        /// <param name="d2">20</param>
        /// <param name="s2">2</param>
        /// <param name="marsh1">76</param>
        /// <param name="stenk1">8,1</param>
        /// <param name="marsh2">45</param>
        /// <param name="stenk2">4,1</param>
        public Prop(float d1, float s1, float d2, float s2, float marsh1, float stenk1,
            float marsh2, float stenk2, float m)
        {
            Dia = d1;
            Sten = s1;
            Diam = d2;
            Stenk = s2;
            Pas = Meybe(d2, s2, marsh1, stenk1, marsh2, stenk2, m);
            V1 = GetChill(m);
            V2 = "";
            V3 = "";
            V4 = $"{d1}x{s1}";
            V5 = $"{d2}x{s2}";
            V6 = "";
            SelectedString2 = "";
            SelectedString = "";
            SelectedString1 = "";
            SelectedString3 = "";
            SelectedString4 = "";
            Metri4 = m / pr2;
            Metri3 = Metri4 / pr3;
            Metri1 = default;
            Metri2 = default;
        }

        /// <summary> ++++++++++++++++++++
        /// 
        /// </summary>
        /// <param name="marsh1">76</param>
        /// <param name="stenk1">8,1</param>
        /// <param name="marsh2">45</param>
        /// <param name="stenk2">4,5</param>
        /// <param name="marsh3">20</param>
        /// <param name="stenk3">1</param>
        /// <returns></returns>
        public float Meybe(float marsh1, float stenk1, float marsh2, float stenk2,
                           float marsh3, float stenk3, float m)
        {
            Big2(marsh2, stenk2, marsh1, stenk1);
            Big3(marsh3, stenk3, marsh2, stenk2);
            pm1 = 1000 / pr2;
            pm2 = pm1 / pr3;
            return pasability = (pm1 + pm2 + 1000) / 1000;

        }

        /// <summary> ++++++++++++
        /// 
        /// </summary>
        /// <param name="d1">140</param>
        /// <param name="s1">15</param>
        /// <param name="d2">20</param>
        /// <param name="s2">2</param>
        /// <param name="marsh1">102</param>
        /// <param name="stenk1">11</param>
        /// <param name="marsh2">76</param>
        /// <param name="stenk2">8.1</param>
        /// <param name="marsh3">45</param>
        /// <param name="stenk3">4.1</param>
        public Prop(float d1, float s1, float d2, float s2, float marsh1, float stenk1,
            float marsh2, float stenk2, float marsh3, float stenk3, float m)
        {
            Dia = d1;
            Sten = s1;
            Diam = d2;
            Stenk = s2;
            Pas = Meybe(d2, s2, marsh3, stenk3, marsh2, stenk2, marsh1, stenk1,m);
            V1 = GetChill(m);
            V2 = "";
            V3 = "";
            V4 = $"{d1}x{s1}";
            V5 = $"{d2}x{s2}";
            V6 = "";
            SelectedString2 = "";
            SelectedString = "";
            SelectedString1 = "";
            SelectedString3 = "";
            SelectedString4 = "";
            Metri2 = m / pr2;
            Metri3= Metri2 / pr3;
            Metri4 = Metri3 / pr4;
            Metri1 = default;
        }

        /// <summary> +++++++++++++++
        /// 
        /// </summary>
        /// <param name="marsh1">20</param>
        /// <param name="stenk1">2</param>
        /// <param name="marsh2">45</param>
        /// <param name="stenk2">4.1</param>
        /// <param name="marsh3">75</param>
        /// <param name="stenk3">8,1</param>
        /// <param name="marsh4">102</param>
        /// <param name="stenk4">11</param>
        /// <returns></returns>
        public float Meybe(float marsh1, float stenk1, float marsh2, float stenk2,
                           float marsh3, float stenk3, float marsh4, float stenk4, float m)
        {
            Big2(marsh2, stenk2, marsh1, stenk1);
            Big3(marsh3, stenk3, marsh2, stenk2);
            Big4(marsh4, stenk4, marsh3, stenk3);
            pm1 = 1000 / pr2;
            pm2 = pm1 / pr3;
            pm3 = pm2 / pr4;
            return pasability = (pm1 + pm2 + pm3 + 1000) / 1000;
        }

        /// <summary> ++++++++++++
        /// 
        /// </summary>
        /// <param name="d1">140</param>
        /// <param name="s1">15</param>
        /// <param name="d2">20</param>
        /// <param name="s2">2</param>
        /// <param name="marsh1">102</param>
        /// <param name="stenk1">11</param>
        /// <param name="marsh2">76</param>
        /// <param name="stenk2">8.1</param>
        /// <param name="marsh3">45</param>
        /// <param name="stenk3">4.1</param>
        public Prop(float d1, float s1, float d2, float s2, float marsh1, float stenk1,
            float marsh2, float stenk2, float marsh3, float stenk3, float marsh4, float stenk4, float m)
        {
            Dia = d1;
            Sten = s1;
            Diam = d2;
            Stenk = s2;
            Pas = Meybe(d2, s2, marsh4, stenk4, marsh3, stenk3, marsh2, stenk2, marsh1, stenk1, m);
            V1 = GetChill(m);
            V2 = "";
            V3 = "";
            V4 = $"{d1}x{s1}";
            V5 = $"{d2}x{s2}";
            V6 = "";
            SelectedString2 = "";
            SelectedString = "";
            SelectedString1 = "";
            SelectedString3 = "";
            SelectedString4 = ""; 
            Metri1 = m / pr2;
            Metri2 = Metri1 / pr3;
            Metri3 = Metri2 / pr4;
            Metri4 = Metri3 / pr5;
        }

        /// <summary> +++++++++++++++
        /// 
        /// </summary>
        /// <param name="marsh1">20</param>
        /// <param name="stenk1">2</param>
        /// <param name="marsh2">45</param>
        /// <param name="stenk2">4.1</param>
        /// <param name="marsh3">75</param>
        /// <param name="stenk3">8,1</param>
        /// <param name="marsh4">102</param>
        /// <param name="stenk4">11</param>
        /// <returns></returns>
        public float Meybe(float marsh1, float stenk1, float marsh2, float stenk2,
                           float marsh3, float stenk3, float marsh4, float stenk4, float marsh5, float stenk5, float m)
        {
            Big2(marsh2, stenk2, marsh1, stenk1);
            Big3(marsh3, stenk3, marsh2, stenk2);
            Big4(marsh4, stenk4, marsh3, stenk3);
            Big5(marsh5, stenk5, marsh4, stenk4);
            pm1 = 1000 / pr2;
            pm2 = pm1 / pr3;
            pm3 = pm2 / pr4;
            pm4 = pm3 / pr5;
            return pasability = (pm1 + pm2 + pm3+ pm4 + 1000) / 1000;
        }

        public float Big1(float d1, float s1, float d2, float s2)
        {
            return pasability = 1;
        }
        public float Big2(float d1, float s1, float d2, float s2)
        {
            pr2 = ((d1 - s1) * s1) / ((d2 - s2) * s2);
            return pr2;
        }
        public float Big3(float d1, float s1, float d2, float s2)
        {
            pr3 = ((d1 - s1) * s1) / ((d2 - s2) * s2);
            return pr3;
        }
        public float Big4(float d1, float s1, float d2, float s2)
        {
            pr4 = ((d1 - s1) * s1) / ((d2 - s2) * s2);
            return pr4;
        }
        public float Big5(float d1, float s1, float d2, float s2)
        {
            pr5 = ((d1 - s1) * s1) / ((d2 - s2) * s2);
            return pr5;
        }
        /// <summary>
        /// Точка входа
        /// </summary>
        public float Meybe()
        {
            return pasability = 1;
        }

        #endregion

        public override string ToString()
        {
            return $"{V3};{V2};{SelectedString2};{Metri4};{Metri3};{Metri2};" +
                   $"{Metri1};{SelectedString};{SelectedString1};{SelectedString3};{SelectedString4};{SelectedString5};" +
                   $"{SelectedString6};{Pas};{V1}";
        }

    }
}
