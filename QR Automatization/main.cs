using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using AForge.Video;
using AForge.Video.DirectShow;
using ZXing;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Numerics;
using System.Diagnostics;

namespace QR_Automatization
{
    public partial class main : MetroFramework.Forms.MetroForm
    {
        readonly DataTable dt = new DataTable();
        readonly DataTable dt2 = new DataTable();
        readonly ListBox framerate = new ListBox();

        OleDbDataAdapter da2;
        DataSet ds2;

        OleDbDataAdapter da3;
        DataSet ds3;

        private FilterInfoCollection Devices;
        private VideoCaptureDevice Source;

        static readonly string[] Scopes = { DriveService.Scope.Drive };
        static readonly string ApplicationName = "Mersin Uni_Order Automatization with QR Code";

        readonly SaveFileDialog save = new SaveFileDialog();

        string filePath, filePath2, filePath3, filePath4;
        public TextBox txt = new TextBox();

        int a = 0;
        int number = 1;
        int t = 0;
        int x = 0;
        int c = 0;
        int adet = 0;
        int gör = 0;
        public main()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        void fillgrid()
        {
            OleDbConnection con;

            con = new OleDbConnection("Provider = Microsoft.ACE.Oledb.12.0; Data Source = ta.accdb");
            da2 = new OleDbDataAdapter("SElect * from hir", con);
            ds2 = new DataSet();
            con.Open();
            da2.Fill(ds2, "ta");
            accessgrid.DataSource = ds2.Tables["ta"];
            con.Close();
        }
        void controlfillgrid()
        {
            OleDbConnection con;

            con = new OleDbConnection("Provider = Microsoft.ACE.Oledb.12.0; Data Source = ta.accdb");
            da3 = new OleDbDataAdapter("SElect * from sira", con);
            ds3 = new DataSet();
            con.Open();
            da3.Fill(ds3, "ta");
            sequenceaccess.DataSource = ds3.Tables["ta"];
            con.Close();
        }
        private void killzombieexcel(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "" + excelFileName)
                {
                    process.Kill();
                    process.Dispose();
                }
            }
        }
        void importexcel()
        {
            var processes = Process.GetProcessesByName("EXCEL");

            foreach (var p in processes)
            {
                if (p != Process.GetCurrentProcess())
                {
                    lstpid_first.Items.Add(p.Id.ToString());
                }

            }

            Excel.Workbook oWB = null;
            Excel.Application oXL = null;

            oXL = new Excel.Application();

            var processes2 = Process.GetProcessesByName("EXCEL");

            foreach (var p2 in processes2)
            {
                if (p2 != Process.GetCurrentProcess())
                {
                    lstpid_last.Items.Add(p2.Id.ToString());
                }
            }

            if (workbooktext.Text == string.Empty)
            {
                return;
            }
            else
            {
                if (preparedgrid.Rows.Count != 0)
                {
                    preparedgrid.Rows.Clear();
                }

                if (outputgrid.Rows.Count != 0)
                {
                    dt.Clear();
                }

                oWB = oXL.Workbooks.Open(workbooktext.Text);

                List<string> liste = new List<string>();
                foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                {
                    liste.Add(oSheet.Name);
                }
                pseudogrid.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                excelfiletext.Text = pseudogrid.Rows[0].Cells[0].Value.ToString();

                OleDbCommand komut = new OleDbCommand();
                string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + workbooktext.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + excelfiletext.Text + "$]", conn);
                MyDataAdapter.Fill(dt);
                outputgrid.DataSource = dt;
                oWB.Close(0);
                oXL.Quit();
                GC.Collect();

                creatingnewexcel();

                substringtext.Text = "";

                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                oWB = null;
                oXL = null;

                for (int g = 0; g < lstpid_last.Items.Count; g++)
                {
                    for (int f = 0; f < lstpid_first.Items.Count; f++)
                    {
                        if (lstpid_last.Items[g].ToString() == lstpid_first.Items[f].ToString())
                        {
                            var processes3 = Process.GetProcessesByName("EXCEL");

                            foreach (var p3 in processes3)
                            {
                                if (p3.Id != Convert.ToInt32(lstpid_last.Items[g].ToString()))
                                {
                                    p3.Dispose();
                                }
                            }
                        }
                    }
                }

                if (lstpid_first.Items.Count == 0)
                {
                    killzombieexcel(lblfilename.Text);
                }
                collectgarbage();
                liste.Clear();
            }
        }
        void creatingnewexcel()
        {
            //upload edilecek excel sütunları oluşturma
            #region
            preparedgrid.ColumnCount = 21;

            preparedgrid.Columns[0].Name = "Flavor Abbr.";
            preparedgrid.Columns[1].Name = "labels";
            preparedgrid.Columns[2].Name = "slider";
            preparedgrid.Columns[3].Name = "ISI";
            preparedgrid.Columns[4].Name = "autocontinue";
            preparedgrid.Columns[5].Name = "ratings.thisRepN";
            preparedgrid.Columns[6].Name = "ratings.thisTrialN";
            preparedgrid.Columns[7].Name = "ratings.thisN";
            preparedgrid.Columns[8].Name = "ratings.thisIndex";
            preparedgrid.Columns[9].Name = "Rating_raw";
            preparedgrid.Columns[10].Name = "Rating_trans";
            preparedgrid.Columns[11].Name = "Rating_RT1";
            preparedgrid.Columns[12].Name = "ParticipantNumber";
            preparedgrid.Columns[13].Name = "SubjectNumber";
            preparedgrid.Columns[14].Name = "SessionNumber";
            preparedgrid.Columns[15].Name = "Condition";
            preparedgrid.Columns[16].Name = "InputFile";
            preparedgrid.Columns[17].Name = "Experimenter";
            preparedgrid.Columns[18].Name = "date";
            preparedgrid.Columns[19].Name = "experiment";
            preparedgrid.Columns[20].Name = "frameRate";
            #endregion

            //boş kalan değerlerin belirleme
            #region
            for (int o = 0; o < 7; o++)
            {
                lstvoidline1.Items.Add(o);
            }
            for (int r = 19; r < outputgrid.Rows.Count; r += 14)
            {
                if (r == 18)
                {

                }
                else
                {
                    lstvoidline1.Items.Add(r);
                }
            }
            for (int j = 20; j < outputgrid.Rows.Count; j += 14)
            {
                lstvoidline1.Items.Add(j);
            }

            for (int z = 0; z < lstvoidline1.Items.Count; z++)
            {
                x = Convert.ToInt32(lstvoidline1.Items[z]);
            }

            for (int i = 0; i < outputgrid.Rows.Count; i++)
            {
                if (outputgrid.Rows[i].Cells[19].Value.ToString() == "")
                {
                    gör++;
                    if (gör == 1)
                    {
                        nullrowtext.Text = i.ToString();
                    }
                }
            }

            if (nullrowtext.Text == "")
            {
                int b = outputgrid.Rows.Count;
                int cikan = b - x;
                for (int u = 0; u < cikan; u++)
                {
                    x++;
                    if (x <= b)
                    {
                        lstvoidline1.Items.Add(x);
                    }
                }
            }
            else
            {
                int b = Convert.ToInt32(nullrowtext.Text.ToString());

                for (int ç = 0; ç < 6; ç++)
                {
                    lstvoidline1.Items.Add(b);
                    b--;
                }
            }
            #endregion

            //framerate logaritmik hata düzeltme
            #region
            filePath = workbooktext.Text;
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            int hh = 0;
            while (excelReader.Read())
            {
                hh++;
                if (hh == 1)
                {

                }
                else
                {
                    if (excelReader.GetValue(19) == null)
                    {
                        lstframerate.Items.Add("");
                    }
                    else
                    {
                        lstframerate.Items.Add(excelReader.GetValue(19));
                    }
                }
            }
            excelReader.Close();
            stream.Close();

            for (int m = 0; m < lstframerate.Items.Count; m++)
            {
                string framerate = lstframerate.Items[m].ToString();
                if (framerate.Contains("E"))
                {
                    logarithmictext.Text = lstframerate.Items[m].ToString();

                    int deger_e, adet_virgul, adet_sifir;
                    string detectseperator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString();

                    string virgül = logarithmictext.Text.Replace(detectseperator, ""); //virgül çıkartıldı (bunun nokta olması gerekiyor ama neden? hayır, kültüre göre separator neyse onun gelmesi gerekiyor)
                    int index = virgül.IndexOf("E"); //E'nin indexsi bulduk
                    string e = virgül.Substring(index); //e'den sonrasını aldık //değer e'li gelmiyor ki
                    string ana = virgül.Replace(e, ""); //ve çıkarttık, sadece ana sayı kaldı şu an

                    string[] sayilar = Regex.Split(e, @"\D+");
                    foreach (string s in sayilar)
                    {
                        if (int.TryParse(s, out _))
                        {
                            int modlanacak = Convert.ToInt32(s);
                            lstnumbere.Items.Add(modlanacak);
                        }
                    }

                    BigInteger kontrol = BigInteger.Parse(ana);
                    if (kontrol < 0)
                    {
                        deger_e = Convert.ToInt32(lstnumbere.Items[c]); //E+15'daki 15
                        adet_virgul = Convert.ToInt32(ana.Length - 2); //mesela 13 bsm.'de 12 virgül koyabilirsin
                        adet_sifir = deger_e - adet_virgul; //kaç tane sıfır gelecek sona onun belirlenmesi
                    }
                    else
                    {
                        deger_e = Convert.ToInt32(lstnumbere.Items[c]); //E+15'daki 15
                        adet_virgul = Convert.ToInt32(ana.Length - 1); //mesela 13 bsm.'de 12 virgül koyabilirsin
                        adet_sifir = deger_e - adet_virgul; //kaç tane sıfır gelecek sona onun belirlenmesi
                    }

                    if (adet_sifir < 0)
                    {
                        Math.Abs(adet_sifir);
                    }

                    for (int k = 0; k < adet_sifir; k++)
                    {
                        string z = ana + "0";
                        BigInteger buyukdeger = BigInteger.Parse(z); //int olmuyor buraya çünkü değer çok fazla oluyor
                        ana = Convert.ToString(buyukdeger);
                    }

                    lstframerate.Items[m] = ana.ToString();
                    c++;
                }
                else
                {

                }
            }
            #endregion

            logarithmictext.Text = "";
            lstnumbere.Items.Clear();
            c = 0;

            //rating_raw logaritmik hata düzeltme
            #region
            filePath2 = workbooktext.Text;
            FileStream stream2 = File.Open(filePath2, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader2 = ExcelReaderFactory.CreateOpenXmlReader(stream2);
            int h = 0;
            while (excelReader2.Read())
            {
                h++;
                if (h == 1)
                {

                }
                else
                {
                    if (excelReader2.GetValue(8) == null)
                    {
                        lstrtraw.Items.Add("");
                    }
                    else
                    {
                        lstrtraw.Items.Add(excelReader2.GetValue(8));
                    }
                }
            }
            excelReader2.Close();
            stream2.Close();

            for (int k = 0; k < lstrtraw.Items.Count; k++)
            {
                string rating_raw = lstrtraw.Items[k].ToString();
                if (rating_raw.Contains("E"))
                {
                    logarithmictext.Text = lstrtraw.Items[k].ToString();

                    int deger_e, adet_virgul, adet_sifir;
                    string detectseperator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString();

                    string virgül = logarithmictext.Text.Replace(detectseperator, "");
                    int index = virgül.IndexOf("E");
                    string e = virgül.Substring(index);
                    string ana = virgül.Replace(e, "");

                    string[] sayilar = Regex.Split(e, @"\D+");
                    foreach (string s in sayilar)
                    {
                        if (int.TryParse(s, out _))
                        {
                            int modlanacak = Convert.ToInt32(s);
                            lstnumbere.Items.Add(modlanacak);
                        }
                    }

                    BigInteger kontrol = BigInteger.Parse(ana);
                    if (kontrol < 0)
                    {
                        deger_e = Convert.ToInt32(lstnumbere.Items[c]);
                        adet_virgul = Convert.ToInt32(ana.Length - 2);
                        adet_sifir = deger_e - adet_virgul;
                    }
                    else
                    {
                        deger_e = Convert.ToInt32(lstnumbere.Items[c]);
                        adet_virgul = Convert.ToInt32(ana.Length - 1);
                        adet_sifir = deger_e - adet_virgul;
                    }

                    if (adet_sifir < 0)
                    {
                        Math.Abs(adet_sifir);
                    }

                    for (int v = 0; v < adet_sifir; v++)
                    {
                        string z = ana + "0";
                        BigInteger buyukdeger = BigInteger.Parse(z);
                        ana = Convert.ToString(buyukdeger);
                    }

                    lstrtraw.Items[k] = ana.ToString();
                    c++;
                }
                else
                {

                }
            }
            #endregion

            //rating_trans tarihli değer düzeltme
            #region
            filePath3 = workbooktext.Text;
            FileStream stream3 = File.Open(filePath3, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader3 = ExcelReaderFactory.CreateOpenXmlReader(stream3);
            int p = 0;
            while (excelReader3.Read())
            {
                p++;
                if (p == 1)
                {

                }
                else
                {
                    if (excelReader3.GetValue(9) == null)
                    {
                        lsttrans.Items.Add("");
                    }
                    else
                    {
                        lsttrans.Items.Add(excelReader3.GetValue(9));
                    }
                }
            }
            excelReader3.Close();
            stream3.Close();

            for (int w = 0; w < lsttrans.Items.Count; w++)
            {
                string rating_trans = lsttrans.Items[w].ToString();

                if (rating_trans.Contains(".") || rating_trans.Contains(":"))
                {
                    char harf1 = char.Parse(".");
                    char harf2 = char.Parse(":");
                    for (int i = 0; i < rating_trans.Length; i++)
                    {
                        if (rating_trans[i] == harf1)
                        {
                            adet++;
                        }
                    }

                    for (int j = 0; j < rating_trans.Length; j++)
                    {
                        if (rating_trans[j] == harf2)
                        {
                            adet++;
                        }
                    }

                    if (adet == 4)
                    {
                        string tarih = rating_trans.Substring(0, rating_trans.Length - 9); //1.01.2021 gibi oldu şu an, sondan aldık bunları
                        string degerimiz = tarih.Substring(0, tarih.Length - 5); //1.01 gibi oldu şu an, sondan sildik yılları
                        lsttrans.Items[w] = degerimiz.ToString(); //vee mutlu son :))
                    }
                    else
                    {

                    }

                    adet = 0;
                }
                else
                {

                }
            }
            #endregion

            //rating rt1 tarihli değer düzeltme
            #region
            filePath4 = workbooktext.Text;
            FileStream stream4 = File.Open(filePath4, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader4 = ExcelReaderFactory.CreateOpenXmlReader(stream4);
            int q = 0;
            while (excelReader4.Read())
            {
                q++;
                if (q == 1)
                {

                }
                else
                {
                    if (excelReader4.GetValue(10) == null)
                    {
                        lstrt1.Items.Add("");
                    }
                    else
                    {
                        lstrt1.Items.Add(excelReader4.GetValue(10));
                    }
                }
            }
            excelReader4.Close();
            stream4.Close();

            for (int l = 0; l < lstrt1.Items.Count; l++)
            {
                string rating_rt1 = lstrt1.Items[l].ToString();
                if (rating_rt1.Contains(".") || rating_rt1.Contains(":") || rating_rt1.Contains(","))
                {
                    char harf1 = char.Parse(".");
                    char harf2 = char.Parse(":");
                    for (int i = 0; i < rating_rt1.Length; i++)
                    {
                        if (rating_rt1[i] == harf1)
                        {
                            adet++;
                        }
                    }
                    for (int j = 0; j < rating_rt1.Length; j++)
                    {
                        if (rating_rt1[j] == harf2)
                        {
                            adet++;
                        }
                    }

                    if (adet == 4)
                    {
                        string tarih = rating_rt1.Substring(0, rating_rt1.Length - 9); //1.01.2021 gibi oldu şu an
                        string degerimiz = tarih.Substring(0, tarih.Length - 5); //1.01 gibi oldu şu an

                        int index = degerimiz.IndexOf(".");
                        string ondalik = degerimiz.Substring(index + 1);

                        int sayi_bsm = Convert.ToInt32(ondalik.ToString());
                        int sayac = 0;
                        while (sayi_bsm > 0)
                        {
                            sayi_bsm /= 10;
                            sayac++;
                        }
                        if (ondalik.StartsWith("0") == true)
                        {
                            sayac++;

                            string tam = degerimiz.Substring(0, degerimiz.Length - sayac - 1);
                            string sayi = tam + ondalik;
                            int y = Convert.ToInt32(sayi);
                            lstrt1.Items[l] = y;
                        }
                        else
                        {
                            string tam = degerimiz.Substring(0, degerimiz.Length - sayac - 1);
                            string sayi = tam + ondalik;
                            int y = Convert.ToInt32(sayi);
                            lstrt1.Items[l] = y;
                        }
                    }
                    else
                    {
                        string detectseperator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString();

                        if (rating_rt1.Contains("/") == true)
                        {
                            int index2 = rating_rt1.IndexOf("/");
                            string ondalik2 = rating_rt1.Substring(index2 + 1, 2);

                            int sayi_bsma = Convert.ToInt32(ondalik2.ToString());
                            int sayaca = 0;
                            while (sayi_bsma > 0)
                            {
                                sayi_bsma /= 10;
                                sayaca++;
                            }

                            if (ondalik2.StartsWith("0") == true)
                            {
                                string tam = rating_rt1.Substring(1, 1);
                                string sayi = tam + ondalik2;
                                lstrt1.Items[l] = sayi;
                            }
                            else //12
                            {
                                string tam = rating_rt1.Substring(1,2);
                                string sayi = tam + ondalik2;
                                lstrt1.Items[l] = sayi;
                            }
                        }
                        else
                        {
                            int index2 = rating_rt1.IndexOf(detectseperator);
                            string ondalik2 = rating_rt1.Substring(index2 + 1);

                            int sayi_bsma = Convert.ToInt32(ondalik2.ToString());
                            int sayaca = 0;
                            while (sayi_bsma > 0)
                            {
                                sayi_bsma /= 10;
                                sayaca++;
                            }

                            if (ondalik2.StartsWith("0") == true)
                            {
                                sayaca++;

                                string tam = rating_rt1.Substring(0, rating_rt1.Length - sayaca - 1);
                                string sayi = tam + ondalik2;
                                int y = Convert.ToInt32(sayi);
                                lstrt1.Items[l] = y;
                            }
                            else
                            {
                                string tam = rating_rt1.Substring(0, rating_rt1.Length - sayaca - 1);
                                string sayi = tam + ondalik2;
                                int y = Convert.ToInt32(sayi);
                                lstrt1.Items[l] = y;
                            }
                        }
                    }

                    adet = 0;
                }
                else
                {

                }
            }
            #endregion
        }
        void writeexceltogrid()
        {
            //upload edilecek excele değerleri yazdırma
            #region
            ArrayList list = new ArrayList();
            foreach (object o in lstvoidline1.Items)
            {
                list.Add(o);
            }
            list.Sort();
            lstvoidline1.Items.Clear();
            foreach (object o in list)
            {
                lstvoidline1.Items.Add(o);
            }

            for (int i = 0; i <= outputgrid.Rows.Count - 1; i++)
            {
                if (lstvoidline1.Items.Contains(i) == true)
                {
                    this.preparedgrid.Rows.Add("", outputgrid.Rows[i].Cells[0].Value.ToString(), outputgrid.Rows[i].Cells[1].Value.ToString(), outputgrid.Rows[i].Cells[2].Value.ToString(), outputgrid.Rows[i].Cells[3].Value.ToString(), outputgrid.Rows[i].Cells[4].Value.ToString(), outputgrid.Rows[i].Cells[5].Value.ToString(), outputgrid.Rows[i].Cells[6].Value.ToString(), outputgrid.Rows[i].Cells[7].Value.ToString(), lstrtraw.Items[i].ToString(), lsttrans.Items[i].ToString(), lstrt1.Items[i].ToString(), outputgrid.Rows[i].Cells[11].Value.ToString(), outputgrid.Rows[i].Cells[12].Value.ToString(), outputgrid.Rows[i].Cells[13].Value.ToString(), outputgrid.Rows[i].Cells[14].Value.ToString(), outputgrid.Rows[i].Cells[15].Value.ToString(), outputgrid.Rows[i].Cells[16].Value.ToString(), outputgrid.Rows[i].Cells[17].Value.ToString(), outputgrid.Rows[i].Cells[18].Value.ToString(), lstframerate.Items[i].ToString());
                }
                else
                {
                    if (t == 13)
                    {
                        a++;
                        if (a >= 30)
                        {

                        }
                        else
                        {
                            lstvoidline2.Items.Add(i.ToString());
                            this.preparedgrid.Rows.Add(lstheadnumber.Items[a].ToString() + number, outputgrid.Rows[i].Cells[0].Value.ToString(), outputgrid.Rows[i].Cells[1].Value.ToString(), outputgrid.Rows[i].Cells[2].Value.ToString(), outputgrid.Rows[i].Cells[3].Value.ToString(), outputgrid.Rows[i].Cells[4].Value.ToString(), outputgrid.Rows[i].Cells[5].Value.ToString(), outputgrid.Rows[i].Cells[6].Value.ToString(), outputgrid.Rows[i].Cells[7].Value.ToString(), lstrtraw.Items[i].ToString(), lsttrans.Items[i].ToString(), lstrt1.Items[i].ToString(), outputgrid.Rows[i].Cells[11].Value.ToString(), outputgrid.Rows[i].Cells[12].Value.ToString(), outputgrid.Rows[i].Cells[13].Value.ToString(), outputgrid.Rows[i].Cells[14].Value.ToString(), outputgrid.Rows[i].Cells[15].Value.ToString(), outputgrid.Rows[i].Cells[16].Value.ToString(), outputgrid.Rows[i].Cells[17].Value.ToString(), outputgrid.Rows[i].Cells[18].Value.ToString(), lstframerate.Items[i].ToString());
                            t = 0;
                            number = 0;
                        }
                    }

                    else if (t == 12)
                    {
                        a++;
                        if (a >= 30)
                        {

                        }
                        else
                        {
                            number = 1;
                            lstvoidline2.Items.Add(i.ToString());
                            this.preparedgrid.Rows.Add(lstheadnumber.Items[a].ToString() + number, outputgrid.Rows[i].Cells[0].Value.ToString(), outputgrid.Rows[i].Cells[1].Value.ToString(), outputgrid.Rows[i].Cells[2].Value.ToString(), outputgrid.Rows[i].Cells[3].Value.ToString(), outputgrid.Rows[i].Cells[4].Value.ToString(), outputgrid.Rows[i].Cells[5].Value.ToString(), outputgrid.Rows[i].Cells[6].Value.ToString(), outputgrid.Rows[i].Cells[7].Value.ToString(), lstrtraw.Items[i].ToString(), lsttrans.Items[i].ToString(), lstrt1.Items[i].ToString(), outputgrid.Rows[i].Cells[11].Value.ToString(), outputgrid.Rows[i].Cells[12].Value.ToString(), outputgrid.Rows[i].Cells[13].Value.ToString(), outputgrid.Rows[i].Cells[14].Value.ToString(), outputgrid.Rows[i].Cells[15].Value.ToString(), outputgrid.Rows[i].Cells[16].Value.ToString(), outputgrid.Rows[i].Cells[17].Value.ToString(), outputgrid.Rows[i].Cells[18].Value.ToString(), lstframerate.Items[i].ToString());
                            t = 0;
                        }
                    }
                    else
                    {
                        if (a >= 30)
                        {

                        }
                        else
                        {
                            lstvoidline2.Items.Add(i.ToString());
                            this.preparedgrid.Rows.Add(lstheadnumber.Items[a].ToString() + number, outputgrid.Rows[i].Cells[0].Value.ToString(), outputgrid.Rows[i].Cells[1].Value.ToString(), outputgrid.Rows[i].Cells[2].Value.ToString(), outputgrid.Rows[i].Cells[3].Value.ToString(), outputgrid.Rows[i].Cells[4].Value.ToString(), outputgrid.Rows[i].Cells[5].Value.ToString(), outputgrid.Rows[i].Cells[6].Value.ToString(), outputgrid.Rows[i].Cells[7].Value.ToString(), lstrtraw.Items[i].ToString(), lsttrans.Items[i].ToString(), lstrt1.Items[i].ToString(), outputgrid.Rows[i].Cells[11].Value.ToString(), outputgrid.Rows[i].Cells[12].Value.ToString(), outputgrid.Rows[i].Cells[13].Value.ToString(), outputgrid.Rows[i].Cells[14].Value.ToString(), outputgrid.Rows[i].Cells[15].Value.ToString(), outputgrid.Rows[i].Cells[16].Value.ToString(), outputgrid.Rows[i].Cells[17].Value.ToString(), outputgrid.Rows[i].Cells[18].Value.ToString(), lstframerate.Items[i].ToString());
                        }
                    }
                    t++;
                    number++;
                }
            }
            #endregion
        }
        void saveexcel()
        {
            object misValue = System.Reflection.Missing.Value;
            Excel.Application appExcel = new Excel.Application
            {
                Visible = false
            };
            Excel.Workbook workbook = appExcel.Workbooks.Add(misValue);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            worksheet.Name = excelfiletext.Text.ToString();

            for (int i = 0; i < preparedgrid.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = preparedgrid.Columns[i].Name.ToString();
            }
            for (int i = 0; i < preparedgrid.Rows.Count - 1; i++)
            {
                for (int j = 0; j < preparedgrid.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = preparedgrid.Rows[i].Cells[j].Value.ToString();
                }
            }


            if (onlineradiobutton.Checked == true)
            {
                string yol = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(); //yol
                string substring = yol.Substring(2, 1); //C:\Users\wasd0\Documents
                string dosyaismi = "[MERGED - Intake Session] " + nametext.Text; //dosya ismi

                string v = yol + substring + dosyaismi;
                string file = v;
                workbook.SaveCopyAs(file);
            }
            else if (localradiobutton.Checked == true)
            {
                workbook.SaveCopyAs(save.FileName);
            }

            workbook.Close(0);
            appExcel.Quit();

            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(appExcel);

            worksheet = null;
            workbook = null;
            appExcel = null;

            collectgarbage();
        }
        [Obsolete]
        void uploadexcel()
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".client_oAuth.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore("Drive.Auth.Store")).Result;
            }

            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string yol = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(); //yol
            string substring = yol.Substring(2, 1); //C:\Users\wasd0\Documents
            string dosyaismi = "[MERGED - Intake Session] " + nametext.Text; //dosya ismi

            string v = yol + substring + dosyaismi;
            string file = v;

            if (File.Exists(file))
            {
                var fileMetaData = new Google.Apis.Drive.v3.Data.File
                {
                    Name = dosyaismi
                };
                FilesResource.CreateMediaUpload request;
                using (var stream = new FileStream(file, FileMode.Open))
                {
                    request = service.Files.Create(fileMetaData, stream, "excel/xlsx");
                    request.Fields = "id";
                    request.Upload();
                    var fileId = request.ResponseBody;
                }
            }
        }
        void collectgarbage()
        {
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);
        }
        void controllingaccess()
        {
            for (int i = 0; i < sequenceaccess.Rows.Count; i++)
            {
                string y1 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                OleDbConnection baglanti1 = new OleDbConnection(y1);
                baglanti1.Open();
                string sil = "delete from sira where harf=@harf";
                OleDbCommand komut1 = new OleDbCommand(sil, baglanti1);

                komut1.Parameters.AddWithValue("@harf", lstheadnumber.Items[i].ToString()); //eh be tahir, önceden lst_order'ı silersen tabi değeri algılayamaz aq sonradna fark ettim bunu çünkü her şey doğru gözüküyor sdfgdfgf
                komut1.ExecuteNonQuery();

                baglanti1.Close();
            }

            sequencetext.Clear();
            lstheadnumber.Items.Clear();
            lblfilename.Text = "";
            pctokay_output.Visible = false;
        }


        private void main_Load(object sender, EventArgs e)
        {
            fillgrid();
            controlfillgrid();

            if (sequenceaccess.Rows.Count != 0)
            {
                for (int f = 0; f < sequenceaccess.Rows.Count; f++)
                {
                    lstheadnumber.Items.Add(sequenceaccess.Rows[f].Cells[0].Value.ToString());
                }

                sequencetext.Text = lstheadnumber.Items[0].ToString() + Environment.NewLine + lstheadnumber.Items[1].ToString() + Environment.NewLine + lstheadnumber.Items[2].ToString() + Environment.NewLine + lstheadnumber.Items[3].ToString() + Environment.NewLine + lstheadnumber.Items[4].ToString() + Environment.NewLine + lstheadnumber.Items[5].ToString() + Environment.NewLine + lstheadnumber.Items[6].ToString() + Environment.NewLine + lstheadnumber.Items[7].ToString() + Environment.NewLine + lstheadnumber.Items[8].ToString() + Environment.NewLine + lstheadnumber.Items[9].ToString() + Environment.NewLine + lstheadnumber.Items[10].ToString() + Environment.NewLine + lstheadnumber.Items[11].ToString() + Environment.NewLine + lstheadnumber.Items[12].ToString() + Environment.NewLine + lstheadnumber.Items[13].ToString() + Environment.NewLine + lstheadnumber.Items[14].ToString() + Environment.NewLine + lstheadnumber.Items[15].ToString() + Environment.NewLine + lstheadnumber.Items[16].ToString() + Environment.NewLine + lstheadnumber.Items[17].ToString() + Environment.NewLine + lstheadnumber.Items[18].ToString() + Environment.NewLine + lstheadnumber.Items[19].ToString() + Environment.NewLine + lstheadnumber.Items[20].ToString() + Environment.NewLine + lstheadnumber.Items[21].ToString() + Environment.NewLine + lstheadnumber.Items[22].ToString() + Environment.NewLine + lstheadnumber.Items[23].ToString() + Environment.NewLine + lstheadnumber.Items[24].ToString() + Environment.NewLine + lstheadnumber.Items[25].ToString() + Environment.NewLine + lstheadnumber.Items[26].ToString() + Environment.NewLine + lstheadnumber.Items[27].ToString() + Environment.NewLine + lstheadnumber.Items[28].ToString() + Environment.NewLine + lstheadnumber.Items[29].ToString();

                lbloutput.Text = "Active =>";
                btn_outputexcel.Enabled = true;
                pctwarning_output.Visible = true;
            }

            Devices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo f in Devices)
            {
                cmbcameras.Items.Add(f.Name);
            }

            cmbcameras.SelectedIndex = 0;
            camerasindextext.Text = "0";

            string y2 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
            OleDbConnection baglanti2 = new OleDbConnection(y2);
            baglanti2.Open();
            string comand = "DELETE FROM hir";
            OleDbCommand cmd = new OleDbCommand(comand, baglanti2);
            cmd.ExecuteNonQuery();
            baglanti2.Close();
        }
        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Source != null)
            {
                Source.Stop();
                collectgarbage();
            }
            collectgarbage();
            killzombieexcel(lblfilename.Text.ToString());
            Application.Exit();
        }
        private void Source_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            pctbox_webcam.Image = (Bitmap)eventArgs.Frame.Clone();
        }
        private void cmbcameras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (camerasindextext.Text == "1")
            {
                Source.Stop();
                timerofcameras.Stop();

                Source = new VideoCaptureDevice(Devices[cmbcameras.SelectedIndex].MonikerString);
                Source.NewFrame += Source_NewFrame; //oha çok iyiymiş lan bu; taba basarak alttaki metodu oluşturdu
                Source.Start();
                timerofcameras.Start();
            }
        }
        private void btn_outputexcel_Click(object sender, EventArgs e)
        {
            try
            {
                workbooktext.Text = "";
                excelfiletext.Text = "";

                collectgarbage();

                OpenFileDialog openfile1 = new OpenFileDialog
                {
                    Filter = ".xlsx|*.xlsx|.xls|*.xls|.csv|*.csv",
                    Title = "Select Excel Folder..."
                };
                if (openfile1.ShowDialog() == DialogResult.OK)
                {
                    this.workbooktext.Text = openfile1.FileName; //dosyanın tamamı

                    string y = openfile1.FileName.ToString();
                    substringtext.Text = Path.GetDirectoryName(y);

                    FileInfo uzanti = new FileInfo(openfile1.FileName);
                    extentiontext.Text = uzanti.Extension;

                    int kes = substringtext.TextLength;
                    nametext.Text = workbooktext.Text.Substring(kes + 1);

                    label.Default.text = "Loading...";
                    label.Default.Save();

                    string y2 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                    OleDbConnection baglanti2 = new OleDbConnection(y2);
                    baglanti2.Open();
                    string comand = "DELETE FROM hir";
                    OleDbCommand cmd = new OleDbCommand(comand, baglanti2);
                    cmd.ExecuteNonQuery();
                    baglanti2.Close();

                    using (loadingscreen frm = new loadingscreen(importexcel))
                    {
                        frm.ShowDialog(this);
                    }

                    pctwarning_output.Visible = false;
                    pctokay_output.Visible = true;
                    pctwarning_merged.Visible = true;
                    pctokay_merged.Visible = false;
                    btn_mergedexcel.Enabled = true;
                    lblsave.Text = "Active =>";
                    this.ActiveControl = nullfocustext;
                    localradiobutton.Enabled = true;
                    onlineradiobutton.Enabled = true;
                    lbloutput.Text = "Done! =>";

                    lblfilename.AutoSize = true;
                    lblfilename.MaximumSize = new Size(100, 103);
                    lblfilename.Text = nametext.Text;

                    lbltitlefolder.Visible = true;
                    lblunderline2.Visible = true;
                    lblfilename.Visible = true;

                    pctwarning_merged.Visible = true;
                    pctokay_merged.Visible = false;
                    lblsave.Text = "Active =>";

                    label.Default.text = "Creating...";
                    label.Default.Save();

                    //Creating ; using (hazirlik frm = new hazirlik(writeexceltogrid))
                    using (loadingscreen frm = new loadingscreen(writeexceltogrid))
                    {
                        frm.ShowDialog(this);
                    }

                    MessageBox.Show(@"""Loading...""" + " and " + @"""Creating...""" + "process was successful.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select file!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch
            {
                #region
                collectgarbage();
                killzombieexcel(lblfilename.Text.ToString());
                MessageBox.Show("A problem was encountered and solved. Please try again." + "\n\n" + "If you see this warning time and again, please restart the program and deleting all " + @"""EXCEL.EXE""" + " from your Task Manager. " + "\n\n" + "Still if you managed to get this warning again, please restart your computer!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                #endregion
            }
        }
        [Obsolete]
        private void btn_mergedexcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (onlineradiobutton.Checked == true)
                {
                    label.Default.text = "Merging and Saving...";
                    label.Default.Save();
                    //Saving ; using (saving frm = new saving(saveexcel))
                    #region
                    using (loadingscreen frm = new loadingscreen(saveexcel))
                    {
                        frm.ShowDialog(this);
                    }
                    #endregion

                    label.Default.text = "Uploading...";
                    label.Default.Save();

                    //upload ; using (upload frm2 = new upload(uploadexcel))
                    #region
                    using (loadingscreen frm2 = new loadingscreen(uploadexcel))
                    {
                        frm2.ShowDialog(this);
                    }
                    #endregion

                    //adjust
                    #region
                    pctwarning_merged.Visible = false;
                    pctokay_merged.Visible = true;
                    lblsave.Text = "Done! =>";
                    this.ActiveControl = nullfocustext;

                    string yol = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(); //yol
                    string substring = yol.Substring(2, 1); //C:\Users\wasd0\Documents
                    string dosyaismi = "[MERGED - Intake Session] " + nametext.Text; //dosya ismi

                    string v = yol + substring + dosyaismi;
                    string file = v;

                    File.Delete(file);

                    DataView dv = ds2.Tables[0].DefaultView;
                    dv.RowFilter = "isim Like '" + lblfilename.Text + "%'";
                    accessgrid.DataSource = dv;

                    if (accessgrid.Rows.Count == 0)
                    {
                        string y1 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                        OleDbConnection baglanti1 = new OleDbConnection(y1);
                        baglanti1.Open();
                        string ekle1 = "insert into hir (isim,l,u) values (@isim,@l,@u)";
                        OleDbCommand komut1 = new OleDbCommand(ekle1, baglanti1);

                        komut1.Parameters.AddWithValue("@isim", lblfilename.Text.ToString());
                        komut1.Parameters.AddWithValue("@l", "");
                        komut1.Parameters.AddWithValue("@u", "upload");
                        komut1.ExecuteNonQuery();

                        baglanti1.Close();
                        komut1.Dispose();
                    }
                    else
                    {
                        string y1 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                        OleDbConnection baglanti1 = new OleDbConnection(y1);
                        baglanti1.Open();
                        string guncelle1 = "update hir set isim=@isim, l=@l, u=@u where ID=@ID";
                        OleDbCommand cmd = new OleDbCommand(guncelle1, baglanti1);

                        cmd.Parameters.AddWithValue("@isim", lblfilename.Text.ToString());
                        cmd.Parameters.AddWithValue("@l", accessgrid.Rows[0].Cells[1].Value.ToString());
                        cmd.Parameters.AddWithValue("@u", "upload");
                        cmd.Parameters.AddWithValue("@ID", accessgrid.Rows[0].Cells[3].Value.ToString());
                        cmd.ExecuteNonQuery();

                        baglanti1.Close();
                        cmd.Dispose();
                    }

                    fillgrid();
                    controlfillgrid();

                    onlinekontroltext.Text = "upload";

                    MessageBox.Show(@"""Merging and Saving...""" + " and " + @"""Uploading...""" + "process was successful. Good luck for the research!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    collectgarbage();
                    #endregion
                }
                else if (localradiobutton.Checked == true)
                {
                    save.OverwritePrompt = false;
                    save.Title = "Save Merged-Excel File";
                    save.DefaultExt = "xlsx";
                    save.Filter = "*.xlsx|*.xlsx";

                    if (extentiontext.Text == ".csv")
                    {
                        string csv = nametext.Text.Substring(0, nametext.Text.Length - 3);
                        save.FileName = "[MERGED - Intake Session] " + csv + "xlsx";
                    }
                    else
                    {
                        save.FileName = "[MERGED - Intake Session] " + nametext.Text;
                    }

                    if (save.ShowDialog() == DialogResult.OK)
                    {
                        label.Default.text = "Merging and Saving...";
                        label.Default.Save();

                        //merge and save ; using (merge frm = new merge(saveexcel))
                        #region
                        using (loadingscreen frm = new loadingscreen(saveexcel))
                        {
                            frm.ShowDialog(this);
                        }
                        #endregion

                        //adjust
                        #region
                        pctwarning_merged.Visible = false;
                        pctokay_merged.Visible = true;
                        lblsave.Text = "Done! =>";
                        this.ActiveControl = nullfocustext;

                        DataView dv = ds2.Tables[0].DefaultView;
                        dv.RowFilter = "isim Like '" + lblfilename.Text + "%'";
                        accessgrid.DataSource = dv;

                        if (accessgrid.Rows.Count == 0)
                        {
                            string y1 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                            OleDbConnection baglanti1 = new OleDbConnection(y1);
                            baglanti1.Open();
                            string ekle1 = "insert into hir (isim,l,u) values (@isim,@l,@u)";
                            OleDbCommand komut1 = new OleDbCommand(ekle1, baglanti1);

                            komut1.Parameters.AddWithValue("@isim", lblfilename.Text.ToString());
                            komut1.Parameters.AddWithValue("@l", "local");
                            komut1.Parameters.AddWithValue("@u", "");
                            komut1.ExecuteNonQuery();

                            baglanti1.Close();
                            komut1.Dispose();
                        }
                        else
                        {
                            string y1 = "Provider=Microsoft.ACE.Oledb.12.0;Data Source=ta.accdb";
                            OleDbConnection baglanti1 = new OleDbConnection(y1);
                            baglanti1.Open();
                            string guncelle1 = "update hir set isim=@isim, l=@l, u=@u where ID=@ID";
                            OleDbCommand cmd = new OleDbCommand(guncelle1, baglanti1);

                            cmd.Parameters.AddWithValue("@isim", lblfilename.Text.ToString());
                            cmd.Parameters.AddWithValue("@l", "local");
                            cmd.Parameters.AddWithValue("@u", accessgrid.Rows[0].Cells[2].Value.ToString());
                            cmd.Parameters.AddWithValue("@ID", accessgrid.Rows[0].Cells[3].Value.ToString());
                            cmd.ExecuteNonQuery();

                            baglanti1.Close();
                            cmd.Dispose();
                        }

                        fillgrid();
                        controlfillgrid();

                        localkontroltext.Text = "local";

                        MessageBox.Show(@"""Merging and Saving..."" process was successful. Good luck for the research!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        collectgarbage();
                        #endregion
                    }
                    else
                    {
                        MessageBox.Show("Please save file!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch
            {
                //ilginç bir hata, aç kapa iş görür
                #region
                collectgarbage();
                killzombieexcel(lblfilename.Text.ToString());
                MessageBox.Show("A problem was encountered and solved. Please try again." + "\n\n" + "If you see this warning time and again, please restart the program and deleting all " + @"""EXCEL.EXE""" + " from your Task Manager. " + "\n\n" + "Still if you managed to get this warning again, please restart your computer!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                #endregion
            }
        }
        private void timerofcameras_Tick(object sender, EventArgs e)
        {
            if (pctbox_webcam.Image != null) //eğer picturebox'a resim düşüyorsa, yani kamera çalışıyorsa
            {
                BarcodeReader reader = new BarcodeReader(); //barkod okuma açılsın
                Result result = reader.Decode((Bitmap)pctbox_webcam.Image); //barkod alınmaya çalışılıyor
                if (result != null) //eğer alınabildiyse
                {
                    try
                    {
                        timerofcameras.Stop(); //tekrar tekrar barkoda tetiklenme gereği yok durdur timer'o
                        Source.Stop(); //kamera ile bağlantını da kes işini gördün

                        string str = result.ToString(); //bu textbox'ı bir değişeken atalım rahat işlem için
                        Regex filter = new Regex(@"([A-Z]+)"); //bize o textboxtaki harfleri sırasıyla verecek kütüphane/paket
                        foreach (var item in str) //textbox içindeki her bir string ifadede
                        {
                            var match = filter.Match(item.ToString()); //filtremize uygun olacak şekide harf ayıklama
                            if (match.Success) //eğer gerçekten harf ise
                            {
                                lstheadnumber.Items.Add(match.Value); //ekle listeye
                            }
                        }

                        if (lstheadnumber.Items.Count == 30)
                        {
                            #region
                            cmbcameras.Enabled = false;
                            sequencetext.Text = lstheadnumber.Items[0].ToString() + Environment.NewLine + lstheadnumber.Items[1].ToString() + Environment.NewLine + lstheadnumber.Items[2].ToString() + Environment.NewLine + lstheadnumber.Items[3].ToString() + Environment.NewLine + lstheadnumber.Items[4].ToString() + Environment.NewLine + lstheadnumber.Items[5].ToString() + Environment.NewLine + lstheadnumber.Items[6].ToString() + Environment.NewLine + lstheadnumber.Items[7].ToString() + Environment.NewLine + lstheadnumber.Items[8].ToString() + Environment.NewLine + lstheadnumber.Items[9].ToString() + Environment.NewLine + lstheadnumber.Items[10].ToString() + Environment.NewLine + lstheadnumber.Items[11].ToString() + Environment.NewLine + lstheadnumber.Items[12].ToString() + Environment.NewLine + lstheadnumber.Items[13].ToString() + Environment.NewLine + lstheadnumber.Items[14].ToString() + Environment.NewLine + lstheadnumber.Items[15].ToString() + Environment.NewLine + lstheadnumber.Items[16].ToString() + Environment.NewLine + lstheadnumber.Items[17].ToString() + Environment.NewLine + lstheadnumber.Items[18].ToString() + Environment.NewLine + lstheadnumber.Items[19].ToString() + Environment.NewLine + lstheadnumber.Items[20].ToString() + Environment.NewLine + lstheadnumber.Items[21].ToString() + Environment.NewLine + lstheadnumber.Items[22].ToString() + Environment.NewLine + lstheadnumber.Items[23].ToString() + Environment.NewLine + lstheadnumber.Items[24].ToString() + Environment.NewLine + lstheadnumber.Items[25].ToString() + Environment.NewLine + lstheadnumber.Items[26].ToString() + Environment.NewLine + lstheadnumber.Items[27].ToString() + Environment.NewLine + lstheadnumber.Items[28].ToString() + Environment.NewLine + lstheadnumber.Items[29].ToString();

                            pctwarning_output.Visible = true;
                            btn_outputexcel.Enabled = true;
                            lbloutput.Text = "Active =>";
                            this.ActiveControl = nullfocustext;
                            sequencetext.Enabled = false;

                            MessageBox.Show("Sequence was taken successfully.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            #endregion
                        }
                        else
                        {
                            #region
                            collectgarbage();
                            MessageBox.Show(lstheadnumber.Items.Count.ToString());
                            MessageBox.Show("This order is not valid! Order's item count isn't 30! Restarting the program... Please try again.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Restart();
                            #endregion
                        }
                    }
                    catch
                    {
                        //qr kod hatalı demek, istediğimiz biçimde değil
                        #region
                        collectgarbage();
                        MessageBox.Show("This QR Code is not valid! Restarting the program... Please try again.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Restart();
                        #endregion
                    }
                }
            }
        }
        private void sequencetext_Enter(object sender, EventArgs e)
        {
            ActiveControl = nullfocustext; //kutuya tıklanmaya çalışırsa diye engelleme amaçlı
        }
        private void btnstarttocatch_Click(object sender, EventArgs e)
        {
            label.Default.text = "Loading Camera...";
            label.Default.Save();

            //loading camera ; using (camera frm = new camera(harfkontrolü))
            if (sequenceaccess.Rows.Count != 0)
            {
                using (loadingscreen frm = new loadingscreen(controllingaccess))
                {
                    frm.ShowDialog(this);
                }
            }

            Source = new VideoCaptureDevice(Devices[cmbcameras.SelectedIndex].MonikerString);
            Source.NewFrame += Source_NewFrame; //oha çok iyiymiş lan bu; taba basarak alttaki metodu oluşturdu
            Source.Start();
            timerofcameras.Start();

            int x = Convert.ToInt32(camerasindextext.Text);
            x = 1;
            camerasindextext.Text = x.ToString();
        }
        private void onlineradiobutton_Click(object sender, EventArgs e)
        {
            DataView dv = ds2.Tables[0].DefaultView;
            dv.RowFilter = "isim Like '" + lblfilename.Text + "%'";
            accessgrid.DataSource = dv;

            if (accessgrid.Rows.Count != 0)
            {
                if (accessgrid.Rows[0].Cells[1].Value.ToString() == "local" && accessgrid.Rows[0].Cells[2].Value.ToString() == "upload")
                {
                    pctokay_merged.Visible = true;
                    pctwarning_merged.Visible = false;

                    lblsave.Text = "Done! =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "local" && accessgrid.Rows[0].Cells[2].Value.ToString() == "")
                {
                    pctokay_merged.Visible = false;
                    pctwarning_merged.Visible = true;

                    lblsave.Text = "Active =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "" && accessgrid.Rows[0].Cells[2].Value.ToString() == "upload")
                {
                    pctokay_merged.Visible = true;
                    pctwarning_merged.Visible = false;

                    lblsave.Text = "Done! =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "" && accessgrid.Rows[0].Cells[2].Value.ToString() == "")
                {

                }
            }
        }
        private void localradiobutton_Click(object sender, EventArgs e)
        {
            DataView dv = ds2.Tables[0].DefaultView;
            dv.RowFilter = "isim Like '" + lblfilename.Text + "%'";
            accessgrid.DataSource = dv;

            if (accessgrid.Rows.Count != 0)
            {
                if (accessgrid.Rows[0].Cells[1].Value.ToString() == "local" && accessgrid.Rows[0].Cells[2].Value.ToString() == "upload")
                {
                    pctokay_merged.Visible = true;
                    pctwarning_merged.Visible = false;

                    lblsave.Text = "Done! =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "local" && accessgrid.Rows[0].Cells[2].Value.ToString() == "")
                {
                    pctokay_merged.Visible = true;
                    pctwarning_merged.Visible = false;

                    lblsave.Text = "Done! =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "" && accessgrid.Rows[0].Cells[2].Value.ToString() == "upload")
                {
                    pctokay_merged.Visible = false;
                    pctwarning_merged.Visible = true;

                    lblsave.Text = "Active =>";
                }
                else if (accessgrid.Rows[0].Cells[1].Value.ToString() == "" && accessgrid.Rows[0].Cells[2].Value.ToString() == "")
                {

                }
            }
        }
    }
}
