namespace QR_Automatization
{
    class notes
    {
        //private void yoket
        #region


        #region
        /*foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
                break;
            }
        }*/
        #endregion

        //IsFileinUse(label1.Text.ToString());

        #region
        /*object fileName = "NewWorkbook.xlsx";
        Excel.Workbook workbook = Application.Workbooks.get_Item(fileName);
        workbook.Close(false);*/
        #endregion

        #region
        /*var processes = from p in Process.GetProcessesByName("EXCEL")
                        select p;

        foreach (var process2 in processes)
        {
            if (process2.MainWindowTitle == "Microsoft Excel -" + ex)
                process2.Kill();
        }*/
        #endregion

        //Globals.ThisWorkbook.Close(false);
        //this.ActiveWorkbook.Close(false, missing, missing);

        #region
        /*System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
        foreach (System.Diagnostics.Process p in process)
        {
            if (!string.IsNullOrEmpty(p.ProcessName))
            {
                try
                {
                    p.Kill();
                }
                catch
                {

                }
            }
        }*/
        #endregion

        #endregion

        //ienumerable
        #region
        /*
        public IEnumerable<string> ReadLines(Func<Stream> streamProvider, Encoding encoding)
        {
            using (var stream = streamProvider())
            using (var reader = new StreamReader(stream, encoding))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    yield return line;
                }
            }
        }
        */
        #endregion

        //String alınmasının sayılı versiyonu, ek kod
        #region
        //string str = metroTextBox1.Text;
        /*string str = "B 1P 7Y 4L 9M 2C 10A 3P 7F 5L 9B 1T 8Y 4R 6M 2P 7A 3L 9Y 4C 10F 5R 6B 1C 10M 2T 8A 3R 6F 5T 8";

        string[] sayilar = Regex.Split(str, @"\D+");
        foreach (string s in sayilar)
        {
            if (int.TryParse(s, out _))
            {
                lst_sayi.Items.Add(s);
            }
        }*/
        #endregion



        //frame logaritmik hata düzeltme
        //nedense buradan çeviremedim değeri bigint'e, başka yöntemlerle bulduk artık çözümü
        #region
        //MessageBox.Show(excelReader.GetValue(19).ToString()); //ABİ DEĞER DİREKT E'Lİ GELİYOR

        /*BigInteger j;
        BigInteger.TryParse(Convert.ToString(excelReader.GetValue(19)), out j);
        var bigintegervalue = j.ToString();
        MessageBox.Show(BigInteger.Parse(bigintegervalue).ToString());*/


        //abi sıfır atıyor ve false dönüyor, NİYE AQ

        //BigInteger.TryParse(Convert.ToString(excelReader.GetValue(19)), out j);

        //BigInteger bigIntFromDouble = new BigInteger(Convert.ToString(excelReader.GetValue(19)));
        //MessageBox.Show(bigIntFromDouble.ToString());

        /*if (BigInteger.TryParse(Convert.ToString(excelReader.GetValue(19)), out _) == true)
        {
            MessageBox.Show("true");
        }
        else
        {
            MessageBox.Show("false");
        }*/

        //BigInteger j = BigInteger.Parse(Convert.ToString(excelReader.GetValue(19)));
        //MessageBox.Show(j.ToString());
        #endregion

        //excel data reader övgümü içerir
        #region
        //vay aq getstring diyince dönüştürme hatası verdi
        //bu excel reader olayını sevdim güzel oldu bunu öğrendiğim
        //en gelişmiş excel okuma yöntemlerinden birisi bu olmalı çünkü bu direkt olarak hücreyi okuyor, yüzeysel olarak okumuyor en güzel tarafı o zaten bu sayede uyuşmazlık durumu olmuyor
        //ama her zaman hücreyi direkt okumak en iyisi demek değil çünkü bazen o rating transta olduğu gibi ne hücrenin içi ne de dışı yarıyor mecburen çevirme yapıyorsun
        //bak gördün mü mesela benim ilk kullandığım 29.04'ü sayısal tuhaf bir değere çeviriyordu bu da tam tahmin ettiğim gibi katı bir şekilde hücrede ne yazıyorsa onu okuyarak tarih şeklinde bana değeri verdi 
        //oysaki bana bu ikisi de lazım değildi bunun için de dönüşüm uygulamam şarttı ya da bana sağladığı bir adım kolaylık ile direkt ilk 4'ü alıp önüme bakabilirdim ama güzel bir fonksiyon da öğrenmiş oldum onun sayesinde
        //ama çooook kral br okuma şekli her halinden belli
        //mesela logaritmik yerleri o da logaritmik almak durumunda kaldı çünkü int'in sınırları belli, onu aşınca program ne yapsın exceldatareader mecburen e'li okudu değerleri
        #endregion

        //eski ve bu algoritma için geçersiz olan ama kendi içinde güzel mantık kurduğumu düşündüğüm yapı, ama faydalarını çok gördüm
        #region
        /*string[] sayilar = Regex.Split(e, @"\D+");
        foreach (string s in sayilar)
        {
            if (int.TryParse(s, out _))
            {
                int modlanacak = Convert.ToInt32(s);
                modlanacak++;
                lst_e.Items.Add(modlanacak);
            }
        }

        int sayi = Convert.ToInt32(lst_e.Items[c]);
        int mod = sayi % 3;

        if (mod == 2)
        {
            string ana_ikisifir = ana + "00";
            lst_framerate.Items[m] = ana_ikisifir.ToString();
        }
        else if (mod == 1)
        {
            string ana_birsifir = ana + "0";
            lst_framerate.Items[m] = ana_birsifir.ToString();
        }
        else
        {

        }
        c++;*/

        /*string ana_dondur = txt_logaritmik.Text.Substring(0, 13);
        string ana = ana_dondur.Replace(",", ""); //tamam sayı bu oldu şu an
        string sifir_dondur = txt_logaritmik.Text.Substring(14, 3);
        string sifir = sifir_dondur.Replace(sifir_dondur, "000");
        string framerate = ana + sifir;*/
        #endregion



        //rating tran hata düzeltme
        //Alternatif versiyon 1
        #region
        /*int gelanam = Convert.ToInt32(metrogrid_alinan.Rows[g].Cells[9].Value.ToString()); //lan boşluklara geldiği için ondan okuyamıyormuş haliyle!!!
          DateTime dt = DateTime.FromOADate(gelanam);
          string alalimseni = dt.ToString();
          string tarih = alalimseni.Substring(0, alalimseni.Length - 9); //1.01.2021 gibi oldu şu an
          string degerimiz = tarih.Substring(0,tarih.Length - 5); //1.01 gibi oldu şu an
          metrogrid_alinan.Rows[g].Cells[9].Value = degerimiz.ToString();*/
        #endregion

        //Alternatif versiyon 2
        #region
        /*string kontrol_saat = rating_trans.Substring(rating_trans.Length - 9); //00.00.00 gibi oldu şu an
        //MAL TAHİR!!! DEĞER TABİ Kİ SIFIRDAN KÜÇÜK OLAMAZ VE PROGRAM HATA VERİR SAYIN AMINA KODUĞUMUN TAHİRİ DİKKAT ETSENE ŞUNA, EKLE ŞU ISNULLOREMPTY'İ
        //yav ekledinde kontrol saatte hata fırlatacaktır program sen kontrol etmeden evvel!!! Keşke bunu da düşünseydin!!!
        if (string.IsNullOrEmpty(kontrol_saat) == false)
        {

        }
        else
        {

        }*/
        #endregion

        //private void yoket()
        #region
        //private void yoket()
        //{
        #region
        /*foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
                break;
            }
        }*/
        #endregion

        //IsFileinUse(label1.Text.ToString());

        #region
        /*object fileName = "NewWorkbook.xlsx";
        Excel.Workbook workbook = Application.Workbooks.get_Item(fileName);
        workbook.Close(false);*/
        #endregion

        #region
        /*var processes = from p in Process.GetProcessesByName("EXCEL")
                        select p;

        foreach (var process2 in processes)
        {
            if (process2.MainWindowTitle == "Microsoft Excel -" + ex)
                process2.Kill();
        }*/
        #endregion

        //Globals.ThisWorkbook.Close(false);
        //this.ActiveWorkbook.Close(false, missing, missing);


        #region
        /*System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
        foreach (System.Diagnostics.Process p in process)
        {
            if (!string.IsNullOrEmpty(p.ProcessName))
            {
                try
                {
                    p.Kill();
                }
                catch
                {

                }
            }
        }*/
        #endregion
        //}
        #endregion

        //tüm zombie exceller öldü daha hala geride kalan varsa
        //MessageBox.Show("geldik");
        //string y = Process.GetProcessesByName("EXCEL")[0].StartInfo.FileName.ToString();
        //MessageBox.Show(y.ToString());

        //MessageBox.Show(x.ToString());
        //hmm, hangisi benim excelim ki o zaman, onun id'sini almak lazım
        //BULDUM LAN, SEN KİMİN ELİNDEN KAÇIYORSUN AQ YA, 
        //EĞER BEN BU EXCEL AÇILMADAN EVVEL EĞER AÇIK OLAN EXCEL VARSA ONUN KODUNU ALIRSAM SONRA FARKLI OLANIN
        //BENİM UYGULAMANIN EXCELİ OLDUĞU ANLAŞILIR VE BU SAYEDE GÖNÜL RAHATLIĞIYLA İLGİLİ EXCELİ SONLANDIRABİLİRİM
        //VE HERHANGİ BİR SORUN DA ORTADA KALMAZ!!!

        //evet başarıyla öldürdü, peki neden kapatmadı, çünkü bu zombie excelleri öldürüyor, ana exceli öldürmüyor
        //bunun için öncekiler lazım bana, zaten dikkatini çekerse tak diye kapandı bu sefer diğerlerinden farklı olarak


        //bir defa bana backgroundworker olmaz çünkü benim metot tek dönmelik tek seferlik bir metot
        //bana bu kod bloğunun süresi lazım, sonra işte timer ile sn. sn. ilerteip backgroundworker'ı öyle kullanabilirim
        //yok be gerek yokmuş, robot sesli ablamız sağolsun bize ışık tuttu



        //string e = virgül.Split('E').Last(); //E+..'lı kısımlar alındı //hayır alınmamış, E'den sonra gelenler alınmış, o yüzden de pek bi işimi görmedi bu güzel kod, yeri geldi mi çok güzel kullanılır bu metot
        //programcı ne kadar çok metot kullandığı ile değil kullandığı metotları ne kadar spesifikleştirebildiği ile anlaşılır metot kullanmakta sorun yok, sorun, kendin özgün olacak şekilde onları kullanma ve yazma


        //lan oğlum neg. logaritmalarda bu hata meydana geliyor çünkü bir karakter daha fazla oluyor ya ondan kaynaklı bir sıkıntı oluyormuş meğerse

        //burada 44000 bilmem ne gibi gelmeyecek tarih direkt hücreyi okuduğumuz için
        //ama aynı zamanda diğer değerler de noktalı geleceği için
        //tarihli değerlerin nokta ve iki nokta üst üsteden barındırmasından yararlanılabilir


        //bi de bunun listboxlu olanını yapalım o zaman abi, çünkü ona düzgün aktarılıyor ama o düzgün aktaramıyorsa sorun onda olabilir


        //workbook.Save(); //istediğim gibi kaydetmiyor çünkü dosyanın ismini kendisi belirliyor, böyle bir kaydetme benim işimi görmez
        //savefile name'de direkt kaydedilecek adres veriliyordu, bana da o lazım ama save dialog kullanmıyorum otomatik kaydedecek
        //bu yüzden bana manual olacak şekilde bir savefilename lazım

        //not listed olacak pct warning açılacak ve bu silme işlemi kamera açılmasından sonra olsun hatta eş zamanlı olsun
        //eş zamanlı olamıyor maalesef
        //işe yaramayan ama beğendiğim multihread, çünkü kamera özelliği böyle çalışmıyormuş
        #region
        /*Thread th1 = new Thread(new ThreadStart(kamera));
        Thread th2 = new Thread(new ThreadStart(harfsil));

        th1.Priority = ThreadPriority.Highest;
        th2.Priority = ThreadPriority.Lowest;

        th1.Start();
        th1.Join();
        th2.Start();*/
        #endregion


        //güzelmiş konulmasa da olabilir belki ek kamera gelirse, e yanda kocaman ok var e bi zahmet görsün onu,
        //bu sayede kamerayı sistemin direkt algıladığını anlasın hoca

        //sen kimsin de benim elimden kaçıyorsun ya, bunun uzunluğu kadar başından silerim geriye bana dosya name'i kalmış olur saf!

        #region
        //eğer varsa güncelleme yapmak lazım o zaman abi, doğru ya, ben bu büyük ayrıntıyı atlamışım aq sdafdgfd
        /*komut1.Parameters.AddWithValue("@isim", label1.Text.ToString());
        komut1.Parameters.AddWithValue("@l", mt_save.Rows[0].Cells[1].Value.ToString());
        komut1.Parameters.AddWithValue("@u", "upload");
        komut1.ExecuteNonQuery();*/
        #endregion

        #region
        //eğer varsa güncelleme yapmak lazım o zaman abi, doğru ya, ben bu büyük ayrıntıyı atlamışım aq sdafdgfd
        /*komut1.Parameters.AddWithValue("@isim", label1.Text.ToString());
        komut1.Parameters.AddWithValue("@l", "local");
        komut1.Parameters.AddWithValue("@u", mt_save.Rows[0].Cells[2].Value.ToString());
        komut1.ExecuteNonQuery();*/
        #endregion
    }
}
