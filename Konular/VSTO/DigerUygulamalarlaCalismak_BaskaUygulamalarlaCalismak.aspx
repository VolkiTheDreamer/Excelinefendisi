<%@ Page Title='Başka Uygulamalarla Çalışmak' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>


<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'></div>
    <h1>Diğer Uygulamalarla Çalışmak</h1>
    <p>Bu bölümde hem diğer uygulamalar(Hesap makinesi, Spotify v.s) nasıl çalıştırılır ona bakacğaız, hem de daha önemlisi, diğer programlama dilleriyle yazılmış kodları nasıl çalıştırırız ona bakacağız.</p>
    <h2 class='baslik'>Giriş</h2>
    <div class='konu'>
        <p>İlk olarak basit bir uygulama nasıl açılıyor ona bakalım. VBA&#39;de <strong>Shell</strong> komutu ile yaptığımız bu işi .Net&#39;te <strong>Process</strong> sınıfı ile yapıyoruz. Bunu yapmanın da birkaç yolu var. Hepsine bakalım.</p>
        <p>Bu arada bu örneği, ayrı bir proje yapmak yerine aşağıdaki InvestPY projesi içine koymayı uygun gördüm. </p>
        <h3>İlk yöntem</h3>
        <p>Notepad ile bir metin dosyasını açacağız. <strong>Process</strong> sınıfının <strong>Start</strong> metodu ile programın adını ve parametreleri string olarak veriyoruz. En basit hali budur.</p>
        <pre class="brush:csharp">
using System.Diagnostics; //bahsekonu sınıf bu namespace içinde

Process.Start(&quot;calc&quot;); //exe uzantısına gerek yok
Process.Start(&quot;notepad.exe&quot;, @&quot;E:\OneDrive\Dökümanlar\GitHub\dotnet\Ugulamalar\InvestPY\myinvestpy.py&quot;);  </pre>
        <h3>İkinci Yöntem</h3>
        <p>Process sınıfından bir nesne yaratıp, parametreleri de aşağıdaki gibi belirliyoruz. Bu yöntemle Process sınıfının çok daha zengin üye listesine erişebiliyoruz.</p>
        <pre class="brush:csharp">
Process notePad = new Process();
notePad.StartInfo.FileName = "jupyter";
notePad.StartInfo.Arguments = "notebook";
notePad.Start();  </pre>
        <h3>Üçüncü Yöntem</h3>
        <p><strong>ProcessStartInfo</strong> sınıfını da devreye sokuyoruz. Önce StartInfo bilgilerini oluşturuyoruz, sonra usign block&#39;u içinde Process sınıfını devreye alıyoruz. İkinci&nbsp; yöntemle arasındaki fark için <a href="https://stackoverflow.com/questions/2890310/what-s-the-difference-between-process-and-processstartinfo-in-c">buraya</a> bakınız. Ancak çoğunlukla ilk yöntem bile yeterli olacaktır.</p>
        <pre class="brush:csharp">
ProcessStartInfo start = new ProcessStartInfo();
start.Arguments = "www.excelinefendisi.com";
start.FileName = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge";           
int exitCode;

using (Process proc = Process.Start(start))
{
    proc.WaitForExit();
    exitCode = proc.ExitCode;
    MessageBox.Show(exitCode.ToString()); //sorunsuz ise 0 çıkar
}       </pre>
    </div>

    <h2 class='baslik'>Başka programlama dili dosyalarıyla çalışıtırmak</h2>
    <div class='konu'>
        <p>Bu sefer, işleri biraz daha ileri götürüp, başka bir programlama diline yazılmış bir scripti çalıştırıp Excel&#39;le nasıl bağlantı kurarız onu göreceğiz. Aslında yukarıdaki mantıktan pek de bir farkı yok. FileName olarak ilgili programlama dilinin derleyicisini, argument olarak da gerekli tüm argümanları verceğiz. Bu argümanlar içnde çalışıtırılacak script dosyası ve bu dosyaya dışardan verilen parametreler de dahil olacak.</p>
        <p>Bu ilk örneğimizde Excel&#39;deki tüm işi <strong>Python&#39;a</strong> yaptıran bir kodumuz olacak. Bunun için PC&#39;mizde python kurulu olması gerekiyor. Bu vesileyle Python öğrenmediyseniz bunu da mulatka öğrenmenizi tavisye ederim. Hem öğrenmesi çok kolay bir dil, hem de çok güzel işler yapılabiliyor. Özellike Veri Bilimi, Makine Öğrenimi ve Yapay Zeka konularında önde gelen dildir. <a href="ThirdPartyKutuphanler_ExcelDNA.aspx">Excel DNA</a> sayfasında gördüğümüz gibi pyxll isimli kütüphane ile de Excel için XLL add-in&#39;ler yazılabilir.</p>
        <p>Bu projede, <a href="https://www.investing.com/">https://www.investing.com/</a> adresinden çeşitli ekonomik verileri çeken bir python kodumuzu var. Python&#39;da webden veri çekme ile ilgili olarak BeautifulSoup diye bir kütüphane var ancak şanslıyız ki birileri bu siteden veri çekecek bir kütüphane(API) yazmış bile, o yüzden python&#39;da çok basit bir kod yazdım. Ancak diğer web siteleri için böyle hazır api olmayabilir, o yüzden BeautifulSoup kullanmak gerekirdi. Kodun çalışması için PC&#39;mizde <strong>pandas, investpy ve openpyxl</strong> adlı python kütüphanelerinin de kurulu olması gerekmektedir.</p>
        <p>Kod, özetle bu siteden Türkiye&#39;ye ait çeşitli tipteki yatırım araçları için ilk 10 kıymete ait biglileri getiriyor, ve bunları Excel&#39;e yazdırıyıor. Bu kısmı tamamen Python yapıyor. Bizim Excel&#39;de sunduğumuz fonksiyonalite ise kullanıcıya çeşitli bilgileri Ribbon&#39;dan girdirmek, bir nevi programı kullanmak için kullanıcıya arayüz sağlamak. Kullanıcıların hiç python bilmememsi, tüm python dosyalarını sizin hazırladığınız bir durumda kullanıcılara python kurdurtmak, sonrasında kütüphane kurdurtmak da ayrı bir sorun olabilir. O yüzden buraya isterseniz, kullanıcıya pythonı <a href="https://stackoverflow.com/questions/35684243/how-to-download-and-run-a-exe-file-c-sharp">indirip kurulumu yaptıran</a>, kütüphaneleri indirtmeyi&nbsp;ve python path&#39;ini öğrenmeyi sağlayan butonlar da koyabilirsiniz. Bunu bir ara ödev olarak hazırlayıp aşağı Ödevler bölümüne koyacağım.</p>
        <p>Çalıştırdıktan sonra şöyle Ribbonumuz şöyle görünür:</p>
        <p><img alt="" src="../../images/VSTO_invest2.jpg" style="width: 408px; height: 107px" /></p>
        <p>Farkettiyseniz Ribbondaki text kutuları oldukça kısa olup girdiğimiz tüm metin burda görünmüyor. Bunun yerine Invest group&#39;unun sağ altındaki dialog launcher ile açılan Settings formuna da bu bilgiler girilebilir. Bunu ödev olarak düşünebilirsiniz. Ben formu proje dosyası içine dahil ettim, siz sadece gerekli ayarlamaları yapın.</p>
        <p>Çalışmayı, python kodu da dahil olacak şekilde <a href="https://github.com/VolkiTheDreamer/excel/tree/main/InvestPY">buradan</a> indirebilirsiniz.</p>
        <p>Şimdi kodlara bakıp sonra açıklamasına geçelim. </p>
        <pre class="brush:csharp">
public partial class Ribbon1
    {
        Process process;
        bool iptal = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //ilk değer atamaları
            this.editBox1.Text = DateTime.Today.AddDays(-10).ToString("dd/MM/yyyy").Replace(".", "/");
            this.editBox2.Text = DateTime.Today.ToString("dd/MM/yyyy").Replace(".","/");
            this.editBox3.Text = @"C:\Invest\sonuclar.xlsx"; //bu ve alttaki settings formuna koyarak da yapılabilir, siz böyle deneyin
            this.editBox4.Text = @"C:\Users\volka\AppData\Local\Programs\Python\Python38\python.exe";
        }

        private async void button1_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.StatusBar = "Lüfen bekleyiniz...";//daha görünür olması için A1&#39;e de yazdırabiliriz
            iptal = false;
            string mesaj = await fetchdata();

            Globals.ThisAddIn.Application.StatusBar = "";

            if (iptal)
            {
                MessageBox.Show("İptal edildi");
            }
            else
            {
                Globals.ThisAddIn.Application.Workbooks.Open(this.editBox3.Text);

                if (mesaj.Length < 100)
                    MessageBox.Show(mesaj);
                else
                {
                    Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                    ws.Name = "Sonuç-Hata mesajı";
                    ws.Cells[1, 1] = mesaj;
                }
            }
        }

        private async Task<string> fetchdata()
        {
            
            string sonuc = await Task.Run(() =>
            {
                string pyfile = @"E:\OneDrive\Dökümanlar\GitHub\dotnet\Ugulamalar\InvestPY\myinvestpy.py";

                ProcessStartInfo start = new ProcessStartInfo();
                start.FileName = this.editBox4.Text;
                start.Arguments = pyfile + " " + this.editBox1.Text + " " + this.editBox2.Text + " " + this.editBox3.Text;
                start.UseShellExecute = false;
                start.CreateNoWindow = true; 
                start.RedirectStandardOutput = true;
                start.RedirectStandardError = true;
                using (process = Process.Start(start))
                {
                    using (StreamReader reader = process.StandardOutput)
                    {
                        string stderr = process.StandardError.ReadToEnd(); 
                        string result = reader.ReadToEnd(); 

                        if (string.IsNullOrEmpty(stderr))
                            return string.Format("Sonuç:{0}", result);
                        else
                            return string.Format("Hata:{0}", stderr);
                    }
                }
            });
            return sonuc;

        }
        
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            process.Kill();
            iptal = true;
            Globals.ThisAddIn.Application.StatusBar = "";            
        }
    }
        </pre>
        <p>Öncelikle şunu belirtmek isterim ki, burdaki <strong>async/await</strong> kullanımı bu konuyu anlamak için de güzel bir fırsat oldu. Siz tabiki bu ifadelerin teorisini iyice araştırmalısınız. Burda yapmaya çalıştığım, data gelene kadar Excel kitlenmesin ve kullanılabilir durumda olsun. Keza, bir şekilde işlem çok uzun sürecek gibi olursa da işlemi İptal butonuyla iptal edebilelim. Keza, o sırada diğer Process group&#39;undaki butonları da çalıştırabilirsiniz.</p>
        <p>Programın esasına gelecek olursak, Excel&#39;in sunduğu ribbon arayüz(veya settings form) aracılığı ile kullanıcıyı Python Console&#39;da çalışmaktan kurtartmış ve aşina olduğu Excel&#39;de kalmasını sağlıyoruz. Parametreleri girip Getir butonuna bastıktan sonra Python kodu çalışıyor. Biz burada ilk başlıkta gördüğümüz 3. yöntemi kullanmış olduk. Yukarıdaki örneklerden farklı olarak gördüğümüz satırlar ve anlamları şöyle:</p>
        <ul>
            <li><strong>UseShellExecute</strong> özelliğine false atadık. Normalde default değer true&#39;dur ve bu bizim VBA&#39;de de bildiğmiz Shell komutunu çalıştırmaya denk veya daha genele ifadeyle cmd komut satırından ilgili komutun verilmesine denk. Ancak biz&nbsp;geriye bir değer(python&#39;daki print komutlarının sonucu) döndüreceğimiz için Shell kullanmayacağız, onun yerine <strong>CreateProcess</strong> fonksiyonunun çalışmasını istiyoruz. Bu ikisi arasındaki farka ait detay bilgiye&nbsp;<a href="https://stackoverflow.com/questions/5255086/when-do-we-need-to-set-processstartinfo-useshellexecute-to-true">buradan</a> ulaşabilirsiniz.</li>
            <li><strong>CreateNoWindow </strong>= true diyerek, bir pencere açılmasını istemiyoruz.</li>
            <li><strong>RedirectStandardOutput ve RedirectStandardError </strong>değerlerine true atayarak hem çalışan programın döndüreceği çıktıyı(Python&#39;da print ifadeleri) hem de olası hata mesajlarını kullanacağız.</li>
            <li>using bloklarından içteki blokta ana mesajla hata mesajını bi değişkene alıyoruz. Bu arada python kodunda try-catch blokları var, ordaki hata değerlerini de ana çıktının bir parçası olarak alıyorum, zira ordaki try-catch blokları for döngüleri içinde, böylece bir hisse senedinin kaydı bulunmaması gibi durumlarda oluşan hatalar nedeniyle program tamamaen durmuyor, print mesajı ile çıktılar console&#39;a yazdırılıyor. İşte biz stadard output ile bu çıktıları yakalıyoruz. Standard error ise program hata verip durduğunda çıkan hata mesajını yakalıyor.</li>
        </ul>
        <p>
            Benzer bir çalışmayı siz bildiğiniz başka diller için de deneyebilirsiniz.</p>
        <p>Bu arada bu çalışmayı komple c#&#39;ta yapmak isteseydik bunun için de <a href="https://html-agility-pack.net/">HtmlAgility</a> isimli efsane bir paket var, nuget&#39;tan bunu da indirebilirsiniz. Tabi bu aslında Python&#39;daki BeautifulSoup&#39;un muadili oluyor, investpy&#39;nin değil. Zaten bu yüzden c# yerine python kullandık. Çünkü hazır bir API olduğu için python&#39;la ilerlemek çok daha makul oldu.</p>
    </div>

    <h2 class='baslik'>İkinci Python örneği</h2>
    <div class='konu'>
        <p>Bu ikinci örneğimizde ise, diğer programa sadece metin parametre göndermekle kalmayacağız, aynı zamanda structered bir data da göndereceğiz. İşte burada karşımıza meşhur <strong>json</strong> yapıları ve <strong>serialization/deserialization</strong> kavramları çıkıyor. </p>
        <p>Örneğimizin konusu şu: iki metnin birbirine ne oranda benzeştiğine bakıp birbirinin aynı olma ihtimaline bakacağız. Amacımız kurumumuzdaki veri yönetişimi faaliyetlerinden biri olan iş sözlüğümüzü oluşturmak. Bunu oluştururken de sözlüğe duplike(mükerrer) terimlerin girilmesini engellemek. Ancak bazen sözlüğe girilen terimler aslında aynı terim olduğu halde yazılışları farklı olabilmekte(ÖR: KK&#39;lı müşteri adedi ve KK müşteri adedi), o yüzden klasik duplike bulma yöntemleri işimize yaramıyor. Bunun için yakın eşleşme(fuzzy match) yapmayı sağlayan kütüphaneler var. Bunun için Python&#39;da <strong>fuzzywuzzy </strong>kütüphanesini kullanıyoruz ancak bunun kurulumu biraz alengirli, o yüzden ben basit olması adına <strong>difflib </strong>kütüphanesini kullanacağım. Zaten bu detaylar şuan sizin için önemli dğeil, siz .Net kısmına odaklanın.</p>
        <p>Bunu yaparken de elimizde bi metin listesi olacak. Bu listeyi, kendisiyle karşılaştıracağız. Karşılaştırma işlemini pythona yaptıracağız. Bu karşılaştırmayı yaptırırken bazı zıt kelimelere bakmasın isteyeceğiz. Ör:&quot;KK limiti <strong>azalan</strong> müşteri adedi&quot; ve &quot;KK limiti <strong>artan </strong>müşteri adedi&quot; terimlerini bizi boşuna göstermesin, zira çıkan listede bunlar büyük oranda benziyor görünecek, bizi gereksiz meşgul etmiş olacak. Bunlar için bir istisna kelimeler listesine ihtiyacımız olacak.</p>
        <p>Json işlemlerini yapabilmemiz için nuget&#39;tan <strong>Newtonsoft.Json</strong> kütüphanesini indiriyoruz. </p>
        <p>Aşağıdaki gibi bir formumuz olacak.</p>
        <p>
            <img alt="fuzzy" src="../../images/vsto_pythonfuzzy.jpg" style="width: 802px; height: 482px" /></p>
        <p>Terim listesi dosyasının içeriği de şöyle:</p>
        <p>
            <img alt="" src="../../images/vsto_pythonfuzzy1.jpg" style="width: 226px; height: 205px" /></p>
        <p>
            İstisna kontrolü sayesinde 4. ve 5.satırdaki ile 8. ve 9.satırdakiler için karşılaştırma skoru görmeyeceğiz.</p>
        <p>Kodlarımıza bakalım:</p>
        <p>Az önce belirttiğim gibi, burda ilaveten bir dictionary&#39;yi json objesie haline dönüştürerek python&#39;a gönderiyoruz, python da gelen bu datayı alıp kendi dictionary formatına çevirecek. Bu arada tabiki istersek pythondan da bir veri yapısını json olarak c#&#39;a alıp, onu deserialize ederek bir .Net objesine(Dictioonary de olabilir başka bir yapı da) döndürebiliriz. Bu örnekte sadece biz python&#39;a gönderimde bulunacağız, python&#39;dan birşey almayacağız.</p>
        
        <pre class="brush:csharp">
public partial class frmFuzzy : Form
{
    Dictionary<string, string> dict = new Dictionary<string, string>();
    public frmFuzzy()
    {
        InitializeComponent();
    }

    private void frmFuzzy_Load(object sender, EventArgs e)
    {   
        dict.Add("artan", "azalan");
        dict.Add("tl", "yp");
        dict.Add("aktif", "inaktif");

        var result = from d in dict
                        select new { d.Key, d.Value };
        this.dataGridView1.DataSource = result.ToList();//kullanıcıya istisna listesi içeriği hakkında bilgi veriyoruz, istenirse buradan yeni key-value ikililieri de girilecek şekilde ayarlanabilir
    }

    private void button1_Click(object sender, EventArgs e)
    {
        string pyfile = @"E:\OneDrive\Dökümanlar\GitHub\dotnet\Ugulamalar\InvestPY\vstofuzzy.py";
        string istisnaJson = JsonConvert.SerializeObject(dict).Replace("\"","\\\"");//Dicitionar>Json dönşüm işlemi burada. Jsonda özel anlamı olan " işaretlerini \" şeklinde gönderiyoruz ki bunları gerçek " gibi algılasın
        Process process;
        ProcessStartInfo start = new ProcessStartInfo(this.txtPythonexe.Text); //Filename propertysi yerine direkt yaratım sırasında da parametre verebiliyoruz
        start.Arguments = string.Format("{0} \"{1}\" \"{2}\" {3} \"{4}\" {5}", pyfile, this.txtSource.Text, this.txtKolon.Text, this.txtEsik.Text, this.txtTarget.Text, istisnaJson); //source, kolon ve target kolonlarında boşluk olabilir diye ilave tırnak ekliyoruz
        start.UseShellExecute = false;
        start.CreateNoWindow = true;
        start.RedirectStandardOutput = true;
        start.RedirectStandardError = true;

        using (process = Process.Start(start))
        {
            using (StreamReader reader = process.StandardOutput)
            {
                string stderr = process.StandardError.ReadToEnd();
                string result = reader.ReadToEnd();

                if (string.IsNullOrEmpty(stderr))
                    MessageBox.Show(string.Format("Sonuç:{0}", result));
                else
                {
                    MessageBox.Show(string.Format("Sonuç:{0}", stderr));
                }
            }
        }

    }
}</pre>
        <p>Python tarafında bu parametreleri aşağıdaki gibi karşılılyoruz</p>
        <pre>
sourceworkbook=sys.argv[1]
kolon=sys.argv[2]        
esik=int(sys.argv[3]) 
targetfile=sys.argv[4]       
istisnaJson=sys.argv[5]</pre>
        
        <p>Kod çalıştırılınca sonuç şöyledir:</p>
        <p>
            <img alt="fuzzy" src="../../images/vsto_pythonfuzzy2.jpg" style="width: 436px; height: 111px" /></p>
        <p>Beklediğimiz sonuçları aldık.</p>
        <p>Böylece bir konunun daha sonuna glemiş olduk.</p>
        <p>Bu arada, bu kodun daha gelişmiş halini(eşanlam kontrolü, kelime köklerini alarak kontrol etme v.s gibi kontrollerin de olduğu) tamamen c# içinde kalacak şekilde hiç pythona bulaşmadan da yaptım. Onu <a href="OrnekProjeler_ISSozluguicinFuzzyMatchveNotasyonKontrolu.aspx">şurada </a>bulabilirsiniz.</p>
    </div>
</asp:Content>
