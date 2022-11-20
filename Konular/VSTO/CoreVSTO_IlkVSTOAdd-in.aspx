<%@ Page Title='İlk VSTO Add-in' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'>
        <table>
            <tr>
                <td>
                    <asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td>
                <td>
                    <asp:Label ID='Label2' runat='server' Text='Görsel Araçlar'></asp:Label></td>
                <td>
                    <asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td>
            </tr>
        </table>
    </div>

    <h1>İlk VSTO Add'inimiz</h1>
    <h2 class="baslik">Giriş</h2>
    <div class='konu'>
        <h3>Kapsam</h3>
        <p>VSTO&#39;da <strong>iki temel çalışma </strong>yapılabiliyor. Bunların ilkinde her tür detayı ele alacağımız örnek çalışmalar olacak. Bu örnek çalışmaları sadece c# için hazırlayacağım. Bununla birlikte, Vb.Net kodunun c# kodundan önemli ölçüde ayrıştığı yerlerde bu kodları da vermeye çalışacağım(Yine de aşağıdaki indirilebilir dosyalar içinde Vb.Net versyionunu bulabilirsiniz). Bazı kısımlar deneyimli .Net kullanıcıları için sıkıcı gelebilir, hatta hatalı ifade ettiğim(veya gereksiz uzun yazdığım) yerler bile olabilir(zira c# konusunda kendimi master görmüyorum), onlar da İletişim menüsünden benle temasa geçerek bilgi verirlerse sevinirim. Bazı kısımlar da deneyimli VBA kullanıcılarına basit gelebilir, gerçi ben herkesi deneyimli VBA kullanıcısı olarak kabul etmek durumundayım. Ancak onlar da .Net dünyasının güzelliklerini keşfedince hayranlıklarını gizleyemeyecekler.</p>
        <p>İkinci tür çalışmada ise kapsamlı örnekler yapacağız., bunları sadece c# ile yapacağım.</p>
        <h3>Çalışma türleri</h3>

        <p>Bahsettiğim iki tür çalışma şunlar: <strong>Document Level(Excel VSTO Workbook) ve Application Level(Excel VSTO Add-in)</strong>. Ben daha çok Application Level&#39;e odaklanacağım ancak özellikle büyük bir proje şeklinde bir çalışma olacaksa bunu Document Level yapıp TaskPane gibi çeşitli fonksyonalite de katabilirsinz.</p>
        <p>İkisi arasındaki temel fark şu: Application Level projeler, sanki Personal.xlsb dosyasına yazdığınız kodlar veya Excel Add-in&#39;ler gibiyken, Document Level projeler ise normal bir workbook içine yazdığınız kodlar gibidir, yani o dosyaya özgüdür.</p>
        <h3>Kütüphaneler</h3>
        <p>Daha önce bahsettiğimiz gibi iki ana kütüphane var:</p>
        <ul>
            <li><strong>Microsoft.Office.Interop.Excel</strong>: Standart Excel nesne modeli ile iletişim kurmayı sağlar. Ağırlıklı olarak bunu kullanacağız. </li>
            <li><strong>Microsoft.Office.Tools.Excel</strong>: VSTO, Excel nesne modelini genişletir, yani normalde olmayan üyeleri ekler. Mesela artık ihtyiaç kalmayan get_xxxxx metodlarını bu kütüphane sağlıyordu. Biz bunu hiç kullanmaycağız diyebilirim. (Ben şimdiye kadar denemek dışında hiç kullanmadım)</li>
        </ul>
        <p>Bu arada yine bir kenara kaydetmenizi isteyeceğim linkler aşağıda. İlerleyen zamanlarda bunlara bakmanızda fayda var.</p>
        <ul>
            <li><a href="https://docs.microsoft.com/tr-tr/visualstudio/vsto/features-available-by-office-application-and-project-type?view=vs-2019">https://docs.microsoft.com/tr-tr/visualstudio/vsto/features-available-by-office-application-and-project-type?view=vs-2019</a></li>
            <li><a href="https://docs.microsoft.com/tr-tr/visualstudio/vsto/general-reference-office-development-in-visual-studio?view=vs-2019">https://docs.microsoft.com/tr-tr/visualstudio/vsto/general-reference-office-development-in-visual-studio?view=vs-2019</a></li>
            <li><a href="https://docs.microsoft.com/tr-tr/visualstudio/vsto/walkthroughs-using-excel?view=vs-2019">https://docs.microsoft.com/tr-tr/visualstudio/vsto/walkthroughs-using-excel?view=vs-2019</a></li>
            <li><a href="https://docs.microsoft.com/tr-tr/visualstudio/vsto/common-tasks-in-office-programming?redirectedfrom=MSDN&amp;view=vs-2019#projects">https://docs.microsoft.com/tr-tr/visualstudio/vsto/common-tasks-in-office-programming?redirectedfrom=MSDN&amp;view=vs-2019#projects</a></li>
        </ul>
    </div>

    <h2 class="baslik">Application Level projeler</h2>
    <div class="konu">
        <p>Şimdi ilk projemizi Application Level olarak yapacağız. Bunlar, adı üzerinde tüm Excel seviyesinde geçerli olurlar, tek bir workbook değil. En tipik örneği kendisine özel bir Ribbon Menüsünün eşlik ettiği projelerdir. Benim Excelent projem böyle bir projedir.</p>
        <h3>Kaynak Dosyalar</h3>
        <p><a href="Giris_VisualStudio.aspx">Visual Studio</a> sayfasında başladığmız ilk örnek üzerinden devam edelim. Projenin tamamını(Vb.Net versiyonunu da)&nbsp;<a href="https://github.com/VolkiTheDreamer/excel">buradan</a> indirebilirsiniz. Aslında burada Excel&#39;le ilgili birçok örnek uygulama yer alacak. O yüzden repository&#39;i komple tek seferde indirmek isteyebilirsiniz diye tüm repo linkini verdim. Klasör isimlerinden hangi projede çalıştığımızı anlayabilirsiniz diye düşünüyorum.</p>
        <p id="github">Bununla birlikte tek tek klasör indirerek gitmek istiyorsanız bunları indirmek için ara çözümlere ihtiyaç var. Özellikle yeni başlayan biriyseniz ve hazır toollara sahip değilseniz linklerde çıkan klasör isimlerini<strong> <a href="https://minhaskamal.github.io/DownGit/#/home">https://minhaskamal.github.io/DownGit/#/home</a> </strong>sayfasına yapıştırmanız durumunda bir zip dosyası inecektir.</p>
        <h3>Adım adım App Level proje oluşturma</h3>
        <p> 
            <img alt="" src="/images/VSTO_vs1.jpg" class="zoomla" /></p>
        <p>Bu pencereyi analiz edelim.</p>
        <p>En üstte, VS tarafından otomatik olarak eklenen Using deyimleri bulunuyor. Onun altında da namespace tanımı var. Namespaceler sınıfları içeren containerlar olarak düşünülebilir. Genelde projemizin adıyla aynı olurlar. Biz şimdi bunun içine odaklanalım.</p>
        <p>İlk satırdaki public partial class ifadesi, aslında bir sınıf oluşturulduğunu ancak bunun bir kısmının bu dokümanda bulunduğunu anlatır. Sınıfın kalan kısmını görmek için sağ taraftaki&nbsp; <strong>Solution Explorer&#39;da Show All Files </strong>butonuna basın ve <strong>ThisAddIn.Designer.cs </strong>dosyasına tıklayın. Bu desinger dosyaları oldukaça karışıktır, neyseki bizim onunla bi işimiz yok, VS sağolsun, bizim için işin hammaliyet kısmını hallediyor. Keza yukarıdaki kod bloğunda <strong>VSTO generated code </strong>yazan kısımda da yine VS bizim için birşeyler oluşturmuş durumda. Buralara hiç takılmadan devam ediyoruz.</p>
        <p>
            <img alt="" src="/images/vsto_showallfiles.jpg" /></p>
        <p>İki metod görüyoruz. Bu iki metod, isimlerinden de anlaşılacağı üzere Application seviyesinde event handlerlardır, yani Excel açıldığında ve kapandığında devreye girerler.</p>
        <p>Şimdi ilkinde araya küçük bir mesaj kutusu ekleyelim.</p>
        <pre class="brush:csharp">
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    MessageBox.Show("Merhaba Excel!");
}</pre>
        <p>
            Bu MessageBox satırını yazdığımızda altı kırmızılı çizecektir, zira henüz bu sınıfın dahil olduğu namespace yani System.Windows.Forms henüz programa dahil edilmedi, bunu <strong>using</strong> deyimi ile ekleyelim. Bunu manuel olarak ekleyebileceğiniz gibi, aşağıdaki gibi <strong>Show potential fixes&#39;</strong>a tıklayıp, &quot;using System.Windows.Forms;&quot; dersek de otomatik eklenir. Bu arada Forms namespace&#39;i aslında projemize dahil durumda. Projeyi ilk oluşturduğumuzda VS, belli bazı sınıfları otomatikman projemize dahil eder. Bunları Solution Explorer&#39;da References altında görebilirsiniz. Hatırlarsanız, biz using ifadesini ekleyerek bu sınıfın içindeki metodlara doğrudan erişim hakkı kazanıyorduk, uzun uzun yazmaktan kurtuluyorduk. Zaten Show potential fixes&#39;a tıkladığınızda bu seçeneği de görebilirsiniz. </p>
        <p>
            <img alt="" src="/images/vsto_usingforms.jpg" /></p>
        <p>
            Şimdi kodumuzu çalıştıralım. Bunun için <strong>Debug&gt;Start Debugging </strong>diyin veya doğrudan <strong>F5</strong> tuşuna basın. Bu arada Excelin o sırada kapalı olması başka add-in&#39;lerle karışmaması adına iyi olacaktır. Programı debug ettiğimizde Excel kendiliğinden açılacaktır.</p>
        <p>
            Excel açıldıktan sonra mesaj kutusunu göreceksiniz, onu kapatın ve Developer menüsünden COM Addi-ins menüsüne tıklayın, Add-inimiz orada göreceksinz.</p>
        <p>
            <img alt="" src="/images/vsto_comaddin1.jpg" /></p>
        <p>Şimdi, VS&#39;yu kapatsak ve sıfırdan bi Excel açsak bile bu Merhaba Excel mesajını görürüz, çünkü artık add-in&#39;imiz bilgisayarımıza kurulmuştur. Bunu kaldırmak istiyorsak add-in projemizi VS içinden <strong>Clean </strong>etmemiz gerkeir. Clean işlemi, derlenmiş bir programı sistemden kaldırır. Ancak clean işleminden önce gelin bir de ne tür dosyalar oluşmuş ona bi bakalım.</p>
        <p>Proje klasörü aşağıdaki gibidir. Burada sln uzantılı olan solution dosyası olup varsa birden fazla proje için kapsayıcıdır. Biz genelde bir solution içinde bir proje ile çalışacağız. csproj uzantılı dosya da proje dosyasmızı temsil eder. Bu ikisiyle de şuan bi işimiz yok. ThisAddIn.cs&#39;i zaten biliyoruz, zira orada çalışıyoruz.</p>
        <p>
            <img alt="" src="/images/vsto_klasor1.jpg" /></p>

        <p>Şimdi bir de bin klasörüne bakalım. Bunlardan dll dosyası bizim esas derlenmiş dosyamızdır, diğerleri kritik dğeil. VS&#39;da bir proje derlendiğinde, kendi başına bağımsız bir progam olacaksa <strong>exe </strong>uzantılı, başka bir programa bağımlı çalışacaksa <strong>dll </strong>uzantılı olur, ki bizim durumumuzda Excel&#39;e bağlı programlar derleyeceğiz.</p>
         <img alt="" src="/images/vsto_klasor2.jpg" /><p>Bin klasörü dışında bir de aynı dosyaları içeren obj klasörü vardır. Bu bin ve obj folderlarının ayrımı ve anlamı teknik bir konu olup şuan bilmenize gerek yok. Merak edenler <a href="https://stackoverflow.com/questions/5308491/what-are-the-obj-and-bin-folders-created-by-visual-studio-used-for">bu sayfadan </a>bakabilirler. Bu arada her iki klasörün içinde de bir <strong>Debug </strong>bir de <strong>Release </strong>klasörü vardır. Geliştirme aşamasında hep Debug modunda olacağız, zten projemiz ilk açıldığında default olarak bu mode seçilidir(En yukarıdaki pencerede Test menüsünün hemen altında Debug yazan yere bakın). Ne zamanki projemiz artık yaygınlaştırmaya hazır, o zaman Release moda geçeriz. Özet farkı şu: VS, bizim için debug modda kullanılmayan bazı optimizasyon çalışmaları yapar ve kodumuzu daha performanslı hale getirir. </p>
        <h3>Projeye Form ekleme</h3>
        <p>Şimdi projemize bir form ekleyip ortamı canlandıralım.</p>
        <p>Öncelikle projemize sağ tıklayıp, <strong>Add--&gt;New Item--&gt;Windows Form </strong>diyelim.(Formlar, sık kullanılan nesneler olduğu için direkt açılan context menüde de bulunur, doğrudan buradan da eklenebilir)</p>
        <p>
            <img alt="" src="/images/vsto_formadd.jpg" /></p>
        <p>
            Formumuza control ekleme işini başka zamana bırakalım. Şimdi bu formumuzun kod sayfasına bakalım. Bunun için formumuz açıkken F7 tuşuna basarak veya formun üzerinde sağ tklayıp <strong>View Code</strong> diyerek açabilirsiniz.</p>
        <pre class="brush:csharp">
public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
    }              </pre>
        <p>
            Gördüğünüz gibi Form1 isimli formumuz Form isimli base classtan türetilmiştir. Aradaki bu &quot;:&quot; işareti <strong>inheritance&#39;ı</strong> ifade eder, yani kalıtımı. Kalıtım, c# gibi pure object oriented dillerin ana unsurlarından biridir. Biz şimdilik bu detaya girmeyeceğiz, sadece ne olduğunu bilmenizi istediğim için bahsettim. İç kısımdaki kodda da Form1 nesnesi yaratılmakta(constructor metod ile). Şimdilik buna da takılmayın, bunları c#&#39;ı detaylı öğrenmeye başladığınızda incelersiniz. Şimdi bu formu nasıl açarız ona bakalım.</p>
        <p>
            Form1 diye oluşturduğumuz form aslında bir sınıf olup bu sınıfı projemizin herhangi bir yerinde kullanabilmek için bu sınıftan bir Form1 nesnesi yaratmamız gerekiyor. Bunun için yukarıda eklediğimiz mesaj satırı altına aşağıdaki iki satıra daha ekleyelim ve bu nesnenin Show metodu ile bu formu ekrana getirelim. Form1, bir sınıf olduğu için, bu sınıftan istediğimiz kadar nesne yaratabiliriz.</p>
        <pre class="brush:csharp">
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    MessageBox.Show("Merhaba Excel!");
    Form1 frm1 = new Form1();
    frm1.Show();
    Form1 frm2 = new Form1();
    frm2.Show();
    Form1 frm3 = new Form1();
    frm3.Show();
}      </pre>
        <p>Form konusuna burada ara verelim, sonra siz detaylı olarak incelersiniz, zira bunun VSTO ile doğrudan bi alakası yok, genel .Net konusudur.</p>
        <h3 id="ribbon">Ribbon ekleme</h3>
        <p>Az önce bi form ekledik, ama bu formu kaparsak mevcut durumda ona bir daha ulaşmanın yolu yok. İşte ribbon arayüzü burada imdada yetişiyor. Normal Excel Add-in&#39;lerde eski commandbar mantığıyla menüler yapabilyorduk, bunun dışında bir de <strong>Custom UI Editor</strong> diye birşeyden bahsetmiştik ama bunun örneğini yapmamıştık. İşte şimdi canlı, renkli, fonksiyonel menüler yapma zamanı. Üstelik Custom UI Editor tekniğine göre VSTO daha avantajlı, zira arakasında tüm .Net dünyası var.&nbsp; </p>
        <p>Şimdi Projemize sağ tıklayıp <strong>New Item--&gt;Ribbon(Visual Designer)</strong> diyelim.(XML&#39;li olan ile daha gelişmiş Ribbon&#39;lar tasarlanabilmekte, ancak bize şuan Visual Designer yeterlidir).</p>
        <p>
            <img alt="" src="/images/vsto_ribbonadd1.jpg" /></p>
        <p>Ribbonumuz aşağıdaki gibi projemize eklenir.</p>
        <p>
             <img alt="" src="/images/vsto_ribbonadd2.jpg" /></p>
        <p>
             Buraya hemen bir buton ekleyip, bu butona tıklanınca da bi adet form1 formu açılmasını sağlayalım. Tabi bu buton bi ribbon butonu olup sol taraftaki toolbaxın en üstüne <strong>ribbon controlleri</strong> diye bi group eklenmiş oldu, butonu oradan alacağız.</p>
        <p>
             <img alt="" src="/images/vsto_ribboncontrols.jpg" /></p>
        <p>
             Bu arada ribbon1.cs dosyasının içi şöyledir.</p>


    <pre class="brush:csharp">    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f = new Form1();
            f.Show();
        }
    }           &nbsp;</pre>

        <p>Gördüğünüz gibi bir tane de otomatik oluşan <strong>Ribbon1_Load </strong>event handlerı var. Buralara neler yazabiliriz, ribbonlarla ilgili başka ne tür detaylar var, bunlara Ribbon detay safyasında bakacağız. Şimdilik biz kodumuzu çalıştıralım.</p>
        <p><img alt="" src="/images/vsto_ribbonshow.jpg" /></p>
        <p>Gördüğünüz gibi şuanki haliyle ribbonnumuz çok da hoş olmayan bir şekilde Add-ins menüsü altına yerleşti. Tamam, bir Excel Add-in gibi eski moda görüntüsü yok, sonuçta elimizde görselliği daha gelişmiş olan bir menü var ama bu haliyle  yetersiz. Az önce belirttiğim gibi bunu detaylı olarak sonra ele alacağız.</p>
        </div>

    <h2 class="baslik">Document Level(Workbook seviyesi)</h2>
    <div class="konu">
    <p>Ben kendi çalışmalarımda document level&#39;a çok odaklanmadım açıkçası ama internette sıklıkla örnekleri olabildiği için bunun için de bi yer ayırmak istedim. İlgili dosyayı <a href="https://github.com/VolkiTheDreamer/excel/tree/main/VSTO_DocLevel">buradan</a> indirebilirsiniz. (<a href="#github">Yukarda</a> GitHub&#39;dan bir klasör nasıl indirilir, bahsetmiştim)</p>
        <p>Bunlarda da Ribbon yapılabilmektedir ancak internet örneklerinde daha çok ActionPane kullanımı olmaktadır. Sadece o dosyayla ilgili işlemler için Form açmak yerine Actionpane açmak daha makul olmaktadır. Biz de burda Actionpane kullanımına odaklanacağız.</p>
        <p>Konu bütünlüğü adına detayları aşağıya ekledim ancak size tavsiyem, bu noktadan sonra devam etmeyin. Bunu bi yere not edin, Core VSTO bölümünü tamamen bitirdikten sonra tekrar gelin. </p>
        <h3>Projeyi yaratma</h3>
        <p>Proje menüsünde <strong>VSTO Workbook</strong> seçelim ve folder/file seçimimizi yapalım. Karşımıza VS içinde br Excel görüntüsü çıkacaktır.</p>
        <p>
            <img alt="doclevel" src="../../images/vsto_doclevel1.png" style="width: 1644px; height: 939px" /></p>
        <p>Sağ taraftaki ThisWorkbook.cs içine girelim ve Startup(VBA&#39;deki Workbook_Open prosedürüne benzer) metodu içine odaklanalım. Buraya, dosya açılır açılmaz ActionPane&#39;in yaratılmasını sağlayan bir kod yazacağız. Ama öncesinde daha basit birşey yazalım ve ActionPane&#39;i bir altta ele alalım.</p>
        <pre class="brush:csharp">
private void ThisWorkbook_Startup(object sender, System.EventArgs e)
{
MessageBox.Show("Merhaba Doc Level");
} </pre>
        <p>
&nbsp;Bu haliyle çalıştırırsak sanki Workbook_Open prosedürüne MsgBox(&quot;Merhaba Doc Level&quot;) yazılmış bir dosya gibi davranır.</p>
        <h3>ActionPane</h3>
    <p>Konuya geçmeden önce Actionpane&#39;lere benzeyen Taskpane&#39;lerden de bahsedeyim. İkisi görüntü olarak aynıdır, Actionpaneler Document Level&#39;lda kullanırken Taskpaneler Application Level&#39;da kullanılır. Daha detay bilgi için <a href="https://docs.microsoft.com/tr-tr/archive/blogs/ericwhite/understanding-the-difference-between-custom-task-panes-and-action-panes">buraya</a> bakabilirsiniz.</p>
    <p>Eklemek için:</p>
    <ul>
        <li>Taskpanelerde olduğu gibi yine Projeye sağ tıklayıp <strong>Add New Item&gt;User Control </strong>diyeceğiz. Adı da MyUsercontrol olsun.</li>
        <li>Bu kontrolün içine bir combobox ve bir de button sürükleyip kaydedelim</li>
        <li>ThisWorkbook.cs içine gelelim ve aşağıdaki kodu yazalım</li>
    </ul>
    <pre class="brush:csharp">
private void ThisWorkbook_Startup(object sender, System.EventArgs e)
{
//MessageBox.Show("Merhaba Doc Level");
uc = new MyUsercontrol();            
this.ActionsPane.Controls.Add(uc);
}      &nbsp;</pre>
    <p>
        Çalıştıralım ve sonucu görelim.</p>
    <p>
        &nbsp;</p>
        <p>
        <img alt="" src="/images/vsto_actionpane.jpg" /></p>
        <p>
            Olur da bu actionpane&#39;i kapattık, sonra bir kez daha açabilmek için bir arayüze ihtiyacımız vardır. Bunun için bu dokümana özgü olarak bir ribbon yaratıp, ordan bu actionpanei açıp kapayan bir togglebutton ekleyebiliriz.</p>
    <pre class="brush:csharp">
private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
{
if (this.toggleButton1.Checked)
{
    Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = true;
}
else
{
    Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = false;
}
}  </pre>
        <p>
            Bu kodu çok daha pratik şekilde yazabileceğinizi biliyorsunuz. Hadi bunu tek satıra indirelim.</p>
        <pre class="brush:csharp">Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = this.toggleButton1.Checked;</pre>
        <p>
            Bu arada, eklediğimiz ribbonun nerde göründüğüne dikkat ettiniz mi? Bunun cevabını bilmiyorsanız yukardaki tavsiyemi dinlememişsiniz, yani CoreVSTO&#39;yu bitirmeden devam etmişsiniz demektir. Zira bu detayları <a href="CoreVSTO_Ribbon.aspx">Ribbon</a> konusunda göreceğiz.</p>
        <p>
            Document Level projelerle ilgili olarak başka birşey görmeyeceğiz. Detaylı bilgiye <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/architecture-of-document-level-customizations?view=vs-2019">buradan</a> ulaşabilirsinz ancak zaten çoğunlukla Application Level bilgiler yeterli olacaktır.</p>
</div>

</asp:Content>
