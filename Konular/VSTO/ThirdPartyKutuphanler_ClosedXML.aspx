<%@ Page Title='ClosedXML' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>


<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'></div><h1>Third Party Kütüphaneler</h1>
    <p>Şu ana kadar VSTO&#39;nun temel prensiplerini gördük, ve siz de aldığınız c# molasından sonra buraya gelmiş olmalısınız. </p>
    <p>Daha önce belirttiğmiz gibi, VSTO ile çalışırken Excel&#39;in kendi API&#39;si ile çalışabileceğimiz gibi, bazı işlemlerin sadeleştirildiği ve yeni özelliklerin eklendiği 3rd party kütüphanelerle çalışabileceğimizi söylemiştik. Bu bölümde bunlardan ilki olan ClosedXML&#39;i göreceğiz. Ama öncesinde bir 3rd party paketi kurmak için Nuget Manager nasıl kullanılıyor ona bakacağız.</p>
    <p>Örnek uygulamayı indirmek için <a href="https://github.com/VolkiTheDreamer/excel/tree/main/VSTO_3rdPartyLibs">tıklayınız</a>.</p>
    <h2 class='baslik'>NuGet Package Manager</h2><div class='konu'><p>Tools menüsünden aşağıdaki seçimi yapalım.</p>
    <p>
        <img alt="nuget" src="../../images/vsto_nuget1.jpg" style="width: 591px; height: 342px" /></p>
    <p>Açılan pencerede <strong>Browse </strong>sekemsine tıklayıp &quot;Closed&quot; yazdığınızda en tepede ClosedXML çıkacaktır. Herhangi bir paketi kurarken download adedine bakmanızda fayda var. Mesela ClosedXML, ben şu an işlem yaptığımıda 7,33 mio kez indirilmiş durumda, gayet iyi bi rakam. Demek ki oldukça popüler ve güvenilir.</p>
    <p>
        <img alt="closedxml" src="../../images/vsto_nuget2.jpg" style="width: 817px; height: 203px" /></p>
    <p>Bunu seçin ve hemen arkasından sağda açılan küçük pencerede kutuyu tıklayın ve Install diyin.</p>
    <p>
        <img alt="nuget" src="../../images/vsto_nuget3.jpg" style="width: 431px; height: 344px" /></p>
    <p>Sonrasında, size bu kütüphane ile birlikte başka nelerin kurulacağına dair bir pencere çıkacaktır. Bunlara <strong>dependency</strong> denmekte olup, bu kütüphanenin çalışması için gerekli diğer kütüphaneler anlamına gelmektedir. Bazı durumlarda zaten sizde varolan bir paketi görebilirsiniz ama dikkat edin versiyon numarası sizdekinden farklıdır. O yüzden bunlara da ok diyip ilerleyin.</p>
    <p>
        <img alt="nuget" src="../../images/vsto_nuget4.jpg" style="width: 482px; height: 484px" /></p>
    <p>Sonsrasında Output penceresinde kurulum adımlarını görebilirsiniz. En altta <strong>==Finished== </strong>ifadesini gördükten sonra emin olmak için Solution Explorer&#39;da References altına bakabilirsiniz.</p>
    <p>
        <img alt="nuget ref" src="../../images/vsto_nuget-ref.jpg" style="width: 332px; height: 235px" /></p>
    <p>Şimdi artık bu kütüphaneyi kullanmaya geçebiliriz.</p>
    </div>
    <h2 class="baslik">ClosedXML</h2>
    <div class="konu">
        <p>Öncelikle ClosedXML&#39;in nasıl kullanıldığını, ne tür fonksiyonalitesi olduğunu görmek için bu paketin <a href="https://github.com/ClosedXML/ClosedXML">github repo&#39;</a>suna gidelim. Bu sayfada çok basit bir örnek vermişler. Önce bunu inceleyelim, akabinde daha fazla bilgi için bizleri yönlendirdikleri <a href="https://github.com/closedxml/closedxml/wiki">wiki sayfasına</a> bakacağız. Bu kütüphanede işin özü, dosya kapalıyken onda işlem yapmaktır.</p>
        <pre class="brush:csharp">
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}</pre>
        <p>
            Buradaki using bloğunun ne olduğunu bildiğinizi varsayıyorum(c# molası verdiğinizde bunu da öğrenmiş olmanız lazım). Özetle, çeşitli yaratma işlemlerinde sonra sonlandırma işlemini de (unutulma durumuna karşın) otomatik olarak yaparlar. Burda bir Worobook nesnesi (bellekte) yaratılıyor. Sonlandırma işlemi de bellekte olmak durumunda, o yüzden using bloğuna alınıyor. </p>
        <p>
            İlk satırı anlatmaya devam edelim. Workbook nesnesi yaratılıyor ama burada gerek Interop gerek klasik Excel kullanımında aşina olduğumuzun aksine, sayfasız bir workbook oluşuyor. Garip, değil mi?</p>
        <p>
            Peki, sonrasında bir worksheet nesnesi yaratılıyor ve yaratım anında sayfaya isim de verebiliyoruz.</p>
        <p>
            Kodun kalanı oldukça basit ve anlaşılır, farklı bir <strong>Cell </strong>kullanımı dışında interop ile de aynı.</p>
        <p>
            Burda farkettiyseniz, hiç <strong>Globals</strong> sınıfından veya <strong>Application </strong>nesnesinden eser yok. Zira bu aslında bir VSTO Add-in uygulamasına ait bir kod değil, herhangi bir c# projesinde kullanılabilecek bir koddur. Projemizi başlatırken &quot;Excel VSTO Add-In&quot; veya &quot;Excel VSTO Workbook&quot; seçmediysek bunlar VSTO değildir. Ancak Excel ile çalışmak için illa VSTO add-in yaratmak zorunda değiliz. Hatta Interop API&#39;sini bile VSTO dışı projelerde kullanabiliriz. Yani şu kombinasyonlar olasıdır:</p>
        <ul>
            <li>Interop API&#39;sini kullandığımız bir VSTO add-in projesi</li>
            <li>Interop API&#39;sini kullandığımız VSTO olmayan bir proje</li>
            <li>3rd Party Excel API&#39;si ile VSTO projesi</li>
            <li>3rd Party Excel API&#39;si kullanarak VSTO olmayan bir proje</li>
            <li>Her iki API&#39;yi kullandığımız bir VSTO projesi</li>
            <li>Her iki API&#39;yi kullandığımız VSTO olmayan bir proje</li>
        </ul>
        <p>
            Proje tipine VSTO dediğimizde, VS bizim için arka planda bir sürü ayarlama yapar, tek olayı bu. Hatta istersek(neden isteyelim ki) genel bir c# projesi açıp bunu da VSTO&#39;ya dönüştürebiliriz. Ve bunu istersek(yine neden isteyelim ki) notepad&#39;de bile yapabiliriz.</p>
        <p>
            Bence Excel API&#39;sini 3rd party paketlerden ayıran en önemli özelliği, kodun çalıştığı PC&#39;de Excel&#39;in kurulu olması gerektiği ve bir workbookla çalışırken onun açık olması gerektiğidir. 3rd party paketlerde ise işlemler genelde bellekte, yani dosyalar kapalıyken, yapılır. Hatta bu sayfada tanıyacağımız ClosedXML&#39;in adından bile bu anlaşılıyor. Bununla beraber dosya açıkken de işlem yapılabilir, ki bizim durumumuzda her iki seçenek te olacak; bazen aktif workbook üzerinde işlem yapacağız, bazen de başka (kapalı) bir dosyaya yazma işlemi yapacağız. Ancak bu paketlerin nimetlerinden faydalanmak için dosya açık olsa bile önce onu kapatıp bellekte işlerimizi yapıp, en son dosyayı tekrar açabiliriz. Tabi açmak için Interop&#39;a başvurmamız gerkeiyor.</p>
        <p>
            Şimdi ClosedXML&#39;i daha detaylıca görmeden önce, yukarıdaki kodu Interop API ile nasıl yapardık ona bakalım.k ona bakalım.</p>
        <pre class="brush:csharp">
Excel.Workbook wb = app.Workwb.Worksheets[1].Name = "Sample Sheet"; //sayfa yaratmaay gerek yok, zaten default bir sayfamız var, biz bunun (1&#39;den fazla olsa da ilkini) adını değiştiriyoruz
app.Range["A1"].Value = "Hello World"; //bu ve alttakinde ise worksheet nesnesi üzerinden değil app nesnesi üzerinden erişiyoruzden erişiyoruz
app.Range["A2"].Formula = "=MID(A1,7,5)";
wb.SaveAs("HelloWorld_Interop.xlsx");
</pre>
        <p>
            Ben iki kodu da ribbonda iki butona atadım. ClosedXML yoluyla yapınca görünürde hiçbirşey olmadı(beklediğim üzere). Dosya arka planda oluştu. Interop ile yapınca, dosyayı bellekte değil direkt olarak o an açık olan Excel oturumu içinde yarattı, ve saveas yaptığıktan sonra da açık olarak kaldı. Aynı etkiyi ScreenUpdating=False diyerek de yapabilirdik, ama maksat hızlı çalışmaksa o zaman Interop yerine diğer kütüphaneleri tercih etmek daha doğru olacaktır. </p>
        <h3>
            Temel işlemlerin bir kısmı</h3>
        <p>Birkaç koddan sonra göreceksiniz ki, bu kütüphane kullanımını büyük ölçüde VBA syntaxına benzetmeye çalışmışlar, o açıdan güzel olmuş.</p>
        <p>
            Burda dikkat edilmesi gereken husus şu. Kütüphaneler zaman içinde evrimleşebiliyor. Verdikleri örneklerin bir kısmı ise güncelleme yapmadıkları için geçersiz olabiliyor. Bunları GitHub kullanmayı biliyorsanız Github üzerinden kütüphaneyi yaratanlara bildirebiliyorsunuz. Forumlarda aradığınızda bu problemin nasıl giderilleceğine dair bilgiler bulunabileceği gibi, &quot;What Is New?&quot; veya &quot;Changes&quot; gibi alanlarda da duyurusunu görebilirsiniz. Sorunu kendiniz de çözmeye çalışabilirsiniz tabi. Mesela aşağıdaki kodların bir kısmında sorun vardı, bunları yorum olarak ekledim.</p>
        <pre class="brush:csharp">
//Yeni dosya yaratma dosya yaratma
var wb = new XLWorkbook();//Sayfası olmayan bir dosyayı bellekte yaratır

//var olan dosyayı açma(bellekte)
var mevcut = new XLWorkbook("MevcutDosya.xlsx");

//Bir dosyaya sayfa ekleme
var ws = wb.Worksheets.Add("Yenisayfa1");
var ws2 = wb.AddWorksheet("Yenisayfa2");

//Range işlemleri: 
ws.Cell("A1").Value = "selam"; //[] değil () kullanıldığına dikkat, zira bunu bir indexli proerty olarak dğeil metod gibi ele almışlar
ws.Cell("A2").Value = new DateTime(1919, 1, 21);
var alan = ws.Range("B1:D20");
var ozel = alan.FirstCell(); // FirstCellUsed, FirstRow, LastColumn gibi çeşitli türevleri de var
var usedrange = ws.RangeUsed();

//format işlemleri
alan.Style.NumberFormat.NumberFormatId = 15; //Bu kütüphanenin dayandığı OpenXML'in öncenden tanımlı formatlarından
alan.Style.NumberFormat.Format = "$ #,##0";
//Zincirleme formatlama
alan.FirstCell().Style
    .Font.SetBold()
    .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

//Table işlemleri
var excelTable = alan.CreateTable();


//Lambda Expressions ile koşullu işlemler yapma
var rows = alan.Rows(r => r.Cell(3).GetString() == "E"); // 3.kolondaki değeri E olanlar
foreach (var row in rows) //wiki'deki örnekte doğrudan ForEach metodu kullanılmıştı ama bu metod List sınıfının bir metodudur, doğrudan kullanılamaz. Biz bunu şimdi klasik for ile yapalım, bir alttaki örnekte List'e çevirip öyle yapalım       
    row.Delete();

var hucreler = alan.Cells(c => c.DataType == XLDataType.Text); // XLDataType, daha önce XLCellValues idi, yeni versiyonda değişmiş
hucreler.ToList().ForEach(c => c.Style.Fill.BackgroundColor = XLColor.LightGray); //List'e çevirerek ForEach kullanmak da bir diğer yöntem      </pre>
        <p>Büyük ve/veya çok sayıda dosya ile çalışıyorsanız wiki&#39;deki performans yönetimi ile ilgili notları da mutlaka okuyun.</p>
        <p>
            Tabi hepsi bu kadar değil. Ben kendimce önemli gördüklerimi buraya aldım, başlangıç için bunlar yeterli olacaktır. Diğer bütün işlem tipleri için ihtiyaç duydukça wiki&#39;ye başvurabilirsiniz.urabilirsiniz.</p>
        <h3>Açık dosyalarla çalışmak</h3>
        <p>
            ClosedXML ile çalışırken o anda Excel&#39;de açık olan bir dosya ile çalışmak istersek, aşağıdaki gibi bir kod ile açık dosyayı elde eden bir kod yazarız. Akabinde bunu ana kodumuza dahil ederiz. Ben bu ve bunun gibi sık kullanılma ihitmali olan fonskyionları bir <strong>Utiliy</strong> paketi içine koydum(<strong><a href="https://github.com/VolkiTheDreamer/dotnet/tree/master/Ugulamalar/VolkansUtility">VolkansUtility</a></strong>), siz de bunu kullanabilirsiniz, aslında kullanmanızı şiddetle tavsiye ederim, çünkü oldukça faydalı kodlar var içinde.</p>
        <p>
            Normalde activeworkbook&#39;u bu kadar dolambaçlı bir şekilde elde etmeye gerek yok tabi. Bunun için hiç de Utility&#39;deki fonksiyona gerek duymadan doğrudan <strong>Globals.ThisAddin.Application.ActiveWorkbook</strong> diyerek de alabilirdik ancak, hem Utility içindeki kodu kullanmak daha kısa, hem de bu kodu VSTO dışındaki başka bir projede de kullanabilirsiniz. Üstelik, Utility içinde başka işimize yarayacak birçok hazır fonksiyon olacak. O yüzden şimdi Utility ile ilerleyeceğiz.</p>
        <pre class="brush:csharp">
public static Excel.Workbook GetActiveWorkbook()
{eWorkbook()
{
    Excel.Application app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
    app.Visible = true;
    return app.ActiveWorkbook;
}</pre>
        <p>
            Elimizde bu kod olduğuna göre şimdi ClosedXML kodu ile birleştirebiliriz.</p>
        <pre class="brush:csharp">
private void button4_Click(object sender, RibbonControlEventArgs e)
{
    //Önce Interop ile başlıyoruz, zira açık olan bi dosyada işlem yapacağız
    Excel.Workbook wb =  ExcelRW.GetActiveWorkbook();//VolkansUtility içinde
    string filepath = wb.FullName;
    wb.Save(); wb.Close(); //ClosedXML ile işlem yapabilmek için dosyayı geçici olarak kaydedip kapatıyoruz

    //Şimdi ClosedXML zamanı
    var cxwb = new XLWorkbook(filepath);            
    var ws = cxwb.Worksheet(1);
    var alan = ws.Range("B1:D20");           
    var hucreler = alan.Cells(c => c.DataType == XLDataType.Text);
    hucreler.ToList().ForEach(c => c.Style.Fill.BackgroundColor = XLColor.LightGray);
    cxwb.Save();
            
    //Şimdi tekrar Interop vakti
    app.Workbooks.Open(filepath);
}</pre>
        <p>
            Bu arada, kapalı bi dosyayı da açmak için yine Interop&#39;dan destek alabiliriz. Mantık yukarıdaki kod ile aynıdır.</p>
        <h3>
            Data işlemleri</h3>
        <p>Burdaki detayları anlamak için c#&#39;taki <strong>DataTable</strong> yapısı başta olmak üzere veri yapılarını iyi bilmeniz gerekiyor. </p>
        <p>
            Bir veri yapısındaki bilgileri Excel&#39;e nasıl alabileceğiniz
            <strong>Inserting Data</strong> başlığı altında açıkça anlatılmış durumda. Ben burdaki iki metoda ait küçük bir farktan bahsetmek istiyorum. Aslında wiki&#39;de bu bilgi var ama gözden kaçabilir diye ben de vurgulamak istedim. <strong>InsertData</strong> kolon başlıklarını eklemez ve Range döndürürken, <strong>InsertTable</strong> başlıkları koyar ve Table döndürür.</p>
        <p>Bunun dışında tersine ihtiyacınız olursa, yani Excel&#39;deki verileri bir veri yapısına veya bir DataTable&#39;a almak istiyorsanız bunun için yine benim Utility paketindeki bir metodu kullanabilirsiniz. Bu paket içinde ClosedXML ile Interop farkını görebileceğiniz iki metod var. <strong>WriteDataTableContentToActiveWBWithInterope</strong> ve <strong>WriteDataTableContentWithClosedXML</strong>. İkisinin satır sayısına bakarsanız 3rd party paketlerin nasıl kolaylıklar sağladığını görürsünüz. Tabi yalnız satır sayısı sizi yanıltmasın, Interop ile kod yazmak uzun sürüyor fakat süre olarak bakıldığında Interop daha hızlıdır, en azından bu örnekler için. Kodun içinde süre tutan bir kısım da vardır, siz de deneyebilirsiniz.</p>
       

    </div>
    <h2 class="baslik">EPPlus</h2>
    <div class="konu">

         <p>
             ClosedXML de dahil olmak üzere tüm Excel paketlerinden daha popüler bir pakettir. Ancak ticari kullanım için ücret isteniyor. Ücretsiz de olsa kullanmak için lisans ayarlaması da gerektiriyor. Bunları nasıl yapacağınız hepsi <a href="https://github.com/EPPlusSoftware/EPPlus">sayfalarında</a> anlatılmış durumda. Bunu çok başlarda kullanmıştım, o yüzden ClosedXML kadar hakim değilim ancak çok geniş bir <a href="https://github.com/EPPlusSoftware/EPPlus/wiki/Getting-Started">wiki dokümanı</a> var. Bol miktarda da <a href="https://github.com/EPPlusSoftware/EPPlus.Sample.NetFramework">örnek</a> yapmışlar. İncelemenizi tavsiye ederim.</p>
    </div>

    <h2 class="baslik">Diğerleri</h2>
    <div class="konu">
         <p>Başka kütüphanelere şöyle bi göz atma fırsatım oldu. İlki hariç çok detaylı kullanmadım, sizin incelemenize bırakıyorum.</p>
         <ul>
             <li><a href="https://github.com/ExcelDataReader/ExcelDataReader">ExcelDataReader&#39;ıııııııııııtaReader&#39;ı</a> Excel&#39;den hızlıca veri okumak için kullanabilrsiniz. Oldukça popüler bir kütüphane. Ben de bunu VolkansUtility paketinde kullanıyorum. Örnek için bu çalışmama bakabilirsiniz.(Bunun ExcelDataReader.DataSet diye bir yardımcı paketi de var, bunu da indirmeniz gerekiyor.</li>
             <li><a href="https://products.aspose.com/cells/net">Aspose.Cells</a>: Bu da bayağı popüler görünüyor.</li>
             <li><a href="https://github.com/paulyoder/LinqToExcel">LinqToExcel</a>: LINQ yapısını sevenler için ideal. c#&#39;a yeni başladıysanız şimdilik pas geçin.</li>
         </ul>
         <p>
             Ve daha başka birsürü kütüphane. Bence bu yukarda saydıklarım oldukça yeterli, daha fazlasına ihtiyacınız olmayacak diye düşünüyorum. Bu arada bunların hemen hepsi arka planda OpenXML denen kütüphaneyi kullanır. Bunu kullanarak biraz daha pratikleştirme yoluna gitmişler.</p>
         <p>
             Dikkat etmeniz gereken nokta şu olmalı. Kullandığınız paket bir süre sonra çok popüler olursa ücretli hale gelebilir(EPPLUS gibi). Çok seviyor ve memnunsanız devam edersiniz, veya ücretsiz bir başkasına geçersiniz, tabi yeni kütüphanyi baştan öğrenmeniz ve kodlarınızı buna çevirmeniz gerekebilir.</p>
        <p>
            &nbsp;</p>
    </div>
    </strong>
</asp:Content>
