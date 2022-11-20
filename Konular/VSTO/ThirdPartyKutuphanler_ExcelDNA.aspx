<%@ Page Title='ExcelDNA' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'></div>
    <h1>ExcelDNA</h1>
    <p>&nbsp;</p>
    <p><a href="https://excel-dna.net/">ExcelDNA</a>, diğer kütüphanelerden fakrlı olarak aslında Interop için de bir altenratif. Diğerleri daha çok Excel kapalıyken okuma/yazma yapan kütüphanelerken, ExcelDNA ise daha çok XLL tarzı add-inler ve özellikle de UDF&#39;ler yapmak için kullanılıyor.</p>
    <p>Buna benzer <a href="https://www.add-in-express.com/">Addın Express</a>, <a href="https://www.planatechsolutions.com/xllplus/">XLL Plus</a> ve <a href="https://www.spreadsheetgear.com/">spreadsheetgear</a> gibi araçlar da var, ancak bunlar ücretli olduğu için burda değinmeyeceğiz. İsteyen araştırabilir. Bu arada bunlardan özellikle ilkinde oldukça faydalı makalelere rastlama ihtimaliniz yüksek, hatta ben de sitemde yer yer onlardan link verdim.</p>
    <h3>UDF&#39;ler</h3>
    <p>Farkettiyseniz, VBA&#39;de oldukça önem veridğim UDF konusuna şimdiye kadar hiç değinmedik. Çünkü VSTO, <strong>UDF&#39;leri tam olarak desteklemiyor</strong>. Bununla birlikte bu sorunu aşmanın birkaç yolu bulunuyor.</p>
    <ul>
        <li>COM add-in. Biz genel olarak VSTO add-inlere COM add-in dedik ama aslında bu ikisi biraz farklıldır. COM Add-inler hakkında <a href="http://www.cpearson.com/excel/COMAddIn2007.aspx">burdan</a> bilgi edinebilirsiniz, ama kafanızı çok da karıtşırmayın, biz bunla ilerlemiyicez.</li>
        <li>Automation Add-in(ExcelDNA&#39;le alakası yok ama konu bütünlüğü adına buraya koydum. Biz ilk olarak bunu göreceğiz)</li>
        <li>XLL addin(bunu da bir sonraki başlıkta göreceğiz). Excel normal koşullarda <a href="https://en.wikipedia.org/wiki/Managed_code">managed code</a>(.Net dili) kullanan bir XLL add-in yazmayı desteklemez. &quot;İlle de C/C++ kullanacaksın&quot; der. Ama demokrasilerde çare tükenmez. ExcelDNA bunun için var. Teşekkürler 
Govert van Drimmelen. (Bu arada <strong>Python</strong> dilini bilenler <a href="https://www.pyxll.com/">pyxll</a> ile de XLL add-in yazabilir)</li>
    </ul>
    <p>
        COM Add-in ve Automation addinler hakkında daha fazla bilgiyi <a href="https://support.microsoft.com/en-us/topic/excel-com-add-ins-and-automation-add-ins-91f5ff06-0c9c-b98e-06e9-3657964eec72">burada</a> bulabilirsiniz.</p>
    <h2 class='baslik'>Automation Add-inlerle UDF yazma</h2>
    
    <div class='konu'>
        <p>Öncelikle şunu söylemeden geçemiyceğim. Bu konuda da tıpkı &quot;<a href="DigerUygulamalarlaCalismak_OfficeUygulamlariylaCalismak.aspx">GC vs Marshall.Release</a>&quot; konusunda olduğu gibi çok kirli ve eksik bilgi var. Bir kaynakta olan detay bilgi diğerinde olmayabiliyor. Ben yine sizleri bu deli işinden kurtarmaya çalışacağım. Benim yaptıklarımı aynen yaparsanız sorunla karşılaşamsınız(Gerçi bu işler belli olmuyor. VS sürümü, Ofice sürümü, windows ve hatta genel olarak işletim sistemi sürümündeki farklılıklar hatalara neden olabilir, bunları bulup araştırmak size kalır) Mesela siz bu yazıyı 2025 veya sonrasında okuyorsanız, belki işler değişmiş olabilir, umarım ben de güncellemiş olurum.</p>

        <p>Şimdi, bu yöntem ile hala klasik .Net ile kodumuzu yazıyoruz. Sadece birkaç ayar yapmamız gerekecek.</p>
        <p>Öncelikle olaya bütünsel yaklaşalım. Birçok kaynakta bu yok, ben baştan vererek sizi zahemetten kurtarayım. Bu sitenin VSTO bölümünüde olduğunuza göre bir VSTO add-in yapmışsınızıdır veya yapmaya çalışıyorsunuzudur. Bu add-inle birlikte kullanılacak da çeşitli UDF&#39;ler yapmak istiyorsunuz. Senaryo muhtemelen şöyledir(ki bende böyle olmuştu.) Excel Add-ininizi COM/VSTO Add-in&#39;e çevirmek istediniz. Tüm menüleri v.s çevirdiniz, ama UDF&#39;leriniz kaldı. Şimdi de onları .Net&#39;e aktarmak istiyorsunuz. Yani özetle UDF&#39;lerinizi aynı VSTO add-in içinde kullanacaksınız. Yani add-inizin yüklendiğinde UDF&#39;leriniz de devreye girsin istiyoruz. Süreç basitçe şöyle işler.</p>
        <ul>
            <li>Öncelikle normal VSTO Add-ininzi yazarsınız. Bu bölüme kadar bunu nasıl yapacağınızı öğrendiniz zaten</li>
            <li>Şimdi <strong>Solution&#39;ımıza ikinci bir proje</strong>(automation add-in) ekleyeceğiz. Bu önemli, çünkü ana VSTO add-in projesinde yer almayacak olan proje seviyesinde ayarlamalar yapacağız, bu yi,üzden ikinci bir projeye ihtiyacımız var.</li>
            <li>Sonra da VSTO Addin projemizin Thisaddin_startup kodu içine Automation add-in projemizdeki add-ini install eden bir kod yazacağız.(Bu aşama zorunlu değil, isterseniz kullanıcılara bunu manuel yapmalarını da söyleyebilirsiniz ama bu çok profesyonel bi yaklaşım olmaz)</li>
        </ul>
        <p>
            Şimdi daha detaylıca neler yapacağımıza bakalım.</p>
        <h3>Automation Add-in projesini oluşturma</h3>
        <p>
            Komple repository&#39;yi daha önce indirmediyseniz burdaki örnek uygulamayı <a href="https://github.com/VolkiTheDreamer/excel/tree/main/VSTO_UDF">şu</a> ve <a href="https://github.com/VolkiTheDreamer/excel/tree/main/MyUDFs">şu</a> olmak üzere iki ayrı linkten indirebilirsiniz, zira iki projemiz olacak.</p>
        <p>
            Adımlar şöyle:</p>
        <ul>
            <li>Öncelikle Visual Studionuz açıksa kapatın ve <span style="text-decoration: underline"><strong>admin</strong></span> olarak tekrar açın. Bunun için start mensünden veya taskbardan VS&#39;ya sağ tıklayıp <strong>Run as Administrator </strong>deyin. (Bu seçenek <strong>More </strong>altında olabilir)</li>
            <li>VSTO Add-in projenizi açın. Şimdi Solution&#39;a sağ tıklayıp <strong>Add&gt;New Project </strong>diyerek yeni bir <span style="text-decoration: underline"> <strong>Class Library </strong></span>ekliyoruz</li>
            <li>Projemizin properties&#39;ine sağ tıklayıp Build kısmını <span style="text-decoration: underline"> <strong>Register for COM Interop</strong></span> yapıyoruz.&nbsp; <img src="../../images/vsto_udf1.jpg" style="width: 793px; height: 627px" /></li>
            <li>Hala build içindeyken <span style="text-decoration: underline"> <strong>Platform Target</strong></span> kısmını da hangi Excel versiyonunu hedefliyorsanız onu seçin. Ben 64 bit Excel kullandığım için 64 seçiyorum.(Bu kısım birçok kaynakta atlanmış, siz atlamayın derim, zira hata alabilirsiniz, veya deneyin belki sizde çalışır. Ben sadece bir kaynakta gördüm, başkaları belirtmediğine göre onlarda farklı bir nedenle hata çıkmıyor olabilir ve sizde de onların benzeri bir durum olabilir)<br />
                <img src="../../images/vsto_udf3.jpg" style="width: 568px; height: 236px" /></li>
            <li>Yine properties içindeyken bu sefer Application kısmına gelip <span style="text-decoration: underline"> <strong>Make assembly COM-Visible</strong></span> yapıyoruz. (Yukardaki maddede parantez içinde belirttiklerim geçerli)<br />
                <img src="../../images/vsto_udf2.jpg" style="width: 602px; height: 430px" /></li>
            <li>Sonra Class1 isimli classımıza gider çeşitli kodlar yazarız. Öncelikle <strong>class&#39;ımızın adını değiştirelim</strong>. Ben <strong>MyFunctions </strong>koydum. Sonra bunun hemen önüne 3 adet <a href="https://www.buraksenyurt.com/post/C-Temelleri-Nitelikleri(Attributes)-Kavramak-bsenyurt-com-dan">attribute</a> yazıyoruz.<br />
                <br />
            <pre class="brush:csharp">
[Guid("3ADF6501-4D91-4B40-A374-23946CE29E6D")] //Bu GUID sizde farklı olacak
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]            </pre>
                <p>Burda bahsi geçen GUID'i <strong><span style="text-decoration: underline">kendi oluşturacağınız GUID</span> </strong>ile değiştirmeniz lazım. Bunu <strong>Tools</strong> menüsünden <strong>Generate/Create GUID</strong> diyerek yapabilrisiniz. Eğer bu menü sizde çıkmıyorsa bunu <a href="https://marketplace.visualstudio.com/items?itemName=kylebahrke.GenerateGUIDforVisualStudio2015">şuradan</a> indirebilirsinz.</p>
            </li>
            <li>Sonra class&#39;ımız içine bir method yazarız, ki bu bizim UDF&#39;imiz olacak. Mesela iki sayıyı toplayan basit bi fonksiyon olsun bu.<br />
                <br />
            <pre class="brush:csharp">
public double Topla(double number1, double number2)
{
    return number1+number2;
}           </pre></li>
            <li>Şimdi bir de Excel bağlantısı olan bir fonksiyon yazalım. Ne de olsa UDF&#39;imiz Excel&#39;de çalışacak. Bu fonksiyon da bir hücredeki metnin içinde kaç kelime olduğunu bulsun. Daha önce VBA&#39;le yaptığımız bir koddu bu.<br />
                <br />
                <pre class="brush:csharp">
public int KacKelime(Excel.Range hucre)
{
    string icerik = hucre.Value.ToString();
    return icerik.Split(' ').Length;
}
                </pre>
            </li>
            <li>Son olarak da yine classımızın sonuna şu 3 metodu yazarız. Bunları <span style="text-decoration: underline"><strong>aynen copy-paste yapın</strong></span> lütfen. Bunlar UDF&#39;imizi <strong>registry</strong> kayıt işlemleri içindir.<br />
                <br />
            <pre class="brush:csharp">
[ComRegisterFunctionAttribute]
public static void RegisterFunction(Type type)
{
    Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
    RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
    key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
}

[ComUnregisterFunctionAttribute]
public static void UnregisterFunction(Type type)
{
    Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
}

private static string GetSubKeyName(Type type, string subKeyName)

{
    System.Text.StringBuilder s = new System.Text.StringBuilder();
    s.Append(@"CLSID\{");
    s.Append(type.GUID.ToString().ToUpper());
    s.Append(@"}\");
    s.Append(subKeyName);
    return s.ToString();
}           </pre>
            </li>
            <li>Hazırlık aşaması bu kadar. Şu an sadece bu projeyi build edersek Excel&#39;de kullanılabilir halde olduğunu görebilriz. Bunun için Projeye(Solution&#39;a değil) sağ tıklayarak <strong>build </strong>deyin. Veya build menüsünden <strong>Build ProjeAdı </strong>deyin.</li>
            <li>Excel&#39;i açın(açıksa kapatıp tekrar açın) ve <strong>Developer&gt;Addins </strong>altındaki <strong>Automation</strong>&#39;a tıkalyın<br />
                <img src="../../images/vsto_udf5.jpg" style="width: 281px; height: 389px" /></li>
            <li>Burda UDF&#39;lerimizin olduğu sınıfı Namespace.ClassAdı şeklinde görmeliyiz. Göremiyorsak bi sorun olmuştur, yukardaki adımlardan birini atlamışsınızdır.<br />
                <img src="../../images/vsto_udf4.jpg" style="width: 481px; height: 412px" /></li>
            <li>Bunu seçtikten sonra add-inimiz Excel Add-in&#39;leri içine gelir. Aşikar ki, bu UDF&#39;leri kullanabilmek için bunun tick işareti seçili olması gerekir. Bu yaptığımız son 2-3 adım, ilgili add-inin manuel kurulumuyla ilgiliydi, ama biz bunu bir sonraki başlıkta nasıl otomatize edeceğimizi göreceğiz.</li>
            <li>Şimdi fonksiyonumuzu test edelim. Herhangi bir excel hücresine =Topla(3;5) yazalım ve sonucu görelim. </li>
            <li>Bu arada fonksiyolarımızın hepsini tek seferde görmek için <span style="text-decoration: underline"><strong>fx</strong></span> butonuna tıklayıp kategorilerin en altına yerleşmiş olan sınıfımızı seçelim.<br />
                <img src="../../images/vsto_udf7.jpg" style="width: 418px; height: 367px" /><br />
                Burada görünmesini istemediğimiz bazı metodlar da var. Bunları aşağıdaki kodla override edelim ve Excel&#39;den gizleyelim.(Bir tek GetType override edilemez, o da nazar boncuğu kalsın). Bu arada bunlar .Net dünyasında her sınıf/nesne için varolan object metodlarıdır, çünkü .Net&#39;te her sınıf Object sınıfını inherit eder. Bizim sınıfımız MyFunctions da bunlara dahildir.<br />
                <br />
                <pre class="brush:csharp">
[ComVisible(false)]
public override string ToString()
{
    return base.ToString();
}

[ComVisible(false)]
public override bool Equals(object obj)
{
    return base.Equals(obj);
}

[ComVisible(false)]
public override int GetHashCode()
{
    return base.GetHashCode();
}
</pre>
                <p>Tekrar build edip Excele baktığımızda GetType dışındakilerin kaybolduğunu görebiliriz.</p>
                <img alt="" src="../../images/vsto_udf8.jpg" style="width: 416px; height: 368px" /></li>
        </ul>

        <h4>Formül description ve Intellisense</h4>
        <p>
            Herşey güzel de, Excel&#39;in built-in fonksiyonları veya VBA&#39;de yazdığımız UDF&#39;lerde olduğu gibi intellisense çıkmadığını farketmişsinizdir. (Burdan okuyarak farkedemezsiniz bunu tabi, kendiniz denediğinde görebilirsiniz. Bir hücreye &quot;=To&quot; yazdığınızda aşağıda açılan bir kutuda &quot;TO&quot; ile başlayan tüm fonksiyonların gelmesini beklerisiniz ama &quot;Topla&quot; gelmez.) </p>
        <p>
            Bu çok hoş bi durum değil. Bu haliyle yaptığımız çalışma VBA&#39;in bile sağladığı bir fonksiyonaliteyi sağlayamıyor durumda. Bu sizin için veya kullanıcılarınız için çok kritik değilse böyle devam edebilirsiniz. </p>
        <p>
            Ama tek sorunumuz bu değil, bir diğeri de automation add-in UDF&#39;lerimize description ve parametre açıklaması da yazamıyoruz. Halbuki bunu da VBA&#39;de yapabiliyorduk.</p>
        <p>
            Bu durumdan memnun değilseniz iki alternatifiniz var. 
            Ya VSTO Addininiz içine VBA xlam addinizi de gömmek ve otomatik kurulmasını sağlamak veya XLL addin yaratmak. XLL&#39;i aşağıdaki Excel DNA başlığında göreceğiz, diğerini de az sonra, ama önce bu yarattığımız automation add-in&#39;i nasıl otomatik kurulur hale getiririz ona bir bakalım.</p>

        
        <h3>
            VSTO Add-in içinden Automation Add-in&#39;i otomatik kurma</h3>
        <p>Şimdi elimizde aşağıdaki gibi iki proje oldu. <strong>Üstteki UDF&#39;lerimizi barındırıyor ve bunu alttaki VSTO Addin projemize <span style="text-decoration: underline">referans olarak eklememiz</span> gerekiyor. </strong>İlk adımımız bu olacak.</p>
        <p>
            <img src="../../images/vsto_udf6.jpg" style="width: 323px; height: 238px" /></p>
        <p>
            İkinci ve son adımımızda ise VSTO_UDF projemizdeki <strong>ThisAddin_Startup</strong> event handlerı içinde aşağıdaki kodları yazıyoruz. Bu kadar.</p>
        <pre class="brush:csharp">
MyFunctions functionsAddinRef = null;
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    MyUDFYukle();
}

private void MyUDFYukle()
{
    functionsAddinRef = new MyFunctions();
    string NAME = functionsAddinRef.GetType().Namespace + "." + functionsAddinRef.GetType().Name;
    string GUID = functionsAddinRef.GetType().GUID.ToString().ToUpper();

    // is the add-in already loaded in Excel, but maybe disabled
    // if this is the case - try to re-enable it
    bool fFound = false;
    foreach (Excel.AddIn a in Application.AddIns)
    {
        try
        {
            if (a.CLSID.Contains(GUID))
            {
                fFound = true;
                if (!a.Installed)
                    a.Installed = true;
                break;
            }
        }
        catch { }
    }

    //if we do not see the UDF class in the list of installed addin we need to
    // add it to the collection
    if (!fFound)
    {
        // first register it
        functionsAddinRef.Register();
        // then install it
        this.Application.AddIns.Add(NAME).Installed = true; //Bunlarda Namespace.Class şeklinde eklemek yeterli
    }
}</pre>
        <p>
            Son kısımda eklediğimiz Register metodunu da MyUDFs içindeki MyFunctions classı içine aşağıdaki gibi ekliyorum.</p>
        <pre class="brush:csharp">
public void Register()
{
    RegisterFunction(typeof(MyFunctions));
}           &nbsp;</pre>
        <p>
            Şimdi Excel&#39;de manuel kurduğunuz bu add-ini kaldırın. Solution&#39;ınızı clean edip tekrar build edin, Excelinizi tekrar açın, ve kontrol edin, UDF&#39;inizin Excel addinler içine gelmiş olması lazım.</p>
        <h3>Mevcut VBA UDF&#39;lerinin kullanımı</h3>
        <p>Diyelim ki, mevcut durumda VBA&#39;de yazdığınız Excel add-in(xlam uzantılı) içinde bir sürü UDF var. Bunları tek tek .Net ile tekrar kodlamak istemiyorsunuz. Üstelik bunlara yazdığınız fonksiyon ve parametre descriptionları da var. Performans açısından da gayet yeterliler. Kullanıcılarınızın, bunları da mevcut VSTO Add-in&#39;inizin bir parçası olarak yüklemelerini sağlayabilir miyiz. Evet. </p>
        <p>Tabi bunu yaparken kullanıcılara manuel bir yükleme yaptırma sürecinden bahsetmiyorum. Zaten öyle yapmak istesek bunu normal VBA altında anlatmıştık.</p>
        <p>Şimdi ilk olarak projemizin Properties&#39;ine tıklayıp Resources sekmesine gelelim ve oraya ilgili xlam dosyamızı resource olarak ekleyelim, ki bu da projemizin bir parçası olarak derlensin.</p>
        <p>Sonra aşağıdaki gibi bir <strong>VBA_Addin_Yukle</strong> fonksiyonu yazıp, ThisAddin_Startup içine şu satırı ekliyoruz. <strong>VBA_Addin_Yukle("VBAAddinForVstoUdf.xlam", Properties.Resources.VBAAddinForVstoUdf)</strong>. Burda IsDirectoryWritable adında yardımcı bir fonksiyondan da yararlanıyoruz(İstersek bunu Utility paketimizin içine de alabiliriz).</p>
        <pre class="brush:csharp">
private void VBA_Addin_Yukle(string vbaAddin, byte[] res)
{
    //ilk kurulduğunda var mı diye baksın, varsa işaretli mi yani installed mu diye de baksın
    try
    {
        bool isExist = false;
        foreach (Excel.AddIn a in Application.AddIns)
        {
            if (a.Name == vbaAddin) //listede varsa ve kurulu değilse kur ve çık, kuruluysa bişey yapmadan çık
            {
                if (!a.Installed)
                    a.Installed = true;
                isExist = true;
                break;
            }
        }

        if (isExist == false)
        {
            Excel.Workbook tempwb = this.Application.Workbooks.Add(); //geçici yaratıyoruz, hiç açık dosya yoksa hata alıyoruz çünkü
            string hedefdosya = "";
            if (IsDirectoryWritable(Application.UserLibraryPath)) //kullanıcının yazma izni var mı diye kontrol ediyoruz
                hedefdosya = Application.UserLibraryPath + vbaAddin;
            else 
                hedefdosya = Environment.SpecialFolder.LocalApplicationData.ToString() + vbaAddin; //buraya kesin izni vardır
                    
            File.WriteAllBytes(hedefdosya, res);
            this.Application.AddIns.Add(hedefdosya).Installed = true; //ekle ve kur tek satırda
            tempwb.Close();
        }
    }
    catch (Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(ex.Message);
    }
}
        
public bool IsDirectoryWritable(string dirPath, bool throwIfFails = false)
{
    try
    {
        using (FileStream fs = File.Create(
            Path.Combine(
                dirPath,
                Path.GetRandomFileName()
            ),
            1,
            FileOptions.DeleteOnClose)
        )
        { }
        return true;
    }
    catch
    {
        if (throwIfFails)
            throw;
        else
            return false;
    }
}           &nbsp;</pre>
        <h3>Publish etme,&nbsp; Deployment ve Uninstall işlemleri</h3>
        <h4>Publish/Deployment</h4>
        <p>Sadece VSTO Add-ini publish etmemiz yeterli. Zira zaten bunun içine UDF projemizi referans olarak veriyoruz.</p>
        <h4>Uninstall</h4>
        <p>Denetim Masasından Uninstall işlemi ile sadece VSTO_UDF&#39;i kaldırmış oluruz. MyUDFs(veya VBA add-inimiz) hala add-ins içinde görünmeye devam eder. Zira bunların organik bir bağı yoktur. Bunları da kladırmak için Excel Add-in&#39;s penceresinde tick işaretlerini kaldırmak yeterli olabileceği gibi, dosyaları da tamamen kaldırmak isterseniz bunların kurulduğu yerleri bulun ve dosyaları silin. Kullanıcılarınıza da bu bilgiyi vermeyi unutmayın.</p>
<h3>Kaynaklar</h3>
<p>Son olarak faydalandığım kaynakları belirtmek isterim:</p>
        <ul>
            <li><a href="https://docs.microsoft.com/tr-tr/archive/blogs/eric_carter/writing-user-defined-functions-for-excel-in-net">En temel kaynak</a> (ama eksik bilgiler var)</li>
            <li><a href="https://www.codeproject.com/Articles/606446/UsingplusC-plus-NETplusUserplusDefinedplusFuncti">https://www.codeproject.com/Articles/606446/UsingplusC-plus-NETplusUserplusDefinedplusFuncti</a> </li>
            <li><a href="https://adamtibi.net/07-2012/using-c-sharp-net-user-defined-functions-udf-in-excel">https://adamtibi.net/07-2012/using-c-sharp-net-user-defined-functions-udf-in-excel</a> </li>
            <li>ve son olarak <a href="https://redlevelgroup.com/excel-automation-add-in/">https://redlevelgroup.com/excel-automation-add-in/</a> Bu site dışında kimse target framework belirtmekten bahsetmemiş. Bunu yapmayıp AnyCPU bıraktığımda herhangi bir hata vermiyordu ancak add-inimi bir türlü automation listesinde görememiştim. Siz kendiniz deneyip görün.</li>
        </ul>        
    </div>


    <h2 class="baslik">Excel DNA ile XLL tabanlı UDF yazmak</h2>
    <div class="konu">
        <p>
            Normalde XLL yazmak için C/C++ bilmek gerekiyor ve bunu <a href="https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit">burada</a> bahsedilen SDK ile yazıyorsunuz. Ben açıkçası bunu hiç denemedim. Zira .Net dışında çıkmak gibi bir niyetim yok. Hele hele Excel DNA gibi, bana C/C++&#39;ın performansını vaadeden araçlar varsa. (Bi ara PyXLL&#39;i deneyeceğim ama)
        </p>
        <p>
            Peki madem Excel DNA bu kadar efsane. Neden Automation add-inle uğraşalım ki?</p>
        <ul>
            <li>Yeni bir paket(kütüphane) yüklemeyip Solution hacmini artırmak istemiyorsunuzdur</li>
            <li>Hatta yeni bir paket nasıl kullanılır, öğrenmek istemiyorsunuzdur</li>
            <li>Hız ve performans sizin için o kadar da kritik dğeildir, zaten küçük fonksiyonlar yazıyorsunuzdur</li>
            <li>Description ve Intellisense konusunu dert etmiyorsunuzdur. Hatta fonksiyonları sadece siz kullanıoyrsunuzdur, o yüzden ne olduklarını zaten biliyorsunuzdur</li>
            <li>Diğer nedenler(<a href="https://stackoverflow.com/questions/26974959/pros-and-cons-of-vsto-vs-excel-dna">Şurda </a>da VSTO ve ExcelDNA karşılaştırması var, bi bakın isterseniz)</li>
        </ul>

        <p>
            Peki, biz yukardaki maddelerden birinin bizi mutlu etmediğni düşündük ve Excel DNA ile çalışmaya karar verdik diyelim ve devam edelim.&nbsp;</p>
        <p>
            Bu arada şunu da tekrar belirtmekte fayda var.
            ExcelDNA ile Ribbon arayüzü de geliştirebilir, Taskpane de yaratabilir, hatta VBA&#39;de kullanabileceğiniz classlar da yazabilirsiniz. Ancak biz Ribbon işini normal <strong>Interop </strong>ile yapmıştık ve bu bizim için gayet de yeterli diyoruz, VBA&#39;de kullanacağımız classlar yazmaya gerek yok diyoruz. Özetle biz burada sadece UDF kısmına odaklanalaım, ki zaten gerçekten ExcelDNA daha çok bu amaçla kullanılmakta. Arzu eden tüm geliştirmesini Interop yerine Excel DNA ile de yapabilir.</p>
        <p>
            <a href="http://www.excel-dna.net/">ExcelDNA&#39;in</a> bir dokümantasyon <a href="https://docs.excel-dna.net/">sitesi</a> ve bir de <a href="https://github.com/Excel-DNA/ExcelDna">github</a> reposu var. Okumaya <a href="https://docs.excel-dna.net/what-and-why-an-introduction-to-net-and-excel-dna/">şu sayfa</a> ile devam etmenizi, sonrar tekrar buraya gelmenizi tavsiye ederim.</p>
        <h3>
            Adım adım XLL add-in oluşturma</h3>
        <p>
            Komple repository&#39;yi daha önce indirmediyseniz burdaki örnek uygulamayı <a href="https://github.com/VolkiTheDreamer/excel/tree/main/UDF_XDNA">şuradan</a> indirebilirsiniz.&nbsp;</p>
        <ul>
            <li>Yeni bis class library oluşturun, ben adına UDF_XDNA dedim. Class1&#39;in adını da MyDNA olarak değiştirdim.</li>
            <li>Nuget manager&#39;ı açın ve Excel DNA yazın. Çıkan listede hem <strong>ExcelDna.Addin</strong> olanı hem de <strong>ExcelDna.Intellisense</strong> olanı seçin. Add-in paketini kurunca <strong>integration</strong> d<span style="font-weight: bold">a </span>kurulacaktır. 32-64 bitle ilgili sıkıntılar olabilir. Sorun yaşarsanız github&#39;daki issues bölümünden bakabilirsiniz, ancak öncelikle <a href="https://medium.com/@efrem.sternbach/excel-dna-or-why-are-you-still-using-vba-a76f565884ff">şuraya</a> bakın derim,&nbsp; orda bu konuya değinilmiş. <strong>Integration </strong>referansı üzerindekyken alttaki Properties gelin, ve <strong>Copy Local = False </strong>ayarlamasını yapın. Zira, bu dosyanın output klasöründe olmasını istemiyoruz, çünkü bu zaten oluşacak xll dosyası içinde gömülü olacak.<br />
                <br />
                <img alt="" src="../../images/vsto_xll3.jpg" style="width: 360px; height: 292px" /></li>
            <li>Add-in&#39;i kurunca &quot;<em><strong>Proje adı</strong></em><strong>-AddIn.dna</strong>" isminde bir dosya da oluşturulacaktır. Bu dosyanını içeriği de şöyle bişeydir:<br />
                <br />
            <pre class="brush:csharp">
&lt;?xml version="1.0" encoding="utf-8"?>
&lt;DnaLibrary Name="UDF_XDNA Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2018/05/dnalibrary">
  &lt;ExternalLibrary Path="UDF_XDNA.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" />
&lt;/DnaLibrary></pre>
            </li>
                <li>Bu dna uzantılı dosyanın Properties&#39;ten <strong>Copy to Output Directory</strong> özelliğine <strong>Do not copy</strong> değerini atayın.<li>Şimdi aşağıdaki gibi basit bir fonksiyon yazın</li>. Sınıfımızın ve metodlarımızın <strong>static </strong>olduğuna dikkat edin.<br />
&nbsp;<pre class="bursh:csharp">
using ExcelDna.Integration;

public static class MyDNA
{
    [ExcelFunction(Description = "ilk basit fonksiyonum", Category = "XLL Functions")]
    public static string MerhabaXLL()
    {
        return "Merhaba XLL dünyası";
    }
}</pre>
               </ul>
        <p>
            ExcelFunction isimli attribute ile fonksiyonumuza açıklama yazmış oluyoruz.</p>
        <p>
            Bu haliyle projemizi build edelim veya debug edip de bakabiliriz. Bakalım fonksiyonumuz uygun yerde görünüyor mu, evet fonksiyon kategorilerinin en altında <strong>XLL Functions</strong> içinde çıktı.</p>
        <p>
            <img src="../../images/vsto_xll1.jpg" style="width: 417px; height: 363px" /></p>
        <p>
            &nbsp;Excel&#39;e yazarken intellisense de çıkıyor. Fonksiyon description alanı da geliyor.</p>
        <p>
            <img src="../../images/vsto_xll2.jpg" style="width: 438px; height: 219px" /></p>
        <p>
            Herşey yolunda gibi. Şimdi daha karışık bir fonksiyon yazalım ve bu sefer hem parametre alsın ve bu paremetrelerin de açıklamalarını yazalım. Öyle ya, Automation add-inde olmayan tüm bu özellikleri&nbsp; görmemiz laızm.</p>
        <p>
            Şimdi şu fonksiyonu yazalım. Burda dikkat edilecek iki husus var:</p>
        <ul>
            <li>Parametre olarak ilk önce ExcelArgument attribute yazdık ve argüman açıklamalarını yazdık.</li>
            <li>ExcelDNA&#39;in kendi <a href="https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/ArgumentTypeBasics">sitesinde</a> belirttiği üzere parametre işleri biraz alengirli. Öncelikle <strong>Range tipinde bir paramerte yazamıyoruz</strong>. Bunun yerine <strong>object tipli iki boyutlu bi dizi</strong> yazıyoruz.
                Ki VBA&#39;den hatırlarsanız her range, tek kolon/satır olsa bile 2D bir dizidir. </li>
            <li>Keza, her .Net değişken tipi de desteklenmiyor. Mesela ayraç parametresini <strong>char </strong>olarak veremedik, bunu da object verip sonra char&#39;a çevirdik. Burda ayraçı optional verdik, ama default değer atamadık, zira object tiplere null dışında default değer atanamıyor, o yüzden kod içinde buna parametre geçirilip geçirilmediğini kendim kontrol ettim. Optional parametrelerle iligli genel bilgiler için aşağıdaki kaynaklara bakın.</li>
        </ul>
        <pre class="brush:csharp">
[ExcelFunction(Description = "Bir metinde kaç kelime olduğunu sayar", Category = "XLL Functions")]
public static int KacKelimeXLL(
    [ExcelArgument(Name = "rng",Description = "Kelime sayısı yazdırılacak olan metin")] object[,] rng, 
    [ExcelArgument(Name = "ayrac", Description = "Hangi ayraçla bölünecek, default olarak boşluktur")] [Optional] object ayrac
    )
{
    char ayrac2;
    if (ayrac is ExcelMissing)
        ayrac2 = ' ';
    else
        ayrac2 = System.Convert.ToChar(ayrac);
    string icerik = rng[0, 0].ToString();
    return icerik.Split(ayrac2).Length; 
}</pre>
        <p>
            Nasıl göründüğüne bakalım.</p>
        <p>
            <img src="../../images/vsto_xll4.jpg" style="width: 646px; height: 423px" /></p>
        <p>
            Yukardaki resimlerden göreceğiniz üzere fonksiyonumuzunz bu sefer parametre açıklamları da geldi. Faket Excel&#39;in yerel fonksiyonlarında(Ör:SUM) olan fonksiyon tooltip&#39;i(bu da intellsisensin bir parçası olarak düşünülüyor) bizde gelmedi. Bunu sağlamak için  <strong>COMIntegration.cs</strong> adında bi dosya yaryatıp içine şunları yazalım.</p>
        <pre class="brush:csharp">
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices;

namespace UDF_XDNA
{
    [ComVisible(false)]
    internal class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
</pre>
        <p>
            Tekrar build edelim. Ve sonuç:</p>
        <p>
            <img alt="" src="../../images/vsto_xll5.JPG" /></p>
        <p>
            Gördüğünüz gibi Excel&#39;in yerel fonkisyonlarında bile olmayan bir özellik geldi. Artık fx tuşuna basıp fonksiyon argümanı okumaya da gerek yok, fonksiyonumu seçtiğimizde o an kaçıncı parametredeysek onun açıklamasını görüyoruz. Tıpkı Visual Studio içinde kodlama yaparken bir fonksiyona parametre gönderdiğimiz sırada oluşan görüntü gibi. </p>
        <p>
            Bu fonksiyon tanımları ve argüman açıklamaları için daha sistematik bir çalışma yapmak isterserniz <a href="https://stackoverflow.com/questions/8172591/doing-documentation-using-excel-dna">şurada</a> ve <a href="https://docs.excel-dna.net/creating-a-help-file/">şurada</a> bashedilen yönteme bakabilirsiniz.</p>
        <p>
            Başka bir örnek de yine ExcelDNA&#39;nın kendi sitesinden alıp biraz modifiye ettiğim bir örnek. Burda seçilen range&#39;deki tüm sayıların toplamını aldıran bir metod yazdık.</p>
        <pre class="brush:csharp">
public static double dnaSumEvenNumbers2D(object[,] arg)
{
    double sum=0;
    int rows;
    int cols;

    rows = arg.GetLength(0);
    cols = arg.GetLength(1);

    for (int i = 0; i <= rows - 1; i++)
    {
        for (int j = 0; j <= cols - 1; j++)
        {
            object val = arg[i, j];
            if (!(val is ExcelEmpty) && (double)val % 2 == 0) //boş olup olmadığını da kontrol etmekte fayda var, yoksa hata alırız 
                sum += (double)val;
        }
    }

    return sum;
}            &nbsp;</pre>
        <h4>
            Range İşlemleri</h4>
        <p>
            İşimiz hala bitmedi, Range nesnesinin daha alengirli kısımları da var. Yukarıda range nesnesini object[,] olarak geçtik ama onun değerini kullanmış olduk. Peki ya onu bir range gibi kullanmak isteseydik? İşte şimdi işler&nbsp;biraz daha değişiyor. Zira şimdi Excel&#39;e bunun bir range referansı olduğu söylememiz gerekiyor. Bunun için de üç yeni kavram hayatımıza girer.</p>
        <ul>
            <li>Metod attribute&#39;ü olarak ekleyeceğimiz <strong>[ExcelFunction(IsMacroType = true)]</strong></li>
            <li>Parametre attribute&#39;ü olarak ekleyeceğimiz <strong>[ExcelArgument(AllowReference=true)]</strong></li>
            <li>ve ilgili range parametresini, ki bunu object olarak vereceğiz,<strong> ExcelReference</strong> sınıfı ile casting&#39;e tabi tutmak.</li>
        </ul>
        <p>
            Örnek bir kod yazalım. Bu kod&#39;da parametrimize gerçekten range olarak ihtiyacımız var, onun değerilye ilgilenmiyoruz. MEsela hüçcre içi renk bilkgisi ile ilgilenelim. Kodumuz şöçyle.</p>
        <pre class="brush:csharp">
[ExcelFunction(IsMacroType = true)]
public static double GetArkarenk([ExcelArgument(AllowReference=true)] object hucre)
{
    ExcelReference rng = (ExcelReference)hucre;
    Excel.Range refrng = ReferenceToRange(rng);
    return refrng.Interior.Color;
}      

//yardımcı fonkisyon
private static Excel.Range ReferenceToRange(ExcelReference xlRef)
{    
    Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;//Application nesnesine erişimi böyle sağlarız                
    //dynamic app = ExcelDnaUtil.Application; //Interop referansı eklemeden böyle de yapabilrdik ama intellinseten yararlanamayız
    string strAddress = XlCall.Excel(XlCall.xlfReftext, xlRef, true).ToString();
    return app.Range[strAddress];
}</pre>
        <p>
            Hucre parametresini object olarak geçtiğimize dikkat edin, object[,] değil. Daha sonra bunu bir ExcelReference nesnesine döndürüyoruz, bunun için casting yapıyoruz. Sonra bu cast edilmiş nesneyi de ReferenceToRange yardımcı fonksiyonuna göndererek gerçek bir Excel.Range elde ediyorum. Burda tabiki <strong>Excel Interop&#39;u referans olarak eklemeyeli </strong>unutmayalım.</p>
        <p>
            Bir başka örnek de seçili bölgedeki hücre sayısını getirsin.</p>
        <pre class="brush:csharp">
[ExcelFunction(IsMacroType = true)]
public static long HucreAdet([ExcelArgument(AllowReference = true)] object alan)
{   
    ExcelReference rng = (ExcelReference)alan;
    Excel.Range refrng = ReferenceToRange(rng);
    return refrng.Cells.Count;
}          &nbsp;</pre>
        <p>
            Bu konuda <a href="https://docs.excel-dna.net/excel-c-api/">şurada</a> biraz daha detay bulabilirsiniz.</p>
                <h3>
            Projeyi yayınlama</h3>
        <p>
            Kullanıcılarımnızn bu UDF&#39;imizi solo kullanmlaarını istiyorsak işimiz basit. Release modda build ettiğimizde herşey hazır oluyor. <strong>\bin\Release</strong> klasöründe UDF_XDNA-AddIn-packed.xll ve bunun 64 bit versiyonu oluşacaktır. Bunlardan uygun olanı kullancınıza vermeniz yeterli, bunu normal bir Excel add-in gibi manuel ekleyebilirler.</p>
        <p>
            Peki ya projemizi VSTO add-inimizle birlikte paketlemek istiyorsak. Yani yukarda Automation add-in ve Excel add-in&#39;lerle yaptığımızın aynısını bu XLL addinle de yapmak istiyorsak? 
            Aslında çok basit;&nbsp; yukarda Excel-VBA ADd-inimizi yüklerken yaptığımızın aynısını yapacağız. İlgili <strong>Packed-xll</strong>(bizim örneğimizde UDF_XDNA-AddIn-packed.xll veya UDF_XDNA-AddIn64-packed.xll) dosyasını bir resource olarak ana VSTO uygulamamıza ekleyip startup sırasında bunun kurulumunu sağlarız. Bunu tekrar buraya yazmıyorum. Yukardaki kodların aynısı olacak.</p>
        <p>
            Diğer kurulum seçenkeleri için <a href="https://docs.excel-dna.net/installing-your-add-in/">şuraya</a> bakınız.</p>

        <h3>
            Kaynaklar</h3>
        <p>
            ExcelDNA,
            3rd party bir kütüphane olmasına rağmen fena olmayan bir kaynağa sahip. Kendi siteleri 3 farklı yerde bulunduığu için çok karışık geliyor bana. Ben elimden geldiğimce derleyip özetlemeye çalıştım, en faydalı kısımların vermeye çalıştım ama daha derinleşmek için aşağıdaki kaynaklara bakabilirisniz.</p>
        <ul>
            <li>UDF paremetreleri için: <a href="https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/ArgumentTypeBasics">https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/ArgumentTypeBasics</a> </li>
            <li>Excel Range işlemlerinin anlatılıdığ yer: <a href="https://docs.excel-dna.net/excel-programming-interfaces/">https://docs.excel-dna.net/excel-programming-interfaces/</a> ve <a href="https://docs.excel-dna.net/excel-c-api/">https://docs.excel-dna.net/excel-c-api/</a> </li>
            <li>Optional parametler: Genel .Net için <a href="https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/named-and-optional-arguments">buraya</a> ve <a href="https://www.c-sharpcorner.com/UploadFile/75a48f/optional-parameter-in-C-Sharp/">buraya</a>, ExceDNA özelinde bilgi almak için <a href="https://docs.excel-dna.net/optional-parameters-and-default-values/">buraya</a> bakın.</li>
            <li>Kapsamlı bir doküman: <a href="http://www.sysmod.com/vba-to-vb.net-xll-add-in-with-excel-dna.pdf">http://www.sysmod.com/vba-to-vb.net-xll-add-in-with-excel-dna.pdf</a> </li>
            <li>
<a href="https://stackoverflow.com/questions/957575/how-to-easily-create-an-excel-udf-with-vsto-add-in-project">https://stackoverflow.com/questions/957575/how-to-easily-create-an-excel-udf-with-vsto-add-in-project</a></li>
            <li>
<a href="https://smurfonspreadsheets.wordpress.com/2010/02/18/xlls-with-exceldna/">https://smurfonspreadsheets.wordpress.com/2010/02/18/xlls-with-exceldna/</a></li>
            <li>
<a href="https://adamtibi.net/07-2012/using-c-sharp-net-user-defined-functions-udf-in-excel">https://adamtibi.net/07-2012/using-c-sharp-net-user-defined-functions-udf-in-excel</a></li>
        </ul>

    </div>

    <h2 class="baslik">Ekstra Performanslı UDF yazımı</h2>
    <div class="konu">
        <p>Çok yakında...</p>
        </div>
</asp:Content>
