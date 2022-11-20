<%@ Page Title='.Net Dilleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'></div>
    <h1>.Net Dilleri: VB.Net ve C#</h1>
    <p>VSTO'yu ilk öğrendiğimde VB.Net yeterli olur diye düşünüyordum, özellikle c#'ın(si 
	şarp diye okunur) bazı eksiklikleri can sıkıcıydı ancak bunlar yıllar içinde giderildikten ve c#'a da iyice ısındıktan sonra c#'ın tercih edilmesi gerektiğini düşünüyorum.</p>
    <p>Öncelikle belirtmek isterim ki, burada bu iki dili detaylıca görmeyeceğiz. Sadece VSTO programlama yaparken işimize yarayacak kısımları bilmemiz yeterli. O yüzden ilerleyen zamanlarda bu dillerle ilgili internette bolca bulunan diğer eğitim materyallerinden faydalanmanızı tavsiye ederim.</p>
    <p>VBA'yi bildiğiniz için VB.Net size biraz daha kolay gelecektir. Ben burada VBA ve VB.Net arasındaki farkları da anlatmaya çalışacağım. Böylece yoldaki kazalardan korunmuş olacaksınız.</p>
    <p>Bununla beraber ilerleyen sayfalardaki kod örneklerinin büyük çoğunluğu c# ile yazacağım. Keza <a href="OrnekProjeler_Konular.aspx">Örnek p<span style="">rojeler</span></a> de c# ile yazdığım projeler olacak. Ancak bunlardan Excelent&#39;ı Vb.Net ile yazdım, ki içlerinde en büyük proje de budur.</p>
    <p><strong>Size önerim VSTO programı yaratmadan önce sıradan VB.Net ve C# programları yaratarak başlayın ve bu dillere aşina olun, ondan sonra VSTO uygulamalarına doğru geçiş yapın.</strong> </p>
    <h2 class='baslik'>VB.Net</h2>
    <div class='konu'>
    <p>Öncelikle VB.Net&#39;ten başlıyoruz. Bununla beraber anlatacaklarımın çoğu, 
	syntaxtaki farklılıklar dışında c# için de geçerlidir.</p>
        <p>VB.Net&#39;i VBA&#39;den en büyük ayıran kısmı syntaxıdır sanırım. En aşağıda detaylı bir karşılatırma bulacaksınız. O yüzden bu kısımda farklara detaylı girmiyorum.</p>
        <p>VB.Net, .Net platformunun ana prensipleri ışığında çalışır ve tam anlamıyla 
		nesne yönelimli(obje oriented) bir dildir. VBA ise VB6 dilinden başka birşey değildir, bunun Ofis programları içine gömülmüş halidir, en büyük eksiği tek başına executable programları çalıştıramamasıdır. VB6 tam anlamıyla bir object oriented dil olmadığı için Microsoft, .Net platformu bile birlikte VB6nın yerine de VB.Net&#39;i getirdi. VB.Net&#39;te tüm diğer object yönelimli dillerdeki gibi 4 ana unsur vardır: Encapsulation, Abstraction, Inheritance ve Polymorphism. </p>
        <p>Bu arada VBA&#39;deki modül kavramı hala var ancak sınıf kavramı da bulunmaktadır. (c#&#39;ta ise modül diye birşey yok, sadece sınıf var)</p>
        <p>VB.net&#39;le ilgili özet bir eğitim sitesine <a href="https://www.tutorialspoint.com/vb.net/index.htm">buradan</a> ulaşabilirsiniz. Benim sitemdeki bilgilerden sorna yetersiz kaldığınız hissederseniz bu siteye bakabilirsiniz. Kendinizi VB.Net konusunda çok da geliştirmenize gerek yok, enerjinizi daha çok c# öğrenmeye harcayın derim ben, 
		ki aynı sitede ve daha birçok sitede c# için de öğrenim kaynağı 
		bulacaksınızdır. Eminim bu dil size başka kapılar da açacaktır. (Java 
		veya c++ öğrenme, mobil development yapma gibi)</p>
        <h3>.Net Programlarının Genel Yapısı</h3>
        <h4>Modül(sadece vb.net) ve sınıf</h4>
        <p>VB.Net programlarında VBA&#39;deki modüller gibi modüller bulunabileceği gibi, .Net dünyasına özgü namesapace ve onların altında sınıflar 
		da bulunabilir. </p>
        <pre class="brush:vbnet">
Imports System &#39;kullanmak istediğim sınıfları içeren namespaceleri programa dahil ediyoruz

Module Module1
   'This program will display Hello World 
   Sub Main()
      Console.WriteLine("Hello World") &#39;Console denen terminal ortamına çıktımız yazdırılır. VSTO&#39;da bunu kullanmayacağız.
      Console.ReadKey()
   End Sub
End Module</pre>
        <p>
            veya modül olmadan şölye;</p>
        <pre class="brush:vbnet">
Imports System &#39;kullanmak istediğim sınıfları içeren namespaceleri programa dahil ediyoruz

Class Personel
   &#39;Sınıf seviyesinde değişkenler
    Private Name As String
    Private Sicil As Integer 

    Private Sub BilgiVer()
        Dim degisken As String
        Console.WriteLine(....)
    End Sub

End Class

Class Test
   Sub Main()
      Dim p As New Personel() 'Yukarıda tanımladığımız sınıftan bi nesne yaratıyoruz
      r.Name="Volkan"
      r.Sicil=123
      r.BilgiVer()
      Console.ReadLine()
   End Sub

End Class        </pre>
        <p>
            Modül ve classlar arasındaki farkı şu <a href="https://stackoverflow.com/questions/881570/classes-vs-modules-in-vb-net">sayfada</a> bulabilirsiniz. 
			Sınıfların da erişim türleri olmaktadır. Bunlardan Shared sınıflar(c#'ta 
			static sınıf olarak geçerler) Modüllere çok benzerler. Bundan şimdi 
			bahsediyorum, çünkü özellikle c# tarafında bol bol statik sınıf 
			kullanacağız. Modüllerin ana kullanım prensibini hatırlayack 
			olursak, c#'taki bu kullanım şeklini de çok daha iyi anlayabiliriz.</p>
        <h4>Sınıflar ve nesneler</h4>
        <p>.Net&#39;te kodlarımız hiyerarşik bir yapı içindedir. En üstte 
		<strong>namespace </strong>denen yapılar vardır, bunları özet olarak 
		<strong>kütüphane </strong>terimiyle özdeş düşünebilirsiniz(tam olarak öğle değil ama basit olması adına şimdilik öyle düşünün). Bu namespacelerin/kütüphanelerin 
		altına başka namespaceler olabilir, bunların da içinde <strong>sınıflar </strong>vardır. Sınıflar, başka sınıfları miras alabilir ve hiyerarşik bir yapı oluşturur. Ör:System namespace'i 
		altında Windows isimli bir namespace ve onun altında da Forms isimli bir sınıf bulunur ve gösterimi System.Windows.Forms şeklindedir.</p>
        <p>Object Oriented bir dilde bu sınıflardan objeler yaratılır. Sınıflar bu nesneler için kalıp görevi görürler. Bu nesneleri yaratabilmek için bu sınıfları programımıza dahil etmeliyiz. Bunun için&nbsp;öncelikle ilgili 
		sınıfın bulunduğu kütüphaneyi Referanslar içine eklememiz gerekir. Referanslara eklemek onu 
		kodumuzda kullanabiliriz anlamına gelir.
		<strong>Imports</strong> deyimi ile programın başına yazarak da bu sınıfları uzun uzun yazmaktan kurtulmuş oluruz. Bu arada bazı kütüphaneler default olarak başta hazır eklenmiş olarak gelir.</p>
        <pre class="brush:vbnet">&#39;Hiç import yapılmazsa
System.Diagnostics.Debug.Write(&quot;Merhaba&quot;)

&#39;Tek seviye import
Imports System
Diagnostics.Debug.Write(&quot;Merhaba&quot;)

&#39;İki seviye import
Imports System.Diagnostics
Debug.Write(&quot;Merhaba&quot;)</pre>
        <p>Biz de kodlarımızda Interop altındaki Excel namespace&#39;ini programımıza dahil edeceğiz.</p>
        <pre class="brush:vbnet">Imports Microsoft.Office.Interop.Excel</pre>
        <p>
            Yaygın import şekli ise Excel&#39;i alias olarak almaktır. Böylece Application gibi birçok namespace&#39;te olabilecek sınıfların önüne bu aliası yazarak karışıklıkları da önlemiş oluruz.</p>
        <pre class="brush:vbnet">Imports Excel = Microsoft.Office.Interop.Excel
Dim hucre As Excel.Range</pre>
        <h3>Veri Tipleri ve Değişken tanımlama/atama</h3>
        <p>Veri türleri VBA&#39;den oldukça farklı olup <a href="https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/">buradan</a> detaylara ulaşabilirsinz. </p>
        <p>Değişken atama şekli de VBA&#39;den biraz farklıdır. Şimdi şu örneklere bakalım:</p>
        <p>Dim adet As Integer = 10 &#39;tek satırda tanımlama ve atama. VBA'de 
		sadece constantlar böyle atanabilyor<br />
            Dim sayi1, sayi2 as Integer 'VBA'de ilki Variant olurdu, VB.Net'te 
		ikisi de Integer</p>
        <pre class="brush:vbnet">
Module DataTypes
   Sub Main()
      Dim n As Integer
      Dim da As Date
      Dim bl As Boolean = True
      n = 1234567
      da = Today
      
      Console.WriteLine(bl)
      Console.WriteLine(CSByte(bl))
      Console.WriteLine(CStr(bl))
      Console.WriteLine(CStr(da))
      Console.WriteLine(CChar(CChar(CStr(n))))
      Console.WriteLine(CChar(CStr(da)))
      Console.ReadKey()
   End Sub
End Module
</pre>
        <p>
            Değişkenin tanımlandığı yere göre ömrü de farklı olacaktır. Sınıf 
			seviyesindeki değişkenler, o sınıftan yaratılan tüm nesnelerde var 
			olacaktır. Local değişkenler sadece ilgili scope'ta var olacaktır.</p>
		<p>
            c#'ta değişken tanımlama biraz daha farklı ve basittir. Önce 
			değişken tipini sonra adını yazarız.</p>
		<pre class="brush:csharp">int i = 0;
string isim = "";</pre>
		<h4>Veri tipi dönüştürme</h4>
        <p>
            Veriler arasında dönüşüm yapmak için VBA&#39;dekine benzer dönüşüm fonksiyonları vardır. Detaylar <a href="https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/type-conversion-functions">burada</a>. Bunun dışında bir de 
			<strong>Convert</strong> sınıfı var. 
			Bir de tam olarak dönüştürme olmayan ama benzer kavramlar olan
			<strong>Casting, Parsing </strong>ve <strong>Boxing</strong> var.</p>
        <ul>
			<li><strong>converting</strong>: bir veri tipini başkasına çevirme. 
			Ör: Integer&#39;ı Floata. int i = Convert.ToInteger(sayi)</li>
			<li><strong>casting</strong>: compilera &quot;bu obje aslında başka birşey, onu bu şeye dönüştür&quot; diyoruz</li>
			<li>
			<p>
            <strong>parsing</strong>: Bir stringi sayı olarak yorumla(genelde böyle), veya daha genel anlamda bir string içinden anlamlı bişey çıkarma amaçlı kullanılır</p>
			</li>
			<li>
			<p>
            <strong>boxing</strong>: değişkenleri object tipine çevirmedir. 
			Tersi Unboxingdir, castinge benzer.</p>
        	</li>
		</ul>
        <h5>
            CType ve DirecCast fonksiyonları(Vb.Net 
			kodlarıdır, c# eşlenikleri de mevcuttur)</h5>
        <p>
            Dönüş tipi Obje olan nesnelerin üyelerine erişmek için onları daha 
			spesifik nesnelere dönüştürmemiz gerekmektedir. Ve nitekim Excel 
			objelerinin de bir kısmının dönüş tipi object'dir. O yüzden sık sık 
			DirectCast/CType fonksiyonları ile dönüşüm yapacağız. c#'ta bu işlem 
			biraz daha farklı bir şekilde yapılmaktadır.</p>
        <pre class="brush:vbnet">CType(obj, sınıf).sınıfınmetodu() 'metod kullanımı
degisken=CType(obj, sınıf).sınıfınpropertysi 'property kullanımı

'veya öncelikle bir nesne değişkenine atarız sonra üyeleri kullanırız
myobj = CType(obj, sınıf)
degisken= myobj.somevalue

&#39;Excel dünyasından bir örnek
Svar= CType(sender, Button).Text
hucre= CType(app.ActiveCell, Excel.Range) 'app de Application nesnesini temsil eden başka bir değişken</pre>
		<p>c#&#39;ta casting, dönüştürlecek sınıf tipi parantez içinde ilgili objenin 
		önüne yazılarak yapılır.</p>
		<pre class="brush:csharp">hucre=(Excel.Range)app.ActiveCell;</pre>
        <h4>Nesne yaratma</h4>
        <p>
            VBA&#39;deki gibidir. Late ve Early binding olmak üzere iki yolu vardır.</p>
        <pre class="brush:vbnet">Dim nesne As New Sınıf() &#39;sondaki paranteze dikkat
&#39;veya iki yarı satırda
Dim nesne As Sınıf()
nesne = New Sınıf() #set yok</pre>
        <h3>Erişim Tipleri</h3>
        <p>Erişim tipleri hem değişkenler hem prosedürler için kullanılabilir. Public, Private, Protected, Friend(c#'ta 
		Internal), Protected Friend</p>
		<p>Bunları siz araştırın lütfen.</p>
        <h3>Hata Yönetimi ve Debugging</h3>
		<p>VBA'deki yetersiz On Error ifadeleri yerine daha güçlü olan TryCatch 
		ve TryCatchFinally blokları kullanılır. </p>
		<ul>
			<li>Try ile Catch arasına yazılan kodlar icra edilir.</li>
			<li>Hata alınırsa Catch bloğuna gelir</li>
			<li>Hata alsa da almasa da bir kod çalışsın istersek de bunu Finally 
			bloğu içine yazarız.</li>
			<li>Hata alındığında tüm koddan çıkılmaması adına iç içe TryCatch 
			blokları kullanmak yaygın bir durumdur</li>
		</ul>
        <pre class="brush:vb">Try
   [ tryStatements ]
   [ Exit Try ]
[ Catch [ exception [ As type ] ] [ When expression ]
   [ catchStatements ]
   [ Exit Try ] ]
[ Catch ... ]
[ Finally
   [ finallyStatements ] ]
End Try</pre>
        <p>Debugginle ilgili olarak tuşlarda bazı değişiklikler var. Ör:Adım adım 
		ilerlemek için F8 değil F11 kullanıyoruz. Hepsini menüden 
		görebilirsiniz.</p>
        <h3>Eventler</h3>
		<p>Eventler, VBA'deki ile hemen hemen aynı. Sadece Event handler'ların 
		sonunda "Handles" vardır. Bunu takiben birden fazla control için event 
		yazılabilr. Ör: Hem A hem B butonlarına tıklandığında aynı eventin 
		çalışması sağlanabilir. Hangi butona tıklandığını da sender parametresi 
		ile elde edebiliriz.</p>
		<pre class="brush:vb">
Private Sub btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click, btn2.Click

End Sub&nbsp;</pre>
        <h3>Dizimsiler(Veri Yapıları)</h3>
        <pre class="brush:vbnet">Dim dizi(10) As Integer &#39;Son indeksi 10 olan 11 elemanlık bir dizi tanımlandı, ancak henüz yaratılmadı??
Dim dizi() As Integer = New Integer(){} &#39;Boş dizi
Dim dizi() Ad Integer = {1,2,3,4,5} &#39;5 elemanlı bir dizi
Dim dizi() As Integer = New Integer() {1,2,3,4,5} &#39;5 elemanlı bir dizi</pre>
        <p>
            VB.Net'te, <strong>Option Base 1</strong> olayı kaldırıldı, yani indeks her daim 0&#39;dan başlar.</p>
        <p>Redim, içiçe diziler ve çokboyutlu diziler aynen var. Ancak Redim aynı zamanda boyutu baştan belirtilmiş dizilerin boyutunu yeniden boyutlandırmak için de kullanılabilir, bu VBA'de yoktu.</p>
        <p>Ubound&#39;a ek olarak <strong>Length</strong> propertysi var, Lbound yok, zaten 
		alt indeks hep 0&#39;dır. Ayrıca <strong>IndexOf </strong>gibi faydalı propertyler ile <strong>Copy, Reverse, Sort </strong>gibi VBA&#39;de olmayan faydalı metodları var.</p>
		<p><a href="../Fasulye/NeNeredeNasil_Diziler.aspx">
		../Fasulye/NeNeredeNasil_Diziler.aspx</a> sayfasından detaylı kullanım 
		örneklerini görebilirsiniz.</p>
        <h5>Diğer veri yapıları</h5>
        <p><strong>System.Collection</strong> ve <strong>System.Collection.Generic </strong>namespaceleri içinde 
		çok sayıda veri yapısı bulunmaktadır. Bunlar oldukça güçlü veri yapıları 
		olup sıklıkla <strong>List</strong> ve <strong>Dictionary</strong>'leri 
		kullanıyor olacağız. Bunlardan <strong>List </strong>VBA'Deki Collection'lara benzer 
		ancak çok daha güçlüdür.</p>
        <h3>Algoritmik yapılar</h3>
        <p>Koşullu yapılar ile döngüler hemen hemen aynıdır,&nbsp; o yüzden detaya girmiyorum. Ancak bilmeniz gereken önemli bir husus var ki, VBA&#39;de bir değişken prosedürün her yerinde yaşamaya devam ediyordu, yani kapsamı prosedür seviyesindeydi. Detayları <a href="../VBAMakro/Temeller_Birazdahaterminoloji.aspx">bu sayfadan</a> hatırlayabilirsiniz. VB.Net&#39;te ise local değişkenler, 
		tanımlandığı koşul veya döngü bloğu içinde geçerlidir. Bununla beraber, 
		sınıf seviyesinde tanımlanan ve instance variable olarak da geçen 
		değişkenler ilgili sınıfın altındaki tüm metodlarda geçerlidir.</p>
        <h3>Stringler ve Tarihler</h3>
        <p>VBA&#39;de olmayan birçok faydalı metod ve property var, bunlara burda değil, örnekler sırasında değineceğiz</p>
        <h3>Formlar</h3>
        <p>VB.Net&#39;te Userform yerine Formlar var ve bunlar bir sınıftır. Sınıf tanımlama şeklinde yaratılırlar, ancak çoğu zaman manuel yaratmak yerine menülerden yaratacağız(aslında tanımlayacağız). Bu sınıftan bir form yaratıp çağırmadığımız sürece o form yaratılmaz.</p>
        <p>Bir sınıf olduğu için aynı form için birden fazla nesne yaratabiliyoruz. Pratikte çok karşılaşılmamakla birlikte bazen gerekebilmektedir.</p>
        <p>VBA&#39;deki vbModal ve vbModeless etkisi,&nbsp; VB.Nette sırasıyla ShowDialog() ve Show() metodlarıyla yapılır.</p>
        <p>Form üzerinde oluşturduğunuz her bir controlün bilgisi &quot;Form1.Designer.vb&quot; dosyasında bulunur.</p>
        <p>Controllerin çoğu aynı olmakla birlikte kimisinin adı(CommandButton--&gt;Button) ve özellik isimleri(CommandButton.Caption--&gt;Button.Text) değişti. Başka bir örnek: ListBox1. AddtIem&nbsp; ListBox.Items.Add() oldu. Diğer değişenbelli başlı controler şöyle :</p>
        <p>
            <img alt="" src="/images/vsto_formcontrols.jpg" /></p>
        <p>Form sınıfı oldukça gelişmiş bir sınıf olup hem üzerine alabileceği control sayısı çok fazla hem de kendisine ait event sayısı fazladır. VBA’deki Userformlarda 22 event varken Form sınıfında 85 event vardır</p>
        <h3>Diğer</h3>
        <p>Diğer bilgileri .Net işlemleri bölümünde bulabilirsiniz.</p>
        

    </div>

    <h2 class='baslik'>C#</h2>
    <div class='konu'>
   <p>Microsoft, .Net platformu ile birlikte Vb.Net'e ek olarak,ve aslında Java&#39;ya rakip olarak c# dilini de kullanıma sürmüştür. Bu dil, C ailesinden bir dil olmakla birlikte Delphi ve Pascal gibi eskinin güçlü dillerinin de güçlü özelliklerini almıştır. Size önerim de kendinizi VB.net’ten ziyade c#&#39;ta geliştirmenizdir. Zaten stackoverflowdaki arama sayılarına baktığınızda c# etiketlemesi, vb.net&#39;ten kat kat daha fazladır.</p>
        <p>C ailesi demek süslü parantezlerin ve noktalı virgüllerin hayatımıza girmesi, &quot;End ....&quot; ifadesinin de çıkması demek . Örnek bir kod bloğu şöyle:</p>
        <pre class="brush:csharp"> private void hello() {
    string mesaj="Merhaba dünya";
    Console.WriteLine(string);
}</pre>
        <p> Koşul ve döngü blokları ise şöyle</p>
        <pre class="brush:csharp"> if () {

}

do {

}</pre>
        <p> <strong>Yukarıda Vb.Net için anlatıtığımız birçok detay c# için de geçerli 
		demiştik. Şimdi, kısaca c#&#39;ta farklı olarak ne var, o</strong><span style="font-weight: bold">nlara 
		bakalım.</span></p>
        <h5 id="missingc4"> Missing parametresi ve c# 4.0</h5>
		<p> Uzun bir süre kodlarımı c# yerine VB.Nette yazmama neden olan bir 
		açıklama göstermek istiyorum size, MSDN'den:</p>
        <blockquote> <em>In general, developers who use Microsoft Visual Basic®&nbsp;.NET have an easier time working with Microsoft Office objects than do developers who use Microsoft Visual C#® for one important reason: Visual Basic for Applications (VBA) methods often include optional parameters, and Visual Basic .NET supports optional parameters. C# developers will find that they must supply a value for each and every optional method parameter, whereas Visual Basic .NET developers can simply used named parameters to supply only the values they need. 
			</em> </blockquote>
        <p>Türkçesi: c#, optional parameterleri desteklemediği ve VBA 
		metodlarının da çoğu optional parametre içerdiği için, c# kodlarında 
		optional parametrelerin yerine geçmesi için çok çirkin görünen
		<span class="keywordler">missing</span> ifadesini eklenmek durumundaymış. 
		Ör: FalancaFonksiyon(arg1,arg2,missing,missing,missing). Üstelik bu 
		missing'lerin sayısı bazen 30-40'ı bulabiliyordu.</p>
        <pre class="brush:csharp">Excel.Workbook wb = ThisApplication.Workbooks.Add(Type.Missing);</pre>
		<p>Ancak güzel haber şu: c# 4&#39; ten sonra artık optional parametre desteği var, yani artık missing eklemek zorunda değiliz.</p>
        <pre class="brush:csharp"> Excel.Workbook wb = ThisApplication.Workbooks.Add();</pre>
		<p> 
		Bir diğer can sıkıcı açıklama ise şuydu:</p>
        <blockquote> <em>In addition, C# doesn&#39;t support properties with parameters other than indexers, yet many Excel properties accept parameters. You&#39;ll find that properties such as the&nbsp;</em><strong><em>Application.Range</em></strong><em>&nbsp;property, available in VBA and Visual Basic .NET, require separate accessor methods for C# developers (the&nbsp;</em><strong><em>get_Range</em></strong><em>&nbsp;method replaces the&nbsp;</em><strong><em>Range</em></strong><em>&nbsp;property.) Watch for differences between the languages like these throughout this document.</em></blockquote>
        <p> Bunun Türkçesi 
		de "c#, parametreli propertyleri desteklemiyor, bu propertler için 
		'get_' ile başlayan metodlar üretilmiştir. Ör: Range propertysi yerine 
		get_Range metodu gibi" olup, bu durum da düzelmiş durumda. Yani artık garip garip get_Range, get_Offset gibi şeyler yazmıyoruz. 
		Zaten yazarsanız, böyle bir metod yok diye hata alırsınız.</p>
		<p> 
		Özetle 
		VSTO dünyasında c# ile çalışmamanız için artık hiçbir neden yok.</p>
		<p> 
		İki dil kıyaslamasını birçok yerde bulabilirsinz. Ben örnek olarak
		<a href="https://en.wikipedia.org/wiki/Comparison_of_C_Sharp_and_Visual_Basic_.NET">
		birini</a> koyuyorum. Gerçi ben c#'ı tek geçiyorum ama bunu yaşadıkça 
		görmeniz lazım.</p>
  </div>



    <h2 class='baslik'>Kıyaslama ve VBA&#39;den .Net dillerine dönüşüm</h2>
    <div class='konu'>
   


        <p>Burada bütün farklara değinmeyeceğim, sadece önemli olduğunu 
		düşündüklerimi buraya aldım. Ama
		<a href="http://help.autodesk.com/view/ACD/2015/ENU/?guid=GUID-C4B063BA-EAAA-430F-BDB5-2C48F1D897E4">
		şu sayfada</a> detaylı bir kıyaslama bulabilirsiniz.</p>
        <p>Aşağıdaki gösterimler vb.net ile c#ta önemli bi fark yoksa sadece vb.neti vereceğim.</p>
        <h3>Metodlar/Fonksiyonlar</h3>
        <h4>Sub/Fonksiyon/Metod ayrımı</h4>
        <p>VBA&#39;de sadece fonksiyon kavramı varken, VB.Net&#39;te hem fonksiyon hem metod, c#&#39;ta ise sadece metod kavramı var. VBA&#39;de tanımlanan fonksiyonların VB.Net karşılığı Modül fonksiyonu veya Shared sınıf metodudur, c#&#39;ta da static sınıf metodur. VBA&#39;in Sub prosedürünün VB.Net karşılığı da yine Sub prosedürdür, c# karşılığı ise void dönüş tipli metodlardır.</p>
        <pre class="brush:vb">Private Sub Prosedur()
'.....
End Sub

Private Function Fonksiyon()
'.....
End Function</pre>
		<pre class="brush:csharp">private void subprosedur_veya_donusdegersiz_fonksiyon() {
   //.....
}

privade int donus_degerli_fonksiyon() {
  //......
}</pre>
        <h4>Parantez kullanımı</h4>
        <p>VBA&#39;da metodlar parantez almazken, fonksiyonlar alabiliyor. Üstelik aynı isme sahip olan metod-fonksiyonlar sözkonusu olup bunlar da kafa karıştırıcı olabilmekteydi. 
		.Net dillerinde ise parantez zorunludur. VB.Net&#39;te fonksiyon kavramı hala varken, c#&#39;ta fonksiyon diye ayrı bir yapı bulunmuyor.</p>
        <pre class="brush:vb">MsgBox &quot;Merhaba&quot; &#39;Fonksiyon1: Dönüş tipi yok, sadece bir iş icra ediyor
cevap=MsgBox(&quot;Emin misin&quot;,vbYesNo) &#39;Fonksiyon2: Dönüş tipi var
ActiveCell.Clear &#39;Metod: Bir nesneye ait fonksiyon</pre>
		<p>VB.Net'e bakalım</p>
        <pre class="brush:vbnet">MsgBox(&quot;Merhaba&quot;) 'Bu arada c#'ta MsgBox kullanımı bulunmuyor. System.Windows.Forms.MessageBox metodu kullanılıyor. Bu Vb.Net'te de kullanılabilmektedir.
cevap=MsgBox(&quot;Emin misin&quot;,vbYesNo)
ActiveCell.Clear()</pre>


        <h4>Parametre tipleri</h4>
        <p>VBA&#39;de parametreler default olarak ByRef olarak geçirilirken, .Net dillerinde ByValue&#39;dur. O yüzden özellikle
		<strong>VBA&#39;den VSTO&#39;ya kopyaladığınız kodlarda buna dikkat etmelisiniz.
		</strong> </p>
        <h4>Fonksiyon dönüşü</h4>
        <p>VBA&#39;de fonksiyonların son satırı fonksiyonun adı ile aynı idi, VB.Net&#39;te bu hala yapılabilirken aynı zamanda 
		<span class="keywordler">Return</span> ifadesi de kullanılabilir. c# metodlarında ise sadece <strong>Return</strong> kullanılır.</p>
        <pre class="brush:vbnet">Function Deneme(param1 As String)
    &#39;....
    Temp=1
    Return Temp &#39;veya Deneme=Temp
End Function</pre>
        <h4>Fonksiyon çağrılması</h4>
        <p>VBA&#39;de fonksiyonlar doğrudan çağrılabilmekte iken, VB.Net&#39;te bir modül içinde tanımlanan fonksiyonlar yine doğrudan çağrılabilirken, bir sınıf içinde tanımlanmış metodlar sınıf adıyla çağrılırlar.&nbsp; C#&#39;ta da sadece sınıf metodları vardır. Tabi burada 
		<strong>Shared(Static)</strong> class yapısından da bahsedeceğiz. Sınıflardan türetilen nesnelerde ise metodu kullanımı zorunludur.</p>
        <pre class="brush:vb">sayi = Sqr(4) 'Built-in Fonksiyon
sayi2 = Math.Sqrt(4) &#39;Shared sınıf metodu, sınıftan nesne yaratmadan doğrudan kullandık
sayi3 = MyFunc(10) &#39;Modül fonksiyonu(UDF)

Dim obje As New MyClass() 'Bu sınıfın Kokal diye bir metodu olduğunu düşünelim
sayı4 = obje.Kokal(4) 'sınıftan bi nesne yaratarak kullandık </pre>
        <h4>Overloading</h4>
        <p>VBA&#39;de olmayıp .Net dillerinde olan bir özellik de metodların overload edilebilmesi. Yani aynı isimde, ancak farklı parametre kümesiyle birden fazla metod olabilir.</p>
        <pre class="brush:vbnet">Benimmetod(Sayi As Integer)
Benimmetod(Sayi As Integer, deger As Long)
Benimmetod(Sayi As Integer, adres As String)
</pre>
        <h3>Propertyler</h3>
        <p>
            VBA&#39;de default propertyler yazılmadan geçilebiliyorken, .Net dillerinde böyle birşey sözkonusu değildir. 
			İlgili property açıkça yazılmalıdır.</p>
        <pre class="brush:vb">deger=ActiveCell 'Value propertysi default propertydir</pre>
		<p>Bir de .Net'e bakalım</p>
        <pre class="brush:vbnet">deger=ActiveCell.Value</pre>
        <p>NOT: Activecell.Value yazmak bu kadar basit değil ama prensipte böyle, basit olması adına böyle yazdım. Daha doğru yazımı sonra göreceğiz. 
		Yukarıda Veri Tipi Dönüştürme ve onun altındaki DirectCast/CType 
		konusunda biraz bahsetmiştim. Detayları VSTO kodlarında göreceğiz.</p>


         <h3>Indeksler</h3>
        <p>
            VBA&#39;de diziler 
            0&#39;dan da 1&#39;den de başlayabiliyor. Collectionlar&#39;ın indeksi ise hep 1&#39;den başlıyor. Form elemanlarında da indeksler 0&#39;dan başlıyor. .Net dillerinde ise indeks hep 0&#39;dan başlar.</p>
        <p>Bununla birlikte, VB.Net Excel&#39;in kendisiyle ilgilendiğinde mecburen Excelin sınırlamalarına boyun eğer ve onun indeks mantığını kullanır. Mesesa sayfaların indeksi Vb.Net&#39;te de 1&#39;den başlar.</p>
		<h4>c# farkı</h4>
		<p>Indeksler Vb.Net'te VBA'de olduğu gibi () işaretleri arasında 
		gösterilirken c#'ta [] işaretleri arasında yazılır.</p>
		<pre class="brush:csharp">adet = Cells[1,1].Value</pre>
		<h3>Nesnelere değer atama</h3>
        <p>
            VBA&#39;de nesnelere Set ifadesi ile atama yapılırken .Net dillerinde ise Set ifadesi bulunmuyor. Ayrıca default propertyler nedeniyle VBA&#39;de bir nesne hem kendisi olarak hem de default propertysi olarak aynı şekilde atanabiliyor, ki bu da karışık bir görüntüye neden oluyordu, 
			halbuki az önce gördüğümüz üzere .Net'te default property kullanımı 
			yoktur.</p>
        <pre class="brush:vb">deger=ActiveCell &#39;default property ataması
Set hucre=ActiveCell</pre>
		<p>.Net'e bakalım</p>
        <pre class="brush:vbnet">
deger=ActiveCell.Value
hucre=ActiveCell &#39;bu kadar basit değil ama prensipte böyle, daha doğru yazımı sonra göreceğiz</pre>

        <h3>Değişkenler ve Veri Tipleri</h3>
		<h4>Tanımlama zorunluğu</h4>
		<p>VBA anlatırken "Değişkenlerinizi, tipi belli olacak şekilde(yani 
		variant değil) tanımlamak faydalı olur" diyorduk, 
		.Net'te artık bu faydanın size zorla dayatılması sözkonusu. c#ta her 
		değişken mutlaka&nbsp; belli bir tiple tanımlanmalıyken, Vb.Net'te biraz size bırakmışlar;
		<strong>Option Strict On</strong> yapılarak bu zorunluluk sağlanır. c# 
		değil de Vb.Net'i kullanmaya karar verirseniz bu ayarı yapmayı 
		unutmayın. Bunu kod olarak da, Property menüsünden de yapabiliyorsunuz.</p>
        <h4>Variant vs Object</h4>
        <p>
            VBA&#39;de hem Variant hem Object tipleri varken, .Net dillerinde sadece Object vardır. Variantlar da objecte dönüşmüştür.</p>
        <h4>
            Tanımlama</h4>
        <p>
            VBA&#39;de deklerasyon(tanımlama) ve atama ayrı satırda yapılırken .Net 
			dillerinde tek satırda yapılabiliyor. </p>
        <pre class="brush:vb">Dim a As Integer
a = 1</pre>
		<p>.Net'e bakalım</p>
        <pre class="brush:vbnet">
Dim a As Integer = 1</pre>
		<p>
		Ayrıca VBA'de aynı veri tipteki değişkenler tek seferde veri tipi sadece 
		bir kez belirtilerek tanımlanamazken .Net dillerinde tanımlanabilir.</p>
        <pre class="brush:vb">Dim a As Integer, b As Integer &#39;Dim a, b As Integer deseydik a Variant olurdu</pre>
		<p>.Net'e bakalım</p>
        <pre class="brush:vbnet">
Dim a, b As Integer &#39;ikisi de integer</pre>
		<pre class="brush:csharp">
int a ,b;
</pre>
        
        <h3>Enumerations</h3>
        <p>VBA'da global olarak tanımlı bulunann xlCenter tarzındaki 
		enumerationlar, .Net'te artık global olarak bulunmaz ve doğrudan 
		kullanılamazlar. Bunlar uzun uzun yazılmalıdır, ancak neyseki 
		intellisense var. Üstelik Visual Studion'nun sürekli gelişen 
		versiyonalrıyla biz = düğmesine bastığımızda ne seçmemiz gerektiği 
		anında çıkıyor.&nbsp;</p>
        <pre class="brush:vb">ActiveCell.HorizontalAlignment = xlCenter
CurrentRegion.Sort(..., xlAscending)
Application.Cursor = xlDefault</pre>
		<p>.Net'teki duruma bakalım</p>
        <pre class="brush:vbnet">ActiveCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
CurrentRegion.Sort(..., Excel.XISortOrder.xlAscending) 
Application.Cursor = Excel.XIMousePointer.xlDefault</pre>



        <p>
            Bir diğer değişiklik de Enumerationların constant değerleri yani 
			sayısal değerleri .Net'te yasaklı durumda. Yani kodlarınızı VBA'den 
			kopyalıyorsanız bunlara elle müdahale etmeniz gerekecektir.</p>
		<pre class="brush:vb">Application.CutCopyMode = 1
CurrentRegion.Sort(..., 1) </pre>
		<pre class="brush:vbnet">Application.CutCopyMode = Excel.XICutCopyMode.xlCopy
CurrentRegion.Sort(..., Excel.XISortOrder.xlAscending)</pre>
        <h3>
            VBA kodlarınızın .Net'e dönüştürülmesi</h3>
        <ul>
            <li>Modül kodlarını aktarmak sıkıntı değil, ancak Userformlarınızı baştan yapmak zorundasınız.</li>
			<li>Kodlarınızı öncelikle VB.Net'e çevirebilirsiiniz. Bunun için 
			standart modül kodlarınızı Visual Studio'da açtığınız bir modül 
			içine kopyalayın. VS, hatalı yerlerin altını çizecektir. Sonra 
			gerekli düzeltmeleri yapın.</li>
			<li>Vb.Net'e çevirdiğinizi kodları
			<a href="http://converter.telerik.com/">şu sayfada</a> c# koduna 
			dönüştürebilirsiniz</li>
            <li>ADO kodlarınızı olduğu gibi bırakabilirsiniz ancak ADO.Net tercih sebebi olmalı, o yüzden yeniden yazmalısınız.</li>
            <li>VBA kodlarını VSTO kodunuz içinde de çalıştırabiliyorsunuz. Bunu 
			hiç denemedim ancak yapılabiliyor.</li>
        </ul>
        <p>Aşağıdaki iki sayfadan daha detaylı bilgi alabilirsiniz.</p>
        <p>
            <a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa192490(v=office.11)">https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa192490(v=office.11)</a></p>
        <p><a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb960898(v=office.12)">https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb960898(v=office.12)</a></p>
          </div>
</asp:Content>
