<%@ Page Title='Excel Nesne Modeli' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Görsel Araçlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>
    <h1>Excel ile çalışmak</h1>
       <p><a href="../VBAMakro/Giris_ExcelNesneModeli.aspx">Excel obje modelini</a> ve <a href="../VBAMakro/DortTemelNesne_Konular.aspx">4 temel nesney</a>i iyi bildiğinizi varsayarak 
		başlıyorum. Bilmiyorsanız önce bu linklere bakın, zira bu sayfada bunların ne anlama geldiği açıklanmayacak.</p>
        <p>Temel kaynağınız: <a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa168292(v=office.11)">Understanding the Excel Object Model from a .NET Developer's Perspective</a> 
		dokümanı ama içindeki bazı bilgiler eski, 
		özellikle <strong>c#'ın desteklemediğini </strong>söylediği şeylerin çoğu artık destekleniyor. 
		O yüzden bu linki bi yere kaydedin ve benim sitedeki okumalarınızı 
		bitirdikten ve biraz pratik yaptıktan sonra mutlaka bu linke tekrar 
		bakın ve konuyu içselleştirin.</p>
		<p>Şimdi .Net dünyasında Excel ile çalışırken iki yolumuz bulunmakta. Bunlar;</p>
		<ul>
			<li>Excel API'si kullanmak. Bunun için bir VSTO projesinde(Excel 
			Add-in) hazır olarak gelen <strong>Microsoft.Office.Interop.Excel</strong> 
			kütüphanesi kullanılır.(Bir de Office.Tools.Excel var ama bunu hiç 
			kullanmayacağız)</li>
			<li>3rd Party kütüphanelerin kullanımı.</li>
		</ul>
		<p>Biz burada ağırlıklı olarak ilk yöntemi kullanacağız. Ama ikinci 
		yöntem hayat kolaylaştırıcı yöntemdir, o yüzden bunlara da değineceğiz. 
		Bu kütüphaneler o kadar güçlüdür ki, Excel'iniz kapalıyken de işlem 
		yaparlar, hatta Excel kurulu olmasa dahi excel dosyaları 
		yaratabilirler. Hayat kolaylaştırıcılıkları ise 1.yöntemde belli işler için 
		uzun uzun yazdığınız kodlar yerine tek satırlık kodlarla işinizi 
		halletmenizi sağlamasındandır.</p>
	<p>NOT: Şimdilik kodları sadece takip edin. Sonraki sayfadaki İlk VSTO 
	Add-inimizi yaparken uygun yerlere bunları kendiniz de koyup 
	deneyebileceksiniz.</p>
    <h2 class='baslik'>Excel API'si(Office.Interop)</h2>
    	<div class="konu">
		<p>Bu yöntemde Excel nesneleriyle çalışabilmemiz için kodunuzun tepedeki 
		using directivelerinin olduğu kısımda şu satırın yazması gerekir.</p>
		<pre class="brush:csharp">using Excel = Microsoft.Office.Interop.Excel;</pre>
		<p>Vb.Net'te Imports kısmında aynısını yapabileceğiniz gibi, <strong>MyProject&gt;References </strong>içinde aşağıdaki gibi 
		ekleyerek de bunu yapabilirsiniz.</p>
        <p>
            <img alt="" src="/images/vsto_ExcelInterop.jpg" /></p>
			<p>
            Bu arada VBA'de kullandığımız MsgBox, InputBox gibi fonksiyonlar 
			VB.Net'te yine kullanılabilir durumdadır. Ancak bir şekilde c#'ta da 
			bunları kullanmak isterseniz, projenize Microsoft.VisualBasic 
			kütüphanesini eklemek gerekecektir. Özellikle InputBox kullanımı 
			için bu kütüphaneye başvuracağız, zira c#'ta böyle bir 
			fonksiyon(metod) yok.</p>
        <h3>
            Globals sınıfı</h3>
        <p>
            Bir sonraki sayfada göreceğimiz gibi, bir Add-in projesi yarattığımızda VS&#39;nin bazı sınıfları otomatik oluşturacaktır. Bunlardan biri de Globals sınıfıdır. Bu sınıf ThisAddIn.Designer.cs dosyası içinde yer alır. Kod bloğunun içinde ne yazdığını bilmemize gerek yok. Bu sınıfla ilgili bilmemiz gereken şey şu: Bunun sayesinde aşağıdaki 
			nesnelere erişim sağlayabiliyoruz.</p>
        <ul>
            <li>Document Level projelerde ThisWorkbook ve Sheet<em>n</em> sınıflarına. Ör: <strong>Globals.ThisWorkbook</strong></li>
            <li>Application level projelerde ThisAddin sınıfı. <strong>Globals.ThisAddin</strong></li>
            <li>Ribbon Designerda tasarladığımız Ribbonlara. <strong>Globals.Ribbons.Ribbon1</strong></li>
        </ul>
        <p>Bu, şu demek oluyor. VBA&#39;de doğrudan kullandığımız ActiveCell veya ActiveSheet gibi nesneler vardı. Document 
		Level bir projede ThisWorkbook sınıfı dışından, Application Level bir projede de ThisAddin sınıfı dışından bunlara direkt ulaşamazsınız. Mesela bir Ribbondan 
		(veya bir başka sınıf içinden) ulaşmak için bunlara Globals sınıfı üzerinden 
		erişmeniz gerekir. Aşağıda çeşitli örnekler var(basitlik adına şuan için conversion yapmıyorum)</p>
        <pre class="brush:csharp">//Doc level proje ThisWorkbook içi
ThisWorkbook.ActiveSheet //veya this.ActiveSheet
ThisApplication.... //veya this.Application

//Doc level proje Ribbon
Glboals.ThisWorkbook.ActiveSheet....
Globals.ThisWorkbook.Application

//App level proje ThisAddin içi
this.Application.ActiveCell....

//App level proje Form1 sınıfı
Globals.ThisAddIn.Application.ActiveCell....</pre>
        <h3>4 Temel Sınıf/Nesne</h3>
        <h4>Application</h4>
        <p>Yukarıda&nbsp; gördüğümüz gibi Application nesnesine Globals sınıfı 
		üzerinden ulaşacağız. Size tavsiyem, her defasında Application&#39;ı bu şekilde uzun uzun yazmak yerine bunu 
		ilgili sınıfta public variable(Ör:app) olarak tanımlayıp sonra bunu 
		kullanmanızdır. Veya çok fazla sınıfı olan bir uygulamanız olacaksa her 
		sınıfta ayrı ayrı değişkenler tanımlamak yerine static bir sınıf 
		oluşturun ve bu sınıfın içine bir kez tanımlayın ve her defasında bunu 
		çağırın. Vb.Net'te çalışıyorsanız "genel modül&quot; ismini vereceğiniz bir modülde de yapabilirsiniz. Burada Vb.Net&#39;in küçük bi avantajı var; 
		bu değişkenin önünde modül adı gibi birşey belirtmeye gerek olmuyor.</p>
        <p><strong>Statik</strong> sınıf detaylarını ve neden static sınıf kullandığımızı <a href="https://docs.microsoft.com/tr-tr/dotnet/csharp/programming-guide/classes-and-structs/static-classes-and-static-class-members">şuradan</a> görebilirsiniz. (static detayını bilenler bu parantezli kısmı geçebilir. Özetlemek gerekirse statik sınıflardan biz <strong>nesne yaratmayız</strong>, ona ait metod ve propertyleri <strong>doğrudan</strong> kullanabiliriz. En bilinen örneği Math sınıfıdır. Çeşitli matematik fonksiyonlarının bulunduğu bu sınıfı kullanmak için bu sınıftan bi nesne yaratmanın bi esprisi yok, sınıfı doğrudan kullanabilmeliyizdir. Halbuki form gibi bir nesnede ise, Form sınıfından bir nesne yaratıyor(aslında Form sınıfını inherit eden Form1 veya başak bir isim verdiğiniz bir form sınıfı), sonra bu nesnenin Show metodunu kullanıyorduk, doğrudan Form(Form1) sınıfını kullanamıyorduk. Bizim örneğimizde de bi utility sınıf oalrak MyStatik adında bi sınıf yaratacağız ve bundaki app değişkenini kullanacağız)</p>
        <pre class="brush:csharp">//Statik sınıfı ve değişkeni tanımlama
using Excel = Microsoft.Office.Interop.Excel;
            
namespace VSTOcsharp
{
    static class MyStatik
    {
        public static Excel.Application app = Globals.ThisAddIn.Application;
    }
}

//projede herhangi bir yerde kullanımı
MyStatik.app.ScreenUpdating=false; //Statik sınıf adını da öne koyuyoruz</pre>
        <p>
            VB.Net tarafında yukarıdaki kullanıma ek olarak aşağıdaki modül kullanımına bakalım.</p>
        <pre class="brush:vbnet">&#39;Genel modül içi
Public app As Excel.Application = Globals.ThisAddIn.Application

&#39;projede herhangi bir yerde kullanımı
app.ScreenUpdating=False &#39;Modül adını öne koymaya gerek yok</pre>
        <h4>
            Workbook</h4>
        <p>Buradan itibaren yukarıdaki <strong>app</strong> değişkenini kullanacağız. 
		VBA'de buna ihtiyacımız yoktu, zira Application nesnesi default 
		nesneydi. VSTO'da ise işler değişiyor. Ancak bir kez app nesnesi elimizdeyken wb ve ws nesnelerine erişim kolay 
		olacaktır.</p>
        <p>Yeni bir dosya açıp bunu bi değişkene atayalım. Bunu tek satırda 
		yapabileceğimizi biliyorsunuz artık.</p>
        <pre class="brush:csharp">
Excel.Workbook wb = app.Workbooks.Add();</pre>
        
        <p>
            VBA'de bunu şöyle yapardık</p>
        <pre class="brush:vb">
Dim wb As Workbook
Set wb = Workbooks.Add</pre>
        
        <p>
            VSTO'daki farkı özetleyecek olursak</p>
	<ul>
		<li>Tek satırda değişken tanımlayıp atama yaptık</li>
		<li>Set ifadesi kullanmadık, zaten bu ifade artık yok. Ayrı satırlarda 
		yapsaydık bile kullanmazdık</li>
		<li>Add metodu sonunda parantez kullandık</li>
		<li>Workbooks collectionu önünde app nesnesini kullandık.</li>
	</ul>
	<p>
            Bir dosyayı açmak için;</p>
        <pre class="brush:csharp">MyStatik.app.Workbooks.Open(@&quot;E:\OneDrive\Uygulama Geliştirme\web sitelerim\Yeni Efendi\Ornek_dosyalar\CF.xlsx&quot;);</pre>
        <p>
            NOT: Buradaki @ işareti, klasör ayracı olan \ işaretini bir kez kullanmamızı sağlar, aksi halde \\ yazmak lazımdı. 
			Zira c#'ta \ işareti özel bir karekter(‘escape character’ denir) olup takip eden başka 
			karakterlere özel anlam katar. Örneğin \n satırbaşı anlamında, \t tab 
			sekmesi anlamındadır. O yüzden gerçekten \ işaretini kullanmak 
			istediğimizde önüne bir \ daha koyarak onun gerçek \ işareti 
			olduğunu vurgularız. Ama bu çok zahmetli olabileceği için @ 
			karakteri ile bu zahmetten kurtuluruz. Vb.Net'te ise buna gerek 
			yoktur, VBA'de olduğu gibi normal bir şekilde \ kullanımı 
			yapılabilir.</p>
        <p>
            Şimdi bir de c# 4.0 öncesinde, optional parametrelerin desteklenmediği döneme bakalım; korkmayın, artık bu kabus bitti :)</p>
        <pre class="brush:csharp">MyStatik.app.Workbooks.Open(@&quot;E:\OneDrive\Uygulama Geliştirme\web sitelerim\Yeni Efendi\Ornek_dosyalar\CF.xlsx&quot;,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing);</pre>
        <h4>Worksheet</h4>
        <p>app nesnesini elde edikten sonra worksheet kullanımı da kolaydır.</p>
        <pre class="brush:csharp">
Excel.Worksheet ws = MyStatik.app.Worksheets[1]; //index'e dikkat
MessageBox.Show(ws.Name); </pre>
        <p>
            Ancak worksheets collection'ının dönüş tipi Sheets olup, Sheets&#39;in de dönüş tipi VBA&#39;den bildiğimiz üzere object olduğundan 
			eğer bir değişkene atama yapmadıysak doğrudan kullanımda takip eden 
			üyelerini hemen göstermez. Yani aşağıdaki gibi kullanırken [1] yazıp 
			noktaya basınca intellisense çıkmaz, çünkü tipi henüz belli değil, 
			worksheet mi charts mı bilinmiyor.</p>
        <pre class="brush:csharp">
MessageBox.Show(MyStatik.app.Worksheets[1].Name); </pre>
        <p>
            Bu aslında garip bir durum, zira her ne kadar Worksheets(Sheets 
			değil) collection'ınını kullanmış olsak bile dönüş tipini Sheets yapmışlar. 
			Gerçi bu, VSTO&#39;ya özgü bir durum değil, VBA&#39;de çalışırken de aynı durum sözkonusudur. Neyse bunun sebebine takılmayalım. 
			Intellisense çıkmasını istiyorsanız bunu değişkene atamanızda fayda 
			var. Ama diyelim ki bunu bi değişkene atamak istemiyorsunuz, zira bu 
			nesneyi başka bir yerde kullanmayacaksınız, sadece bir kereliğine 
			bir özelliğine erişeceksiniz, ama yine de intellisensin de çıkmasını 
			istiyorsunuz. Bu da mümkün, ama bu örneği Range nesnesinde göreceğiz, oradaki ilgili açıklama worksheet'te 
			(ve tüm diğer nesnelerde de) aynen geçerlidir.</p>
        <h4>Range</h4>
        <p>Workbook ve Worksheet biraz kolaydı. Range&#39;de bazen <strong>
		casting</strong> yapmak durumunda kalacağız. Şimdi ilk olarak basit bir 
		örnekle başlayalım.</p>
	<p>Mesela aktif hücrenin adresini ve değerini öğrenmek istiyoruz. Bunun için yazacağımız kod aşağıdaki gibi olabilir.</p>
        <pre class="brush:csharp">
Excel.Range hucre = MyStatik.app.ActiveCell;
double deger = hucre.Value2;
MessageBox.Show(hucre.Address + " adresindeki hucrenin değeri:" + deger.ToString());
</pre>
        <p>Bu kodda çok kompleks birşey yok aslında. İlk satırda statik sınıfımızdaki app değişkeni üzerinden ActiveCell'e ulaşıyoruz, ikinci satır gayet açık. Son satırdaki 
		MessageBox.Show yazabilmek için de tepede <strong>using System.Windows.Forms</strong> olması gerektiği aşikar. </p>
        <p>
            Şimdi biraz daha şık bir kod nasıl yazılır, ona bakalım. Üstelik biraz daha teferruatlı konulara da girelim. Böylece hem biraz daha c# pratiği hem de VSTO pratiği yapmış olalım. </p>
        <p>
            Öncelikle, bir hata bloğu koyacağız, ki bir hücre seçilmemişken uyarı versin, ayrıca seçilen hücre nümerik bişey içermediğinde de farklı bi uyarı versin, 
            yani aslında iki hata bloğumuz olacak. Bir de, ölçeceğimiz değerin tipini integer yapalım istiyoruz. </p>
	<p>
            Bu sefer farklı olarak Activecell yerine Selection nesnesini kullanacağız&nbsp;ve kodların nasıl değiştiğini göreceğiz.</p>
        <pre class="brush:csharp">
try
{                
    Excel.Range hucre = MyStatik.app.Selection; //castinge gerek yok, çünkü zaten değişkenin tipini belirtiyoruz
    int deger = (int)(hucre.Value2); //double'dan integera dönüşüm
    MessageBox.Show(hucre.Address + " adresindeki hucrenin değeri:" + deger.ToString());
    //değişken atamasız durum
    MessageBox.Show(MyStatik.app.Selection.Value2.ToString()); //casting yapmadığımız için intellisense çıkmaz
    MessageBox.Show(((Excel.Range)MyStatik.app.Selection).Value2.ToString()); //intellisense çıkar
}
catch (NullReferenceException)
{
    MessageBox.Show("Seçili bir hücre yok, lütfen bir hücre seçip tekrar deneyin.");
}      
catch (Exception ex)
{
    if (ex.HResult==-2146233088)
    {
        MessageBox.Show("Şuan nümerik değeri olan bir hücrede bulunmuyorsunuz.");
        MessageBox.Show(String.Format("HRresult:{0},\n\nMesaj:{1}", ex.HResult.ToString(), ex.Message));
    }
    else
    {
        MessageBox.Show(String.Format("HRresult:{0},\n\nMesaj:{1}", ex.HResult.ToString(), ex.Message));
    }                
}
</pre>
        <p>
&nbsp;Şimdi de burayı inceleyelim</p>
        <ul>
            <li>Selection&#39;ın dönüş tipi Range değildir(Peki object mi? 
			Birazdan göreceğiz). Bunun için yapılabilecek 3 şey var. Aslında bu yazdıklarımı 2010'dan önce yazsaydım farklı bişeyler yazacaktım, şimdi yazdığımda ise 
			durum farklı, bunun sebebi dillerin zaman değişiyor olması<ul>
                <li>İlk olarak herhangi bir farklılığın olmadığı case, burada hücre değişkeninin tipini zaten Excel.Range belirlediğim için ilave bir işleme gerek yok</li>
                <li>Casting işlemi: (Excel.Range) ifadesini kullanarak casting yapıyoruz. Böylece compliera diyoruz ki&nbsp; &quot;MyStatik.app.Selection&quot;&#39;dan dönen şey bir objedir ama aslında bu bir Range nesnesidir, hadi bunu Range nesnesine dönüştür. Casting detayı için&nbsp; için <a href="Giris_NetDilleri.aspx">.Net dilleri sayfasına</a> bakın, özellikle orada link verdiğim wordpress sayfasına da bakın. Bu arada bunun 
				Vb.Net karşılığı CType(app.Selection, Excel.Range) olur. Biz bunu yaparak ilgili nesneyi Range&#39;e döndürdüğümüz için range nesnesinin tüm üyelerine erişebiliriz.</li>
                <li>Gelelim 2010 öncesi ve sonrası duruma: 2010'dan önce Cast etmeden kullanamıyorduk, ancak 2010'dan sonra castinge gerek olmadan da kullanabiliyoruz. 
				Peki neden ve farkı ne? 2010&#39;da c# 4.0 geldikten sonra,
				<span class="keywordler">dynamic</span> veri tipi diye birşey çıktı. Detayına girmeyeceğiz ancak, 
				bu özellik object tipli bir nesnenin gerçek nesneye 
				dönüştürülmeden de kullanılabilmesi imkanı verdi ve bunu da tipinin runtime sırasında compiler tarafından otomatikman belirlenmesiyle yapmaya başladı. 
				İşte bu Selection nesnesinin dönüş tipi de artık dynamictir. 2010&#39;dan önce object idi ve mutlaka cast edilmesi gerekiyordu. Özetle castinge gerek olmadan kullanabilirsiniz ama bunu tavsiye etmem, hem compilerı yormuş olursunuz, hem de intellisenseden faydalanamazsınız. Bununla birlikte dynamic veri tipinin çok faydalı olduğu yerler de vardır, bunları zaman içinde göreceğiz.</li>
                </ul>
            </li>
            <li>Gelelim ikinci satırdaki <strong>deger </strong>değişkenine. Bunu int tanımladık. Şimdi eğer ilk satırda elde ettiğimiz 
			<strong>hucre </strong>değişkeni olmasaydı ikinci satırdaki kodu &quot;int deger=<span style="color: red; font-size: large"><strong>(</strong></span><span style="color: #009933"><strong>(</strong></span>Excel.Range<span style="color: #009933"><strong>)</strong></span>MyStatik.app.Selection<strong><span style="font-size: large; color: red">)</span></strong>.Value2;&quot; şeklinde yazardık. Kırmızı parantezler içinde yazan kısım zaten bizim hucre değişkeni içn yazdığımızın aynısı. 
			Kırmızı parentez içine alarak bunu ayrı bir nesne haline getiriyoruz. 
			İşte bundan sonra bunun üyelerine ulaşabiliyoruz. Bu kırmızı paranteze almayıp doğrudan 
			Selection'dan sonra Value2 yazsaydık, c# derleyicisi bunu şöyle anlayacaktı: &quot;MyStatik.app.ActiveCell.Selection.Value2.ToString() 
			değerini Range&#39;e dönüştür&quot;, ve hata verecekti.</li>
            <li>2. catch bloğunda String formatlama da kullandık, bunu kendiniz yorumlamaya çalışın. 
			Bu arada else bloğuna gelecek bir örnek bulamadığım için if kısmına da yazdım, sonucunu görün diye.<br /></li>
        </ul>
        <h4>Cells collectionı</h4>
        <p>Cells, collection olarak kullanıldığında dönüş tipi Range olup herhangi bir şekilde 
		castinge gerek yoktur. Ancak parametreli kullanıldığında, yani satır ve sütun verilerek bir hücre elde edilmeye çalışıldığında bunun dönüş tipi 
		dynamic(2010'dan önce object) dönüş tipli olduğu için intellisense çıkmaz, çıkması için (Excel.Range) ile cast edilmelidir, 
		tabi eğer değişkene atama yapılmadan kullanılacaksa. Değişken ataması yapıldığında 
		ise ayrı bir castinge bunda da gerek yoktur, zira zaten değişkenin 
		tipini belirliyoruzdur.</p>
        <p>
            <img alt="" src="/images/vsto_returntype.jpg" /></p>
        <pre class="brush:csharp">MessageBox.Show(((Excel.Range)MyStatik.app.Cells[1, 1]).Value);</pre>

    </div>
    
        <h2 class='baslik'>3rd Party Kütüphaneler</h2>
    	<div class="konu">
<p>İlk projemize başlamadan daha fazla kod örneği ile sizi boğmak istemedim. 
Normalde konu bütünlüğü adına burada koymayı tercih ederdim ama sizi bir an önce 
ilk projenizle de buluşturmak istiyorum. Bu konuya hem VSTO ile hem c# ile haşır neşir olduktan sonra <a href="ThirdPartyKutuphanler_Konular.aspx">şurada</a> gireceğiz.</p>
			<p>Sonraki sayfada görüşmek üzere...</p>
</div>
</asp:Content>
