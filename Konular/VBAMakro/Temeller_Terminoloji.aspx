<%@ Page Title='Temeller Terminoloji' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Temeller'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Terminoloji</h1>
    <h2 class="baslik">Giriş</h2>
    <div class="konu">
<p>İlk kodlarımıza geçmeden önce kısa bir terminoloji bilgisi edinmemizde fayda var. Böylece ne yaptığımızı daha iyi anlıyor olabileceğiz.</p>

	<p>Visual Basic veya herhangi bir programlama dili kullanmış olanlar bilir, 
	bir program yazarken programımızı belirli anahtar kelimeler üzerine kurarız. 
	Bunlar temel olarak şu şekilde sınıflandırılabilir:</p>
	<ul>
		<li>Prosedürler</li>
		<li>Değişkenler</li>
		<li>Sabitler</li>
		<li>Diziler</li>
		<li>Koşullu yapılar</li>
		<li>Döngüler</li>
		<li>Hata ayıklayıcılar</li>
		<li>Kullanıcıyla iletişim</li>
		<li>Olaylar</li>
		<li>Nesneler</li>
		<li>Metodlar</li>
                <li>Özellikler</li>
		<li>Fonksiyonlar</li>
	</ul>
	<p>Şimdi ilk etapta bilmemiz gerekenlere bir bakalım.</p>
    </div>
<h2 class='baslik'>Prosedür ve Modüller</h2>
<div class='konu'>
	<h3>Prosedür</h3>
	<p><span>Gerek kendimizin yazdığı kodlar olsun gerek makro kaydederek 
	oluşturduğumuz kodlar olsun hepsi bir prosedürdür. İki tür prosedürümüz var.</span></p>
	<ul>
		<li><strong>Sub prosedürler:</strong>Çalışır ve birşeyler yapar, ama bir değer döndürmezler. 
		(Türkçede yordam olarak geçer)</li>
	</ul>
	<pre class="brush:vb">
	Sub prosedüradı()
	'Kodlar buraya yazılır
	End Sub	</pre>

	<ul>
		<li><strong>Function Prosedürler:</strong>Çalışması sonucunda (genelde) değer döndüren 
		prosedürlerdir. Bunlar da kendi içinde ikiye ayrılır. İlk grubun Excelin 
		yerleşik fonksiyonlarından hiçbir farkı yoktur, bunlara <a href="Fonksiyonlar_ExcelicinUDFKullaniciTanimliFonksiyonlar.aspx">Kullanıcı tanımlı fonksiyonlar</a>(UDF) denir. Genelde bir parametre/argüman alırlar. İkinci grupta ise VBA içinde kullandığımız ve döndürdüğü değeri yine VBA içinde kullanmaya devam ettiğimiz <a href="Fonksiyonlar_VBAicinUDF.aspx">VBA Functionlar</a> yer alır. </li>
	</ul>
	<pre class="brush:vb">
	Function functionadı()
	'Kodlar buraya yazılır
	End Function	</pre>
	<p>Makro konusunu işlerken ağırlıklı olarak Sub prosedürleri işliyor olacağız, yer yer Functionlara da değineceğiz. Fonksiyonları az önce verdiğim linklerde ayrıca inceliyor olacağız.</p>

	<p><strong>Fonksiyona</strong> benzeyen bir de <strong>metod</strong> kavramı var, ki ikisi 
	genelde birbiri yerine kullanılmaktadır, bununla birlikte terminolojide 
	aralarında küçük bir fark bulunur: Fonksiyonlar bağımsız çalışabilir, ancak 
	metodlar mutlaka bir nesneye ihtiyaç duyarlar, bağımsız çalışamazlar. </p>
	<p>Ör:<strong>Application.InputBox </strong>ifadesindeki InputBox, bir
	<strong>metod</strong> olup Application nesnesine ihtiyaç duyar. Aşağıdaki 
	tanımladan da görüleceği üzere Application'ın hem ikonundan hem de önündeki 
	Class ifadesinden bir alttaki Inputboxtan farklı olduğu görülmekte.</p>
	<p><img src="../../images/vba_terminonoloji1.jpg"></p>
	<p><strong>Normal</strong> <strong>Inputbox</strong> ise bir <strong>
	fonksiyon</strong> olup nesneye ihtiyaç duymaz. Bu fonksiyon, Interaction
	<strong>modülü</strong> içinde tanımlanmıştır, bir class içinde değil.&nbsp;</p>
	<p><img src="../../images/vba_terminonoloji2.jpg"></p>
	<p>Yukarıda, fonksiyonlar "genelde" bir değer döndürür dedik. Bazen, sebebini 
	anlayamadığım bir şekilde, bazı kişilerin(hatta Microsoft'un kendisinin bile) 
	değer döndürmeyen Function prosedürler yazdığını görüyoruz. Bu hem Modül 
	fonksiyonları hem de Class metodları için geçerlidir. Mesela, aşağıda Range 
	classının Activate metodunu görüyoruz. Deklerasyonu <strong>Function 
	Activate()</strong> şeklinde yapılmış.</p>
	<p><img src="../../images/vba_terminonoloji3.jpg"></p>
	<p>AddComment metodu ise sonunuda <strong>As Comment </strong>ifadesine 
	sahip. İşte bu "As" ile başlayan kısım bize fonksiyonun "dönüş değerini" verir.</p>
	<p><img src="../../images/vba_terminonoloji4.jpg"></p>
	<p>Eğer bir metod geriye bir şey döndürüyorsa bu arkaplanda kesinlikle bir Function presedür 
		olarak hazırlanmıştır, ancak değer döndürmüyorsa Sub prosedür olarak 
	hazırlanmış olabileceği gibi sebebini hala anlamadığım bir şekilde Function 
	prosedür olarak da hazırlanmış olabilir.</p>
	<h3>Modül</h3>
	<p>Prosedürlerimizi yazdığımız yerlere <strong>Modül </strong>denmektedir. 
	Bunlar Standart modül, Class Modül ve Userform modüller olabilir. Ancak biz 
	eğitimin büyük kısmında standart modüllerde çalışacağız. Class modüllere çok 
	az kod yazacağız. Eventler bölümünde ise Workbook ve Worksheet isimli class 
	modüllere(evet bunlar da class modüldür) kod yazacağız. </p>
	<p>VBE içinde Project penceresinde herhangi bir workbookta sağ tıklayıp 
	Insert'e gelince gördüğümüz seçenekler modül seçenekleridir.</p>
	<p><img src="../../images/vba_terminonoloji5.jpg"></p>
	</div>

<h2 class='baslik'>Yorum Satırları</h2>
<div class='konu'><p> Yukarıdaki örneklerde gördüğünüz üzere ' işareti ile 
	yorumlar yazılabilmektedir. Yorumlar yeşil renkte görünürler.</p>
	<p> Yorumlar önemlidir, özellikle bazı makroları arada bir çalıştırıyorsak 
	ve makronun nasıl çalıştığını aklımızda tutamıyorsak yorumlarda bunu 
	açıklamak akıllıca olacaktır. Bir diğer gerekçe de, biz gittikten sonra yerimize gelecek kişiye de yol gösterici birşeylerin olması gerektiğidir.</p>
	<p> NOT: Bir diğer yorum ekleme yöntemi <span class="keywordler">REM</span> 
	ifadesini kullanmaktır, buna neden gerek duyulmuş hiçbir fikrim yok, ben 
	genelde ' işaretini kullanıyorum.</p>
	<pre class="brush:vb">
Sub prosedüradı()
REM açıklamalar buraya yazılır
End Sub	</pre>
	
	</div>

<h2 class='baslik'>Sabırsızlananlar için küçük bir ara</h2>
<div class='konu'><p>Hadi şimdi biraz kod yazalım. <span>Boş bir dosya açalım.
	</span>Önceki konularda gördüğümüz 
	gibi Personal.xlsb dosyamıza gidip Modules'e sağ tıklayalım ve yeni bir modül 
	ekleyelim. Sonra aşğıdaki kodu oraya yapıştıralım.
	<span style="text-decoration: underline">Kodun içinde herhangi bir yerdeyken</span> 
	F5 tuşuna basalım <span>ve kodumuzu çalıştıralım. (Bunun yerine üstteki araç 
	çubuğunda Play işaretine de tıklayarak makronuzu çalıştırabilirsiniz)</span></p>

	<pre class="brush:vb">
Sub ilkörnek()
Range("A1").Select
Selection.Value = 10

Range("A2").Select
Selection.Value = 20

Range("A3").Formula = "=A1+A2"
Range("A4").Value = Range("A1").Value + Range("A2").Value
End Sub	</pre>

	<p>Bu örnekte önce A1 hücresini seçtim, sonra mevcut olan bu seçime 
	yani A1'e 10 değerini yazdırdım. Sonra A2 hücresini seçip oraya da 20 
	değerini yazdırdım. A3 hücresine ise A1+A2 toplamını formül olarak 
	yazdırdım, ancak A4 hücresini yine aynı toplamı değer olarak yazdırdım.</p>
	<p>Gördüğünüz üzere <strong>Value</strong> değerini hem değer atamada hem de 
	değer okumada kullanabildim. İşte bazı özelliklerin hem değer yazma hem 
	değer okuma durumu varken, bazısında sadece değer yazma, bazısında ise 
	sadece değer okuma olabilmektedir. <a href="https://msdn.microsoft.com/en-us/library/office/ff194068.aspx">MSDN'de Excel Nesne modeli</a> 
	incelemesi yaparken bunlar şu şekilde gösterilir.</p>
	<ul>
		<li>Return or sets : Hem atanır, hem okunur</li>
		<li>Sets: Sadece atanır</li>
		<li>Returns: Sadece okunur</li>
	</ul>
	<p>Şimdi bir kod daha yazalım. Bu sefer yeni modül eklemek yerine aynı 
	modülde, yukardaki kodun hemen altına şu kodları yapıştıralım.</p>

	<pre class="brush:vb">
Sub kucukyap()
Set Aktif = Selection
 
For Each s In Aktif
    buyuk = StrConv(s.Value, vbLowerCase)
    s.Value = buyuk
Next s
 
End Sub	
</pre>

<p>Bu kod da seçtiğimiz hücrelerdeki tüm harfleri küçük harfe dönüştürür. Deneme yapmak için A1 hücresine 
VOLKAN, B1 hücresine kendi adınızı, C1 hücresine de EXCEL yazın, sonra bu üçünü 
seçip VBE'ye gelip F5 ile kodu çalıştıralım.</p>
	<p>Şimdilik bu kadar yeter, isterseniz devam etmeden önce siz de biraz Makro 
	Kaydedici ile alıştırma yapın.</p>
	</div>
</asp:Content>
