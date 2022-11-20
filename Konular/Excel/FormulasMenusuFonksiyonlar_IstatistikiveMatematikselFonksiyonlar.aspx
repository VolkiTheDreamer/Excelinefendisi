<%@ Page Title='FormulasMenusu1 IstatistikiveMatematikselFormuller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='6'></asp:Label></td></tr></table></div>

<h1>Istatistiki ve Matematiksel Fonksiyonlar</h1>
	<p>Excelin müthiş bir matematiksel fonksiyon kütüphanesi bulunmaktadır. Tabi biz burada hepsine girmeyeceğiz, 
ancak MIS ağırlıklı çalışan kişilerin hedef kitlem olduğunu düşünerek, bize lazım olabilecek 
fonksiyonları inceleyemeye 
dahil edeceğiz. Bu şu demek, mutlak değer formülü olan <strong>ABS</strong> tam anlamıyla matematiksel bir 
fonksiyon olmakla birlikte zaman zaman biz MIS tarzı işlerde çalışanların da gereksinim duyduğu bir fonksiyon olmaktadır, o yüzden onu burada ele alacağız. Ancak sinüs, cosinüs gibi trigonometrik fonksiyonlar v.s kapsamımız dışında olacaktır.</p>

<p>Bu kategoriyi anlatırken kendi içimde mantıksal alt gruplamalar da yaptım, 
alt bölümleri de buna göre oluşturdum.</p>
	<p>Bu sayfadaki örneklerin yer aldığı koşullu ve istatistiki fonksiyon 
	dosyalarını <a href="../../Ornek_dosyalar/Formuller/istvekosul.zip">buradan</a> 
	indirebilirsiniz.</p>


<h2 class='baslik'>Küsurat, Tamsayı ve Yuvarlama formüller</h2>
<div class='konu'>
<p>Bu kısımda küsuratı olan sayılar nasıl ele alınır, ayrıca diğer yuvarlama işlemleri nasıl yapılır, bunlara bakacağız.</p>



<table class="alterantelitable">
<th width='15%'>Fonksyion</th><th width='30%'>Syntax</th><th>Ne işe yarar</th>
<tr><td>TRUNC</td><td>TRUNC(sayı;[basamak])</td><td>Sayıyı istenen miktardaki küsurattan kurtarır, yuvarlama yoktur. 
	Basamak belirtilmezse küsüuratı tamamen kaldırır.</td></tr>
<tr><td>INT</td><td>INT(sayı)</td><td>Sayıyı en yakın aşağı yönlü tam sayıya yuvarlar. Pozitif sayılarda trunc ile aynı görevde.</td></tr>
<tr><td>ROUND</td><td>ROUND(sayı;basamak)</td><td>Belirtilen basamak kadar yuvarlama yapar. Yuvarlamanın yönü yuvarlanacak kısmın 5'in ne kadar altında/üstünde olmasına göre değişir. Basamak parametresi 0 da olabilir, + ve - de.</td></tr>
<tr><td>ROUNDDOWN</td><td>ROUNDDOWN(sayı;basamak)</td><td>Belirtilen basamak kadar aşağı yuvarlama yapar. Basamak parametresi 0 da olabilir, + ve - de.</td></tr>
<tr><td>ROUNDUP</td><td>ROUNDUP(sayı;basamak)</td><td>Belirtilen basamak kadar yukarı yuvarlama yapar. Basamak parametresi 0 da olabilir, + ve - de.</td></tr>
<tr><td>MROUND</td><td>MROUND (SAYI;ÇARPAN)</td><td>Sayıyı, belirtilen çarpanın en yakın katına yuvarlar. 
	<strong>Sadece pozitif sayılarda çalışır.</strong></td></tr>
<tr><td>CEILING</td><td>CEILING(sayı,[hassasiyet])</td><td>Sayıyı, belirtilen çarpanın en yakın YUKARI katına yuvarlar</td></tr>
<tr><td>CEILING.MATH</td><td>CEILING.MATH(sayı,[hassasiyet],[mode])</td><td>Sayıyı, belirtilen çarpanın en yakın YUKARI katına yuvarlar. Excel 2013le geldi. Mode parametresi, negatif sayılar için</td></tr>
<tr><td>FLOOR</td><td>FLOOR(sayı,[hassasiyet])</td><td>Sayıyı, belirtilen çarpanın en yakın AŞAĞI  katına yuvarlar</td></tr>
<tr><td>FLOOR.MATH</td><td>FLOOR.MATH(sayı,[hassasiyet],[mode])</td><td>Sayıyı, belirtilen çarpanın en yakın AŞAĞI  katına yuvarlar. Excel 2013le geldi. Mode parametresi, negatif sayılar için</td></tr>


<tr><td>EVEN</td><td>EVEN(sayı)</td><td>Pozitif sayıları, kendinden en büyük 
	çift sayıya yuvarlar. Negatifleri aşağıya</td></tr>


<tr><td>ODD</td><td>ODD(sayı)</td><td>Pozitif sayıları, kendinden en büyük tek 
	sayıya yuvarlar. Negatifleri aşağıya</td></tr>


</table>



<p>Aşağıdaki çeşitli örnekleri inceleyelim</p>
	<p><img src="/images/fonkistatistik7.jpg"></p>
	<p>Örnek dosyayı <a href="../../../Ornek_dosyalar/Formuller/yuvarlamalar.xlsx">
	buradan</a> indirebilrsiniz.</p>
	<p>Burada özellikle <strong>MROUND</strong> ve <strong>CEILING</strong> farkına değinmekte fayda var. MROUND, 
	en yakın kata(yukarı veya aşağıda) yuvarlarken, CEILING yukarı yönlü 
	yuvarlar. Bu durumda 67.256 sayısının 5.000'in katı şeklinde yuvarlarken 
	MROUND en yakın kat olarak 65.000'e veya 70.000'e yuvarlama konusunda seçim 
	yapmalıdır. 65.000, 2.256 birim uzaklıkta iken 70.000, 2.754 birim 
	uzaklıktadır, o yüzden 65.000e yuvarlar. CEILING ise hep yukar yuvarladığı 
	için 70.000e yuvarlar.</p>
	</div>
	
	
	<h2 class="baslik">Grup(Aggregate) Fonksiyonları</h2>
	<div class="konu">
	<h3>Toplamlar</h3>
	<p><span class="keywordler">SUM</span>: Belirli bir alandaki sayıları 
	toplar. Başlangıç noktası ile bitiş noktası arasında ":" işareti olur. 
	Birbirine komşu olmayan alanları toplamak için aralara ";" konur. Parametre 
	olarak sabit değer de girilebilir.</p>
	<pre class="formul">=SUM(A1:A5)
=SUM(J4:J10;K4:K6;10) //komşu olmayan ve sabit parametre</pre>
	<p>Filtrelenmiş alanlarda SUM formülünü yazarsanız aradaki gizlenmiş 
	değerler de sonuca dahil olur. Sadece filtreli grubun toplamını almak için 
	<span class="keywordler">SUBTOTAL </span>fonksiyonunu kullanmanız gerekir.</p>
	<p>Yoksa aşağıdaki gibi sorun yaşarsınız. 486.satırın formülünde şu 
	yazmaktadır.</p>
	<pre class="formul">=SUM(D277:D346)</pre>
	<p><img src="/images/formul_istatistik1.jpg"></p>
	<p>Gerçi bazı şanslı durumlarda filtreli alan sıralı olabilir, o yüzden 
	formülü yazdığınızda ilgili alanda arada filtrelenmiş satır bulunmaz. O&nbsp; zaman şansınız 
	yaver gidebilr.</p>
	<p><img src="/images/formul_istatistik2.jpg"></p>
	<p>Formülümüz şudur:</p>
	<pre class="formul">=SUM(D11:D16)</pre>
	<p>Siz siz olunuz, işinizi şansa bırakmayın, filtreli alanlarda her zaman 
	SUBTOTAL kullanmaya çlışın.</p>
	<h3>Adetler</h3>
	<p>Aşağıdaki hücre grubunda çeşitli sayma işlemleri yapalım.</p>
	<p><img src="/images/formul_istatistik3.jpg"></p>
	<p><span class="keywordler">COUNT</span>: Bir alandaki sayısal değer içeren 
	hücreleri sayar.</p>
	<p><span class="keywordler">COUNTA</span>: Bir alandaki dolu hücreleri 
	sayar.</p>
	<p><span class="keywordler">COUNTBLANK</span>: Bir aladaki boş hücreleri 
	sayar.</p>
	<p>Bir alandaki toplam hücreyi saydıran bir fonksiyon olmamakla birlikte, 
	aşağıdaki formülle dolaylı yoldan saydırabiliriz. Satır ve sütundan oluşan 
	bir karenin alanını hesaplıyoruz aslında.</p>
	<pre class="formul">=COLUMNS(J13:M15)*ROWS(J13:M15)</pre>
	<h3>Diğer</h3>
	<p><span class="keywordler">PRODUCT</span>: Belirli bir alandaki değerleri 
	birbiryle çarpar. 10 parametreden oluşan bir katsayıyı tek tek çarpmak 
	yerine bunu kullanmak pratiktir.</p>
	<pre class="formul">=A2*B2*C2*D2*E2*F2*G2*H2*I2*J2
//bunun yerine
=PRODUCT(A2:J2)</pre>
	<p>Bunun SUM ile birleşmiş hali olan SUMPRODUCT'a 
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx#sumproduct">buradan</a> 
	ulaşabilirsniz.</p>
	<p><span class="keywordler">SUBTOTAL</span>:Filtrelenemiş listelerde sadece 
	filtreli verinin toplamını(veya diğer gruplanmış fonksiyon sonucnu) verir. 
	Evet, yanlış okumadınız, fonksiyonun adı her ne kadar SUBTOTAL olsa da alt 
	toplam almaktan başka işler de yapar. Filtreli gurubn ortalamasını, 
	maksimumunu v.s de alır. Zaten aşağıdaki resimde göründüğü gibi "(" 
	karakterine basınca hemen parametre listesi ortaya çıkar. Mesela 1, ilgniç bir 
	şekilde Toplam aldırmaya değil, ortalamayı veriyor. Toplam için 9 seçeneği 
	seçilmelidir.</p>
	<p>Bu arada filtrelenmiş alanların hemen altındayken menüden <strong>Σ(Toplam)</strong> işaretine 
	basılınca otomatikman SUBTOTAL(9;....) şeklinde bir formül oluşur.</p>
	<p><img src="/images/formul_istatistik4.jpg"></p>
	<p>Yukarıdaki gibi bir listede sık sık bölge ve ürün&nbsp; değiştirerek 
	diptoplamın ne olacığı görülmek istenirse iki tane
	<a href="DataMenusu_SiralamaveFiltreleme.aspx#slicer">Slicer</a> konarak bu 
	iş halledilir.</p>
	<p><strong>NOT</strong>:Bu formül, Data Menüsündeki SUBTOTAL aracında da kulanılan formül 
	olup, oradaki kullanımına <a href="DataMenusu_Outline.aspx">buradan</a> 
	ulaşabilirsinz.</p>
	<p><span class="keywordler">AVERAGE</span>:Belirli bir alandaki hücre 
	grubunun ortalamasını verir. Yanlız dikkat edilmesi gereken bir nokta var. Bu 
	fonksiyon, boş değerleri dikkate almaz. 10 hücre varsa ve 9u doluysa toplamı 
	9a bölererk ortalama hesaplar,10'a değil.</p>
	<p>Mesela, satışçıların 12 aylık 
	ortalama satış rakamını alırken bazı satışçlar bazı aylar hiç satış yapmamışsa veri 
	kaynağında burası boş da gelebilir 0 da. Boş gelme durumu daha çok tabular 
	formdaki bir listenin pivot table yapılması sırasında oluşacaktır(Eğer ki 
	özellikle "boş hücreler 0 görünsün" işaretlemesi yapılmadıysa)</p>
	<p>Yukarıdaki örnekte 0 yerine boş gelirse hatalı bir ortalama hesaplanabilir, zira onların 0 
	olarak işleme girmesi gereklidir. Ama bazı durumlarda gerçekten boş 
	hücrelerin ortalama hesabına dahil olmaması istenebilir. Mesela kişinin uzun 
	süreli bir hastalığı olduysa o dönemlerin boş geçilerek ortalamaya dahil 
	olmaması ve ortalamayı düşrüşmemesi gerekir.</p>
	<p>Mesela aşağıdaki tabloda ilk satırın ortalaması 115 olurken ikincisi 138 
	olacaktır.</p>
	<p><img src="/images/formul_istatistik5.jpg"></p>
</div>

	<h2 class="baslik">En küçük ve En büyükler</h2>
<div class="konu">
	<p><span class="keywordler">MIN</span>: Bir alandaki değerlerin en küçüğünü 
	verir. Sabit değerler de parametre oalrak girilebilir ve birbirlerinden ";" 
	ile ayrılır.</p>
	<p><span class="keywordler">MAX</span>: Bir alandaki değerlerin en büyüğünü 
	verir. MIN'deki açıklamalar geçerli.</p>
	<p>MIN ve MAX'ın söylemimize ters bir şekilde işlediğini bilmek önemlidir. Mesela 
	bir hesabın sonucunda çıkan değer <strong>en az </strong>100 olsun denirse, bunun için MIN 
	değil MAX kullanılmalıdır. Mesela, A5*B5 formülünün sonucu 67 çıktıysa, bunu 100 
	yapmak için formülümüz şöyle olmalı</p>
	<pre class="formul">=MAX(A5*B5;100) //67 VE 100'den hangisi büyükse onu alır, yani 100ü</pre>
	<p>Keza bir hesabın sonucunun <strong>en çok </strong>100 çıkması istniyorsa 
	da MAX değil MIN kullanılmalıdır. A5*B5 180 çıkıyorsa formülümüz şöyle 
	olmalıdır:</p>
	<pre class="formul">=MIN(A5*B5;100) //100</pre>
	<p>MIN ve MAX'ın bir diğer kullanımı da basit IF'li yapılara 
	alternatiftir. Yani A1'in 
	değeri 10'dan büyükse 10 olsun, yoksa A1 olsun demek istediğimiz birçok kişi 
	şöyle yazar:</p>
	<pre class="formul">
=IF(A1&gt;10;10;A1) //A1 3'se sonuç 3, 15se 10
//bunun yerine daha kısa olan şu fomrülü deneyin
=MIN(A1;10)</pre>
	<p><span class="keywordler">SMALL(dizi;n)</span>: Bir grup sayı içinde en 
	küçük n. sayıyı verir.</p>
	<p>Sayılar, bir alandaki hücreler olabileceği gibi, bir dizi sonucu da 
	olabilir. </p>
	<pre class="formul">=SMALL(A1:A10;3)
=SMALL({10;75;30;254;20};2) //20</pre>
	<p><span class="keywordler">LARGE(dizi;n)</span>:SMALL'daki açıklamalar 
	aynen geçerli olup, en büyük n.elemanı döndürür.</p>
	<p>SMALL ve LARGE'ın
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllerinde </a>kullanımı da oldukça yaygındır. O sayfayı da incelemenizi 
	tavsiye ederim.</p>
	<h4>Uç değerler hariç ortalama</h4>
	<p>Aşağıdaki formülle, belli bir kümedeki en küçük/büyük 1 değer hariç 
	ortalama alınır. Zira belli ki mart ayında tek seferlik yüksek bir satış 
	olmuş, Ekimde de büyük çaplı bir sorun olduğu için satışlar düşük gitmiş. 
	Ortalamayı hesaplarken bunları hariç tutmakta fayda var. </p>
	<p><img src="/images/fonkistatistik10.jpg"></p>
	<pre class="formul">=(SUM(B2:B13)-MIN(B2:B13)-MAX(B2:B13))/10
//veya
=(SUM(B2:B13)-SMALL(B2:B13;1)-LARGE(B2:B13;1))/10</pre>

<p>Bu formülde sabit bir değer olan 10'a böldük çünkü hem alttan hem üstten 1'er 
uç değeri çıkardık, geriye 10 ay kaldığını biliyoruz, ancak burası ay yerine 
sayısı sürekli değişebilen bir küme olsaydı 10 yerine (COUNT(B2:B13)-2) 
girebilirdik.</p>
	<p>Peki, ya uç değer olarak ikişer tane yani toplamda 4 değeri hariç tutmak 
	isteseydik nasıl yapardık? Formülü biraz uzatmamız gerekirdi:</p>
	<pre class="formul">=(SUM(B2:B13)-SMALL(B2:B13;1)-SMALL(B2:B13;2)-LARGE(B2:B13;1)-LARGE(B2:B13;2))/8</pre>
	<p>Peki ya 3,4 v.s uç değer hariç tutmak istesek? Tabiki yukarıdaki 12 
	adetlik kümede teknik olarak en fazla 5 uç değer hariç tutulabilir. Ama 
	mesela 100 adetlik kümelerde uçlardan 5'er tane yani toplamda 10 tane değeri 
	hariç tutup kalan 90 adedin ortalamasını almak istesek nasıl yapardık? Tek 
	tek 5 tane SMALL 5 tane de LARGE formülü yazacak değiliz, değil mi? Bunun 
	için <a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">
	dizi formüllerini</a> kullanacağız. Onu ilgili bölümde göreceğiz. Ancak dizi 
	formülünün kullanımı da karışık gelirse aşağıdaki gibi bir UDF de 
	kullanabilirsiniz.</p>
<h4 class="baslik">Uçhariçortalama ortalama</h4>
	<div class="baslik">
		<pre class="brush:vb">
	Function uçhariçort(alan As Range, Uç As Variant)
	Dim aratoplam As Double
	Dim enbüyükler As Double
	Dim enküçükler As Double
	 
	For i = 1 To Uç
	    enbüyükler = enbüyükler + WorksheetFunction.Large(alan, i)
	Next i
	     
	For i = 1 To Uç
	    enküçkler = enküçkler + WorksheetFunction.Small(alan, i)
	Next i
	      
	aratoplam = WorksheetFunction.Sum(alan) - enbüyükler - enküçkler
	uçhariçort = aratoplam / (alan.Count - Uç * 2)
	 
	End Function	
		</pre>
	</div>
</div>
	
	<h2 class="baslik">Koşullu Fonksiyonlar</h2>
<div class="konu">
<p>Bu bölümde bir veye birden çok koşul olması durumunda elimizdeki <span class=" keywordler">SUM</span>, <span class=" keywordler">COUNT </span>
gibi formüllerin koşullu türevlerini kullanmayı öğreneceğiz.</p>

<p>Türevi olan fonksiyonlar sözkonusu olduğunda bazen SUM'da olduğu gibi hem tek koşul(<span class=" keywordler">SUMIF</span>) hem çok koşul için(<span class=" keywordler">SUMIFS</span>) türev fonksiyon görebilirken, bazılarında ise sadece tek koşul türevi(<span class=" keywordler">AVERAGEIF</span> 
gibi) görülebilmektedir.(2016 versiyonunda AVERAGEIFS eklenmiştir.) Bununla 
birlikte bazılarının(MAX) ise tek koşullu biçimi(MAXIF) atlanarak doğrudan çok 
koşullu(MAXIFS) versiyonları sunulmuştur.</p>

<p>Bazılarında ise an itibarıyle türev fonksiyon hiç bulunmamaktadır. 
<span class=" keywordler">RANK</span> için ne <strong>RANKIF</strong> ne de <strong>RANKIFS</strong> mevcuttur. 
Bunlar için alternatif yöntemleri deneyeceğiz.</p>
	<p>Şimdi tek bir SUMIF ve SUMIFS örneğini burada inceleyelim. Diğerleri ise (<span class="keywordler">COUNTIF,COUNTIFS,MINIFS,MAXIFS,AVERAGEIF,AVERAGEIFS</span>) 
	bu ikisinin benzer kullanımına sahip olacak, onları bu
	<a href="../../../Ornek_dosyalar/Formuller/kosullufonk.xlsx">dosyayı</a> 
	indirip inceleyebilirsiniz.</p>
	<p>Tek koşullu foksiyonların genel syntaxı şöyle:<strong><em>XXX</em>IF(arama 
	alanı,aranan,hesaplanacak kolon)</strong></p>
	<p>Çok koşullularda ise en başta hesaplanacak kolon olur, sonrasında ise 
	"arama kolonu-aranan değer" çiftleri girilir. <strong><em>XXX</em>IFS(hesaplanacak 
	kolon,arama alanı1,aranan1,arama alanı2,aranan2,.....</strong>)</p>
	<p><span class="keywordler">COUNTIF </span>ve <span class="keywordler">COUNTIFS</span>'te hesaplama kolonu olmadığı için onları bu 
	genellemeden ayırabiliriz. Onun dışındaki kullanımları aynıdır, sadece 
	hesaplama kolonu seçilmez.</p>
	<p>Kolon parametreleri tüm kolon(A:A), belli bir alan(A2:A100), bir table 
	kolonu(Ör:Table1[Bölge]) veya bir Name(Ör:Bölgeler) olabilir.</p>
	<p><img src="/images/fonkistatistik8.jpg"></p>
	<pre class="formul">=SUMIF(Table1[Bölge];G3;Table1[Aylık Gerç])
=SUMIFS(D:D;A:A;F3;C:C;G3)</pre>
	<p>Yukarıda belirttiğim gibi diğer koşulu fonksiyonları indirdiğiniz dosyadan 
	inceleyebilrisiniz. Excel 2016da gelen fonksiyonları, daha eski sürüm 
	Excel'i olanların kullanabilmesi için alternatif çözüm olarak dizi 
	formüllerini önereceğiz. Onlara
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">buradan</a> 
	bakabilirsiniz.</p>
	<p>Bunlar için ayrıca UDF de yazılabilir. Aşağıdaki özelif isimli UDF'imi 
	görebilirsiniz.</p>
	<h4 class="baslik">Özel if</h4>
	<div>
	<pre class="brush:vb">
Function özelif(İşlem As String, BakılacakAlan As Range, kriter As Variant, İşlemAlan As Range)
'işlem yerine, MAX, MIN, MEDIAN gibi aggreagete fonksiyonları yazılır
On Error GoTo hata
Dim aynıwb As Boolean, aynıws As Boolean
Dim strx As String
 
aynıwb = IIf(BakılacakAlan.Parent.Parent.Name = ActiveWorkbook.Name, True, False)
aynıws = IIf(BakılacakAlan.Parent.Name = ActiveSheet.Name, True, False)
 
 
If aynıwb Then
    If aynıws Then
        If IsNumeric(kriter) Then
            strx = İşlem & "(IF(" & BakılacakAlan.Address & "=" & kriter & "," & İşlemAlan.Address & "))"
        Else
            strx = İşlem & "(IF(" & BakılacakAlan.Address & "=""" & kriter & """," & İşlemAlan.Address & "))"
        End If
        'strx = "SumProduct(Max((" & BakılacakAlan.Address & "=""" & kriter & """)*(" & MaxAlan.Address & ")))"
    Else
        If IsNumeric(kriter) Then
            strx = İşlem & "(IF(" & "'" & BakılacakAlan.Parent.Name & "'!" & BakılacakAlan.Address & "=" & kriter & "," & "'" & BakılacakAlan.Parent.Name & "'!" & İşlemAlan.Address & "))"
        Else
            strx = İşlem & "(IF(" & "'" & BakılacakAlan.Parent.Name & "'!" & BakılacakAlan.Address & "=""" & kriter & """," & "'" & BakılacakAlan.Parent.Name & "'!" & İşlemAlan.Address & "))"
        End If
    End If
Else
    If IsNumeric(kriter) Then
        strx = İşlem & "(IF(" & "'[" & BakılacakAlan.Parent.Parent.Name & "]" & BakılacakAlan.Parent.Name & "'!" & BakılacakAlan.Address & "=" & kriter & "," & "'[" & BakılacakAlan.Parent.Parent.Name & "]" & BakılacakAlan.Parent.Name & "'!" & İşlemAlan.Address & "))"
    Else
        strx = İşlem & "(IF(" & "'[" & BakılacakAlan.Parent.Parent.Name & "]" & BakılacakAlan.Parent.Name & "'!" & BakılacakAlan.Address & "=""" & kriter & """," & "'[" & BakılacakAlan.Parent.Parent.Name & "]" & BakılacakAlan.Parent.Name & "'!" & İşlemAlan.Address & "))"
    End If
End If
 
özelif = Evaluate(strx)
Exit Function
 
hata:
özelif = "hata"
End Function	
	</pre>
	<p>Kullanımı da aşağıdaki gibi olup, MAXIF görevi görmektedir. A2:A10'da D1'deki değeri arayıp, 
	eşleşen satırlar için B2:B10'daki MAX değeri döndürür. Excel 2007de SUMIF ve AVERAGEIF geldiği için 
	sadece MAXIF ve MINIF amacıyla kullanılabilir.(Excel 2016'da MAXIFS ve MINIFS geldiği için gereksiz hale gelmiştir.)</p>
	<pre class="formul">=özelif("MAX";A2:A10;D1;B2:B10)</pre>
	</div>
	<h4>Kombine kriterler</h4>
	<p>Bazı kriterler sabit bir rakam(veya hücre) olmaktan ziyade &lt; veya &gt; 
	işaretlerinden oluşur. Böyle durumlarda kriter tırnak içine alınır. Eğer 
	işaretten sonraki değer bir hücreden gelecekse bu da &amp; işaretiyle 
	birleştirilir.</p>
	<pre class="formul">=SUMIFS(A:A;B:B;K2;C:C;"&gt;="&amp;L2)
=COUNTIFS(A:A;K2;C:C;"&lt;10")</pre>
	<h4>COUNTIFS'in gruplu sıralamadaki(RANKIFS amaçlı) kullanımı</h4>
	<h5>RANK</h5>
	<p>Öncelikle <span class="keywordler">RANK </span>fonsiyouna bakalım. RANK 
	fonksiyonu, belirli bir gruptaki sayıların büyüklük sırasını verir. 
	Syntax'ı şöyledir. RANK(sayı, hücre grubu;[sıralama yönü])</p>
	<p>Sıralama yönü parametresi verilmez veya 0 girilirse büyükten küçüğe göre 
	sırayı verir, yani en büyük rakamın sırası 1 olur. Bu parametre 1 girilirse 
	en küçük rakamın sırası 1 olur.</p>
	<p>NOT:Bu fonksiyon 2010'da yerini <span class="keywordler">RANK.EQ</span>'ya bırakmıştır ancak geriye dönük 
	uyumluluk nedeniyle hala kullanılmaktadır. Gördüğünüz gibi hem RANK hem 
	RANK.EQ aynı değere sahip sayılara aynı sırayı verir. 2010'la bir de
	<span class="keywordler">RANK.AVG</span> 
	geldi, bu ise aynı değere sahiplere buçuklu bir sıra verir.</p>
	<p><img src="/images/fonkistatistik9.jpg"></p>
	<h5>RANKIFS</h5>
	<p>İlginçtir ki, RANK'ın yukardaki diğer fonksiyonlar gibi koşullu versiyonu 
	yoktur. Ancak biz alternatif yöntemlerle bu amaca hizmet eden formüller 
	yazabiliyoruz. Şimdi yukarıdaki görüntüde E kolonuna şu formülü yazdığımızda 
	bu bize, her bölgenin her üründeki şube sırasını verecektir.</p>
	<pre class="formul">=COUNTIFS(A:A;A2;C:C;C2;D:D;"&gt;"&amp;D2)+1</pre>
	<p>Formülün mantığı şöyle: Başkent bölgesinde Ürün1'de kendisinin değerinden 
	büyük kaç şube var. En büyük değer için bu 0 çıkacaktır, çünkü ondan daha 
	büyük rakamı olan şube yoktur. Sonuca 1 ekleyerek nihai sırayı buluyoruz.</p>
	<p>Bir diğer yöntem de aşağıdaki gibi bir UDF yazmak olacaktır.</p>
	<h4 class="baslik">Rankif için tıklayın</h4>
	<div>
	<pre class="brush:vb">
Function rankifs(ParamArray kriterler())
On Error GoTo hata
    Dim i As Integer
    Dim formül As String
    Dim str As String
   
    If (UBound(kriterler) + 1) Mod 2 <> 0 Then
        rankifs = "Eksik Parametre. Kriter ve alan sayıları çiftler halinde girilmelidir"
        Exit Function
    End If
   
    For i = LBound(kriterler) To UBound(kriterler)
        If i = UBound(kriterler) Then
            formül = formül & """>""" & "&" & kriterler(i)
        Else
            If i Mod 2 = 0 Then
                formül = formül & kriterler(i).Address & ","
            Else
                formül = formül & """" & kriterler(i) & """" & ","
            End If
            'formül = formül & IIf(i Mod 2 = 0, kriterler(i).Address, """" & kriterler(i) & """") & ";" 
        End If
    Next i
   
    str = "=COUNTIFS(" & formül & ")"
    rankifs = Evaluate(str) + 1
    Exit Function
   
hata:
rankifs = Err.Description
End Function	
	</pre>
	</div>

	<h4>Aynı kolon için çok kriter belirtmek</h4>
	<p>Koşullu kriterleri genelde tek bir kolonda tek bir değeri aramak için 
	kullanırız ancak bazen birden çok değeri de aramamız gereken durumlar 
	olacaktır. Bunun için bu konuda göstereceğimiz iki yöntem olacak, diğer 
	yöntemleri ise
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllleri</a> sayfasında göreceğiz.</p>
	<p>İlk yöntemimiz ilgili kriter kadar SUMIF'i yanyana yazıp toplamaktır. Bu 
	yöntem çok sayıda kriter olması durumunda çok pratik olmayacaktır. İkinci 
	yöntem ise SUM ve SUMIF'i beraber kullanmaktır. Kriterler süslü parantez 
	içine yazılır.</p>
	<pre class="formul">=SUMIF(Table1[Bölge];"Başkent 1";Table1[Aylık Gerç])+SUMIF(Table1[Bölge];"Başkent 2";Table1[Aylık Gerç])
=SUM(SUMIF(Table1[Bölge];{"Başkent 1";"Başkent 2"};Table1[Aylık Gerç]))</pre>
</div>


<h2 class="baslik">İstatistiki Fonksiyonlar</h2>
<div class="konu">
	<p>Aşağıda bu fonksiyonların nerede ve ne zaman 
	kullanabilecğeinize ait bilgiler bulunmakta olup, detay kullanımlarını bu
	<a href="../../../Ornek_dosyalar/Formuller/istatistikfonk.xlsx">dosyayı</a> 
	indirirek görebilirsiniz. Bunları ayrıca Data menüsündeki
	<a href="DataMenusu_VeriAnalizi.aspx">Data Analysis</a> eklentisi ile de 
	formülsüz şekilde elde edebilirsiniz.</p>
	<p><span class="keywordler">MEDIAN</span>:Bir grup sayının en ortasındaki 
	değeri verir. Ortalamayı aşırı saptıran değerlerin olduğu bir kümede AVERAGE 
	ile ortalama almak yerine MEDIAN da kullanılabilir. Ortalamayı saptıran 
	değerlerin sayısı 1-2'den fazla değilse yukarıdaki gibi "uç değerler hariç 
	ortalama"yı hesaplayan formül de kullanılabilir ama MEDIAN kullanımı daha 
	basittir.</p>
	<p><span class="keywordler">STDEV</span>:Standart sapmayı verir. Bunun 
	çeşitli türevleri var, eminim incelemeye değerdir ancak ben burada kendi 
	dünyamda kullandığım örneği vereceğim. Ben, standart sapmanın kendisinden 
	ziyade bunun ortalamaya bölümünü kullanıyorum. Bu oran bize bir gruptaki 
	dalgalanma/sapma 
	oranını verecektir.</p>
	<p>Diyelim bir banka şubesindeki 4 satış temsilcisindeki müşteri sayıları 
	sırayla 880, 770, 910 ve 840 olsun. Bunların müşteri sayılarında aşırı bi 
	dalglanma olmadığını söylemek hemen mümkündür, bunun matematiksel ifadesi 
	ise şöyledir.</p>
	<pre class="formul">=STDEV.P(B2:E2)/AVERAGE(B2:E2)</pre>
	<p>Tüm listeye bu formülü uyguladığımızda tek çırpıda sırıtanları hemen 
	görebiliriz. Burada sabit bir sınır olmamakla birlikte duruma göre %20/%30 
	un üstündekilere sapma var gözüyle bakabilirsiniz. Bu tamamen sizin 
	hassasiyet derecenize bağlı. Ör:"Ben her ay 3 şubede yeniden müşteri 
	dağıtımı yaparım" derseniz bu örnekte Şube2, Şube6 ve belki bir de Şube12de 
	işlem yapabilirsiniz.</p>
	<p><img src="/images/fonkistatistik11.jpg"></p>
	<p><span class="keywordler">MODE/FREQUENCY</span>:MODE, bize belirli bir 
	grupta en çok tekrar eden elemanı verirken, FREQUENCY belirli bir aralıktaki 
	değerlerin ilgili kümede kaç kez geçtiğini söyler. MODE'la bulunan değerin 
	kaç kez geçtiğini ise COUNTIF içine sarmalayarak bulabilirsiniz.</p>
	<p>En çok tekrar eden değerin birden fazla olduğu durumlar için MODE.MULT 
	fonksiyonu devreye sokulmuştur. MODE.SNGL ise MODE'un yerini almıştır, ancak 
	geriye uyumluluk adına MODE da hala varlığını korumaktadır.</p>
	<p>A2:A18 arasındaki değerlerin 
	15;15;17;19;20;22;23;23;24;25;25;25;25;26;26;26;29;32 olduğu bir örnekte 
	formüllerimiz şöyledir</p>
	<pre class="formul">=MODE.SNGL(A2:A18)
=COUNTIF(A:A;MODE.SNGL(A:A)) //kaç kez geçiyor</pre>
	<p>Belirli aralıklardaki baremlerden oluşan bir grupta frekans değerlerini 
	elde etmek için FREQUENCY'yi
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formülü</a> şeklinde uygularız. Bu formül bizi uzun bir COUNTIFS yazmaktan 
	kurtarmaktadır.(Bütün bunları
	<a href="DataMenusu_VeriAnalizi.aspx">Data Analysis</a> aracındaki histogram 
	ile de yapabiliyoruz)</p>
	<p>Aşağıdaki Frequency kolonundaki formül ile onun yanındaki kolonun en 
	altındaki commentboxlı hücredeki formül aşağıdaki gibidir. Detaylara örnek 
	dosyadan ulaşabilirsiniz.</p>
	<pre class="formul">{=FREQUENCY($A$2:$A$19;D8:D11)}
=COUNTIFS(A:A;"&lt;="&amp;D12;A:A;"&gt;="&amp;C12)</pre>
	<p><img src="/images/fonkistatistik12.jpg"></p>
	<p>FREQUENCY fonksiyonu dizi formülü şeklinde kullanıldığında bir değerin 
	sonraki tekrarlarını 0 adet gösterir. Bunun detaylarına
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllerinde </a>değineceğiz.</p>
	<p>NOT:MODE fonksiyonu sadece sayılar için çalışmaktadır. Metinsel ifadelerden 
	hangisinin en çok ve kaç kez geçtiğini bulmak için yine
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllerini</a> kullanırız.(Bu dizi formülleri de artık çok olmaya 
	başladı!!)</p>
	<p><span class="keywordler">NORM.DIST</span>:Bir kümedeki verilerin normal 
	dağılım gösterip göstermediğni belirlemenin yollarından biri, dağılımın 
	grafiğini çizmek ve bunu yorumlamaktır. Dağılım grafiğini çizmek için 
	bunları çizime uygun hale getirmek ve sonra Scatter türündeki grafik ile 
	çizmek gerekir.</p>
	<p>Aşağıdaki gibi, şubelere verilen hedeflerin normal dağılıma uygun olup 
	olmadığını görmek için C kolonunda şu formülü yazarak grafik değerlerini 
	elde ettik. Gördüğünüügibi grafik değerelerinde kümenin ortalama ve standart 
	sapma değerlerini de kullanıyoruz.</p>
	<pre class="formul">=NORM.DIST(B2;$E$2;$F$2;FALSE)</pre>
	<p><img src="/images/fonkistatistik13.jpg"></p>
	<p><span class="keywordler">KURT/SKEW</span>:Bir kümedeki verilerin normal 
	dağılım gösterip göstermediğini belirlemenin bir diğer yolu da basıklık ve 
	çarpıklık katsayılarına bakmaktır. Çarpıklık(skewness) normal dağılımda 
	0'dır. Sonuç negatifse dağılım sağa çarğık, pozitifse sola çarpıktır. 
	Basıklık(kurtosis) normal dağlımda 0'dır; negatifse basık bir dağılım, 
	pozitifse sivri bir dağılım sözkonusudur.</p>
	<p>Dağılımın normal dağılımdan manidar düzeyde farklılaşmadığını söylemek 
	için sonuçların iki fonkisyon için de -1 ve +1 arasında olması 
	gerekmektedir.&nbsp;Bu değerleri yine
	<a href="DataMenusu_VeriAnalizi.aspx">Data Analysis</a> eklentisinden 
	topluca görebiliyoruz.</p>
	<p>Yukardaki normal dağılım örneğindeki kümeye uyguladığımızda bu sonuçların 
	da normal dağılıma uyduğunu görmüş bulunuyoruz.</p>
	<pre class="formul">=KURT(B2:B46) //-0,63
=SKEW(B2:B46) /-0,08</pre>
	<p><span class="keywordler">CORREL</span>: İki veri kümesinin birbiriyle 
	ilintili olup olmadığını gösterir. Örnek olarak, banka şubelerinin 
	hedeflemesinde kullanılacak input listesine bakabiliriz. Mesela Ürün1 
	kalemine ait bir hedefleme yapacağız diyelim, bu kalemin inputu olarak 
	elimizdeki diğer ürünlerden hangilerini kullanabiliriz sorusunun cevabını 
	arıyoruz. Aşağıdaki görüleceği üzere Ürün1 ve Ürün2 oldukça korele yani 
	ilintili. Yani bir şube ne kadar Ürün2 satıyorsa Ürün1'i satma oranı, diğer 
	şubelerde de benzerlik gösteriyor. Ürün3 ve Ürün4'ün ise Ürün1 rakamlarıyla 
	tüm şubelerde benzer bir orana sahip olmadığını yani birbirleriyle alakalı 
	olmadığını görüyoruz. Dolayısıyla hedeflemeye input olarak sadce Ürün1 dahil 
	edilmelidir.</p>
	<pre class="formul">=CORREL($B$2:$B$18;C2:C18)</pre>
	<p><img src="/images/fonkistatistik14.jpg"></p>
	<h4>TAHMİNLEMELER:</h4>
	<p>Excel bize oldukça fazla sayıda istatistiki tahminleme aracı sunuyor ama 
	biz burada kısaca iki tanesine bakacağız. Yine tabiki ilgilenenler
	<a href="DataMenusu_VeriAnalizi.aspx">Data Analysis</a> eklentisini 
	inceleyebilirler. Gerçi bu linkte de çok yeterli bilgil bulamayabilirsiniz, 
	ben bu sitenin kapsamında olmadığı için çok detayalara girmiyorum, arzu 
	edenler başka kaynaklara da bakabilir.</p>
	<p><span class="keywordler">FORECAST</span>(yeni x, bilinen y'ler, bilinen 
	x'ler): Geçmiş dönem satışlarını bidiğiniz durumlarda önümüzdeki dönem 
	satışlarının ne olacağını tahmin etmek istediğiniz dönemlerde bu fonksiyonu 
	kullanırız. </p>
	<pre class="formul">=FORECAST(E$1;$B$2:$B$18;$A$2:$A$18)</pre>

	<p>Bunun bir de sezonsallık etkisini dikkate aldığı versiyonu da var. O da 
	şöyle olup detayı örnek dosyada inceleyebilirsiniz. </p>
	<pre class="formul">=FORECAST.ETS(E$1;$B$2:$B$18;$A$2:$A$18)</pre>
	<p>Bir de benzer sonucu veeren ve kullanımı daha kolay ola
	<span class="keywordler">TREND </span>fonsiyonu var. FORECAST'tan farkı, 
	bunu dizi formülü olarak da kullanbilmeniz. Ancak bunun yerine FORECAST'ı 
	sağa doğru kaydırarak da ayn sonucu elde edebilirsiniz. Başka farkları da 
	var ancak bunun için istatistiki terminiolojiye daha aşina olmak gerekir.</p>
	<pre class="formul">
=TREND(B2:B18;A2:A18;E1:Q1) //dizi formülü olarak girilmeli
//veya
=TREND($B$2:$B$18)</pre>
</div>
	<h2 class="baslik">Diğer Fonksiyonlar</h2>
	<div class="konu">
	<p><span class="keywordler">ABS(sayı)</span>:Sayıların mutlak değerini 
	verir. </p>
	<pre class="formul">=ABS(-10) //10</pre>
	<p>Gelir ve giderin(negatif olarak gösterildiğini varsayalım) bir arada olduğu 
	bir veri kümesinde sıralama yapmak istediğinizde negatifler en alta gitmesin 
	diye önce bunların mutlak değerini almak isteyebilirsiniz.</p>
		<p><img src="/images/fonkistatistik15.jpg"></p>
		<p>Mutlak değer alıp sıralarsak durum şöyle olur.</p>
		<p><img src="/images/fonkistatistik16.jpg"></p>
		<p>Burdan da görülür ki dikkate alınacak en önemli kalem Faiz 
		Gideri.(Tabi listenin daha kalabalık olduğunu düşünürsek ABS'nin etkisi 
		daha anlaşılır olacaktır)&nbsp;</p>
		<p>ABS'nin bir diğer önemli kullanımı da başlangıç noktası negatif olan 
		bir noktadan pozitif olan bitiş noktasına olan değişim oranıdır. Mesela 
		bir şube geçen sene 100(A2) lira zarar ediyor ve bu sene 50(B2) lira kar 
		ediyorsa geçen seneye göre olan değişimi şöyle hesaplanır.</p>
		<pre class="formul">=(B2-A2)/ABS(A2) //%150
//klasik formüle hesaplasaydık aşağıdaki gibi hatalı olurdu
=(B2-A2)/A2 //-%150</pre>
		<p><span class="keywordler">RAND(), RANDBETWEEN(ilk,son)</span>:RAND 0-1 
	arası rasgele sayı üretirken, RANDBETWEEN iki sayı arasında rasgele bir sayı üretir. 
	RAND özellikle rasgele olasılık değeri üretmekte kulanılabilir. 
	RANDBETWEEN'i ise genelde örnek sayı kümeleri oluşturmakta kullanıyorum. 
	Mesela bu siteyi hazırlarken oluşturduğum örnek dosyalardaki rakamlarda hep bu 
	fonksiyonun izi var.</p>
	<pre class="formul">=RAND() //Mesela 0,25734
=RANDBETWEEN(10;20) //Mesela 12</pre>
	<p><span class="keywordler">SIGN(sayı)</span>:Sayı pozitifse 1, negatifse 
	-1, 0'sa 0 üretir. Bunun da dizi formüllerinde güzel bir kullanım örneği 
	var. Yer yer belirttiğim gibi hiçbirşey için özellikle bu sitede anlatılan 
	konular için "benim işime yaramaz" demeyin. Ben zaten buraya 20 yıl boyunca 
	bir şekilde kullandığım şeyleri koydum. SIGN da bunlardan biri.</p>

</div>



</asp:Content>
