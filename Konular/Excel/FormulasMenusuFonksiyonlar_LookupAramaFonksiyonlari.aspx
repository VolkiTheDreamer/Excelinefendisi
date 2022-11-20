<%@ Page Title='FormulasMenusu1 LookupAramaFormulleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='5'></asp:Label></td></tr></table></div>


<h1>Lookup(Arama) Fonksiyonları</h1>
	<p>Benim şahsi gözlem ve tecrübeme göre arama/bulma/referans grubundaki fonksiyonlar az bilinen, bilinse dahi az kullanılan, kullanılsa 
dahi 
tam kapasitede kullanılmayan fonksiyonlardandır. Ne yazık ki, bu durumda olan 
birçok Excel kullanıcısı işlerini çok daha uzun yollardan yapmaktadır. Bu da 
ilgili kurumda üretkenlik kaybına neden olmaktadır.</p>
	<p>Şimdi bu fonksiyonları (daha etkin) kullanarak 
	işimizi nasıl daha verimli yaparız, bunları göreceğiz. Sayfa boyunca 
	kullanılan örnekleri içeren dosyayı
	<a href="../../Ornek_dosyalar/Formuller/vlookup.xlsx">buradan</a> 
	indirebilirsiniz.</p>
	<p>Bu bölümde anlatılanların bir kısmına ait video eğitimini
	<a href="https://www.udemy.com/derinlemesine-excel-vlookup-ve-index-match-fonksiyonlar/learn/v4/overview">
	Udemy sayfamda</a> bulabilirsiniz.</p>
    <p>EDIT(2021): Office 365 ile gelen <strong>XLOOKUP</strong> fonksiyonunu INDEX-MATCH fonksiyonlarından hemen sonraya koydum, ki hayatımızı nasıl kolaylaştırdığını daha iyi görün.</p>

<!--********************************************************************************************************************************************--><h2 class="baslik">
	VLOOKUP</h2>
<div class='konu'>
<h3>Giriş</h3>
<p><strong>Vlookup</strong>, bana göre ortalama bir Excel kullancısı için en önemli fonksiyondur. Gerçi zaman geçtikçe 
bunun da eksikliklerini görecek ve alternatif yollar aramanız gerekecek. Bu 
kısımada öncelikle Vlookup'ı daha etkin nasıl kullanırız, ona bakacağız; 
sonrasında da kendisinin yetersiz kaldığı durumlarda neler yapabileceğimize.</p>
	<p><span class=" keywordler">Vlookup(aranan,aramayeri,kaçıncıkolon,[eşleşmetipi])</span> şeklinde bir syntaxa sahip olan Vlookup'ta arama 
	işlemi "arama yeri"ndeki ilk kolonda yapılır. Aranan değer bulunduğununda, buna karşılık gelen esas istediğimiz sonuç, 
	ilgili alanın belirtilen kolonunda bulunur ve getirilir. Eğer aranan değer ilk kolonda bulunamazsa 
	formül sonucu #N/A döner.</p>

<h4>Arama alanı</h4>
<p>Arama alanı olarak, tüm kolon seçilebileceği gibi $B$1:$C$10 şeklinde sadece ilgili alanın seçimi şeklinde de olabilir. Ayrıca bir Table'a da lookup işlemi yapılabilir(Table detayları için 
<a href="HomeMenusu_Tablolar.aspx">buraya</a> bakabilirsiniz.) Yine Named Range kullanımı da okuma kolaylığı sağlayacaktır.</p>

<p>Aşağıda çeşitli vlookup işlemleri görülmektedir</p>
<img src="/images/excelvlookup1.jpg" class="zoomla">
<h4>Kolon indexi ve sağa doğru kaydırma</h4>
<p>Vlookup formülünüzü yazarken genelde sabit bir kolon indeksi gireriz. Formülümüzü kaydırarak kopayalasak bile bu indeks sabit olduğu için hep aynı kalır. 
Ancak, bu formül içeren hücreyi sağa doğru kaydırırken kolon indeksinin de artarak gitmesini istediğimiz durumlar olacaktır. Aşağıdaki tabloda olduğu gibi; 
en sağ hücredeyim ama indeks hala 2.</p>
	<p><img src="/images/excelvlookup2.jpg"></p>

	<p>Böyle durumlarda genelde arama alanının üstüne yardımcı bi satır açılır 
	ve kolon indeksi oradan alınır. </p>
	<p><img src="/images/excelvlookup3.jpg"></p>
	<p>Bu bir çözümdür ancak şık bir çözüm değildir. 
	Ana sayfada ne demiştik, işleri sadece doğru yapmak yetmez, hızlı ve zarif de yapmalıyız. 
	Gerçi&nbsp; bu yöntem hızlıdır, ama zarif değildir. Eğer acil ve/veya geçici 
	bir bilgiye ihtiyacınız varsa bu tür yöntemler kabul edilebilir. Ancak daha 
	kalıcı bir çözüm istiyorsak böyle yöntemler uygun değildir. Zira bu sayfayı 
	print almak istediğimizde bu yardımcı satır da basılır ve bu da çok hoş 
	olmaz.</p>
	<p>Böyle bir durumda COLUMN veya daha iyisi MATCH formülünden faydalanabiliriz.&nbsp;Önce 
	COLUMN yöntemine bakalım. (Bu fonksiyonu daha detaylıca aşağıda ele alacağız.)</p>
	<pre class="formul">
=VLOOKUP(aranan,alan,COLUMN()-/+x,0)</pre>

<p><img src="/images/excelvookup3.jpg"></p>
	<p>Bu örnekte x yerine birşey yazmak gerekmedi ancak bazen kolonun yerine 
	göre bir değer girmek gerekebiliyor. Column(M2)-5 gibi.</p>
	<p>İkinci çözüm MATCH fonksiyonu ile sağlanır. Üstelik bunda COLUMN'da 
	olduğunun aksine +/- bir diğer girme zorunluluğu yoktur.</p>
	<pre class="formul">
=VLOOKUP($A10,$A$1:$M$5,MATCH(B9;$B$1:$M$1;0),0)</pre>

	<p>Burda MATCH, Ocak ayının B1:M1 içindeki yerine bakıyor, kaçıncı kolonda 
	olduğunu bulup onu getiriyor, yani 2'yi. Formülü kaydırdığımızda da diğer 
	ayların kolon sırasını da ona göre getiriyor.</p>
	<p>MATCH'in COLUMN veya yardımcı satır yöntemine göre bir üstünlüğü de&nbsp; 
	kolonların yeri değişse bile doğru sonuç getirmesidir. Diğer iki yöntemde 
	ise kolonların yeri değişirse, mesela aylar alfabetik sıralı gelse veya 
	Aralıktan Ocak'a doğru sıralanmış olsa, formül hatalı sonuç getirir.</p>
	<p>Bir diğer alternatif de aşağıdaki gibi kayarak ilerlemeyi sağlayan bir UDF yazmaktır. 
	UDF açıklamlarını fonksiyonlar sayfasında detaylıca ele&nbsp; alacağımız 
	için burada tekrar yapmak istemedim.</p>
<h4 class="baslik">Süperlookup UDF'i için tıklayınız</h4>
<div class="konu">
<p>Kullanım örneği için 
<a href="../VBAMakro/Fonksiyonlar_ExcelicinUDFKullaniciTanimliFonksiyonlar.aspx#superlookup">buraya</a> tıklayınız.</p>

<pre class="brush:vb">
Function süperlookup(alan As Range, sütun As Range, aranan As Range)
'paramterlerin sırası klasik vlookupa göre farklıdır
On Error GoTo hata
    süperlookup = alan.Columns(1).Find(aranan, lookat:=xlWhole).Offset(0, sütun.Column - alan.Columns(1).Column).Value
    Exit Function
hata:
    süperlookup = "Bulunamadı"
End Function
</pre>
</div>

<h4>Eşleşme tipi</h4>
<p>Vlookup'ın son parametresi genellikle tam eşleşmeyi sağlamak için 0 (veya False) şeklinde girilir. Ancak bize tam eşleşme değil de en yakın eşleşme lazımsa bu değer 1 
(veya True) olarak girilebilir ya da bu parametrenin varsayılan değeri zaten 1 olduğu için hiç girilmeyebilir. En yakın eşleşme olarak kendisinden en küçük değere bakar. 
Böyle bir duruma ne zaman ihtiyacımız olur bunu birazdan göreceğiz.</p>

<p>Yanlız burada dikkat edilmesi gereken bir nokta var, o da eğer <strong>parametre olarak 1 kullanılacaksa arama alanının sıralı olmasına</strong> dikkat edilmelidir. 
Aksi halde aşağıdaki gibi beklenmeyen sonuçlar çıkabilir. Hatta işin pis tarafı, bazen çıkan sonuç bir hata değil, gayet normal beklenen değerler olabilir ama yanlıştır(Aşağıdaki 
örnekte 500 çıkması gerekirken 105 çıkması gibi).</p>
	<pre class="formul">=VLOOKUP(A2;D:E;2) //B2'deki formül</pre>
	<p><img src="/images/excelvlookup4.jpg"></p>
	<p>Bu yöntem şu şekilde işler. Listede en yukardan aramaya başlar, kendinden 
	küçük bir değer görünce bunu bi kenara yazar, sonra bi aşağı satıra bakar, 
	hala kendinden küçük ama bi öncekinden büyükse bu sefer yeni satırdaki 
	değeri alır, <strong>ta ki kendinden büyük bir değere denk gelinceye kadar</strong>. O anda 
	durur ve o ana kadar kendisinden küçük olup ona en yakın değer hangisiyse 
	onu baz alır. Bu yukardaki örnekte 2'yi arıyoruz, 2 aşağılarda kalmış. En 
	yukarda 1'i görür, sonraki değer 10 olup 2den büyük olduğu için aramayı durdurur. 
	Gördüğünüz gibi son parametrenin 1 olarak kullanımı tehlikelidir ve çok 
	dikkat gerektirir.</p>

	<h3 id="iciceif">İçiçe IF yerine Vlookup </h3>
<p>Vlookup'ın son parametresinin 1 olarak kullanılması tehlikelidir tehlikeli 
olmasına ama doğru kullanıldığında bizi bir sürü dertten kurtarır. Bunlardan 
biri de içiçe IF 
formülü yazmaktan kurtarması durumudur. Aşağıdaki örnek üzerinden gidelim. </p>
	<p><img src="/images/excelvlookup5.jpg"></p>

<pre class="formul">
//bu formül yerine
=IF(A2&lt;2,7;2,73;IF(A2&lt;2,85;2,877;IF(A2&lt;3;3,024;IF(A2&lt;3,15;3,171;IF(A2&lt;3,3;3,318;IF(A2&lt;3,45;3,465;IF(A2&lt;3,6;3,612;IF(A2&lt;3,75;3,759;IF(A2&lt;3,9;3,906;IF(A2&lt;4,05;4,053;4,054))))))))))

//bu formül çok daha kısa ve dolayısıyla verimlidir
=VLOOKUP(A2;F:G;2;1)
</pre>

<p>Dikkat ettiyseniz F kolonundaki arama alanı sıralıdır. Bunu bir önceki kısımda 
özellikle vurgulamıştık.</p>
	<h3>Arama alanında araya kolon ekleme</h3>
<p>Hali hazırda Vlookup uyguladığınız bir sayfada, arama alanında araya bir yere yeni kolon eklenirse 
ve bu yeni kolonun sağında kalan kolonlardan 
bir data çekiyorsanız bunlar patlar. Çünkü formülün içindeki kolon index numarası değişmez. </p>

<pre class="formul">=VLOOKUP(A2;E:F;2;0)</pre>
<p>Şimdi E:F arasına 3 kolon ekleyelim, E:F'nin E:I odluğunu ama 2'nin değişmediğini görüyoruz.</p>

<pre class="formul">=VLOOKUP(A2;E:I;2;0)</pre>
<p>Bunu engellemek için <span class=" keywordler">MATCH</span> fonksiyonunu 
formülün içine eklemekte fayda var. Formülü MATCH kullanarak yazalım:</p>

<pre class="formul">=VLOOKUP(A2;E:F;MATCH($B$1;$E$1:$F$1;0);0)</pre>
<p>Şimdi E:F arasına 3 kolon ekleyelim. Formül hala doğru çalışıyordur, zira MATCH'li formül otomatikman genişledi.</p>
<pre class="formul">=VLOOKUP(A2;E:I;MATCH($B$1;$E$1:$I$1;0);0)</pre>

<p>Bir diğer seçenek de INDEX-MATCH ikilisini kullanmaktır, ki bunu aşağıda göreceğiz.</p>

<h3>Hızlı Vlookup için optimizasyon çalışması</h3>
<p>Vlookup çok faydalı bir fonksiyon olmakla birlikte, özellikle büyük data kümelerinde kullanımı dikkat gerektirir. 
Bunun için bazı tüyolarımız olacak.</p>

<ul>
<li>Arama listesi mümkünse aynı dosyada, hatta mümkünse aynı sayfada olsun. 
Özellikle geçici bir lookup işlemi yapacaksanız geçici veriyi aynı sayfaya 
alabilir ve sorasında silersiniz.</li>
<li>Arama alanı olarak her ne kadar tüm kolon seçimi(B:E gibi) daha pratik olsa da performans kaygınız varsa sadece ilgili alanın seçimi 
(B2:E10 gibi) daha verimli olacaktır. 
Özellikle birkaç onbin satırdan fazla bir data kümesine lookup çekiyorsanız sadece ilgili alan eşleşmesi yapın.</li>
<li>Arama listesi sıralı olsun.</li>
<li>Aşağıdaki hızlı lookup formülünü kullanın</li>.
<p>Bu formülde önce aranan yeri sıralamamız gerekir. Sıralanmış listeye, yakın 
eşleşme moduyla baktığımızda eşleşme sağlanırsa bu eşleşme sonucunu getir,
 sağlanmazsa NA getir diyoruz. Tam eşleşme modu tüm listeye baktığı için çok daha uzun sürer.</p>
<pre class="formul">
=IF(VLOOKUP(A2;Sheet1!A:A;1;1)=A2;VLOOKUP(A2;Sheet1!A:B;2;1);NA())
</pre>
<li>Yukarıdaki formül karışık geldiyse aşağıdaki UDF olarak hazırlanmış fonksiyonu da kullanabilirsiniz. Ancak 
Excel'in yerel fonksiyonları kullanıldığı için üstteki yöntem daha hızlıdır. 
Aşağıda bi hız karşılaştırması var zaten, orda en hızlı yöntemin hemen üstteki 
formül olduğunu görürsünüz.
<h4 class="baslik">Hızlılookup UDF'i için tıklayınız</h4>
<div class="konu">
<pre class="brush:vb">
Function hızlılookup(aranan As Range, alan As Range, kolon As Integer)

If WorksheetFunction.VLookup(aranan, alan, 1, 1) = aranan Then
    hızlılookup = WorksheetFunction.VLookup(aranan, alan, kolon, 1)
Else
    hızlılookup = "NA"
End If
End Function

</pre></div>
</li></ul>

<p>Deneme olması adına 500bin satırlık sırasız bir listede çeşitli lookup işlemleri yaptım. Veriler şöyle:</p>

<table class="alterantelitable">
<th>Sıralı mı</th>
<th>Eşleşme Tipi</th>
<th>Süre</th>
<tr><td>Evet</td><td>Tam</td><td>%1e 2 sn'de geliyor</td></tr>
<tr><td><strong>Evet</strong></td><td><strong>Yakın</strong></td><td><strong>%100e 1 sn'de geliyor</strong></td></tr>
<tr><td>Evet</td><td>Yakın(UDF)</td><td>%1e 2,5 sn'de geliyor</td></tr>
<tr><td>Hayır</td><td>Tam</td><td>%1e 2 sn'de geliyor</td></tr>
<tr><td>Hayır</td><td>Yakın</td><td>Hatalı sonuç verir</td></tr>

</table>


<p>PC konfigürasyonuna göre hızların değişeceği aşikardır, ama oranlar üç aşağı 
beş yukarı aynı kalır.</p>
    <p><strong>NOT</strong>: VBA ile UDF yazmak yerine XLL ile UDF yazarak da performansı artırabilirsiniz. İlginizi çekerse <a href="../VSTO/ThirdPartyKutuphanler_ExcelDNA.aspx">buradan</a> bakabilirsiniz.</p>
	<h3>Eksiklikler</h3>
<p>Gördüğünüz ve/veya bildiğiniz üzere, Vlookup sola doğru arama işlemi yapmıyor. Her zaman pozitif bir değer veriyorsunuz ve sağa doğru arama yapıyorsunuz. 
Bu sorunu aşmak için ya <strong>INDEX-MATCH</strong> ikilisi kullanılır, ki bunları bir aşağıdaki bölümde göreceğiz, veya bir UDF tanımlamanız gerekir. Benim bunun için hazırladığım bir fonksiyon var. VBA bilenler bu kodu inceleyebilirler. </p>

<h4 class="baslik">Terslookup UDF'i için tıklayınız</h4>
<div class="konu">
<pre class="brush:vb">
Function terslookup(aranan As Variant, hedefalan As Range, kaçıncı_kolon)
Dim index1 As String, match2 As String
Dim aynıwb As Boolean, aynıws As Boolean

aynıwb = IIf(hedefalan.Parent.Parent.Name = ActiveWorkbook.Name, True, False)
aynıws = IIf(hedefalan.Parent.Name = ActiveSheet.Name, True, False)


If aynıwb Then
    If aynıws Then
        match2 = hedefalan.Columns(hedefalan.Columns.Count).Address
        index1 = hedefalan.Columns(hedefalan.Columns.Count + 1 - Abs(kaçıncı_kolon)).Address 'niye abs, olur da negatif girmezler diye önce
    Else
        match2 = "'" & hedefalan.Parent.Name & "'!" & hedefalan.Columns(hedefalan.Columns.Count).Address
        index1 = "'" & hedefalan.Parent.Name & "'!" & hedefalan.Columns(hedefalan.Columns.Count + 1 - Abs(kaçıncı_kolon)).Address
    End If
Else
    match2 = "'[" & hedefalan.Parent.Parent.Name & "]" & hedefalan.Parent.Name & "'!" & hedefalan.Columns(hedefalan.Columns.Count).Address
    index1 = "'[" & hedefalan.Parent.Parent.Name & "]" & hedefalan.Parent.Name & "'!" & hedefalan.Columns(hedefalan.Columns.Count + 1 - Abs(kaçıncı_kolon)).Address
End If


If IsNumeric(aranan) Then
    strx = "INDEX(" & index1 & ",Match(" & aranan & ", " & match2 & ", 0))"
Else
    strx = "INDEX(" & index1 & ",Match(""" & aranan & """, " & match2 & ", 0))"
End If
terslookup = Evaluate(strx)

    
End Function
</pre>
</div>

<p><strong>İkinci</strong> olarak, çoklu kritere göre Vlookup yapmak için yardımıcı kolon gerekiyor, bu da tek seferde hedefe ulaşmamızı engelleyen 
ve şık olmayan bir yöntem anlamına geliyor. Bunun üstesinden zarif ve etkin bir 
biçimde gelmek için yine <strong>INDEX-MATCH</strong> ikilisini kullanırız.</p>
	<p>Zarif olmayan yönteme bi bakalım, aşağıda INDEX-MATCH yöntemiyle siz 
	bilahare karşılaştırma yaparsınız.</p>
	<p>Aşağıdaki örnekte Bölge1'in Toplam ürün satışına göre 1. şubesinin Ürün1 rakamını getirelim.</p>
	<p><img src="/images/excelindex3.jpg"></p>
	<p>Önce N kolonuna yardımcı kolonu açalım,</p>
	<p><img src="/images/formulavlookupcokkriterhelper.jpg"></p>
	<p>Formülümüz şöyle olacaktır:</p>
	<pre class="formul">=VLOOKUP("Bölge1_1";N:R;2;0) //Aranan değeri elle girdim ama bunu bir hücreden de okutabilirdik</pre>
<p>Dediğim gibi bu yöntem şık değildir, üstelik tablo yapınızı değiştirdiği için 
tehlikeli de olabilir, zira başka yerlerdeki formülleriniz de bu tablodan data 
çekiyorsa sıkıntılar oluşabilir.</p>
	<p><strong>Üçüncü</strong> olarak Vlookup her zaman ilk gördüğü eşleşme sonucunu getirir, ancak siz özellikle ilk eşleşme değil de sonraki eşleşme 
veya tüm eşlemelerin sonucunu yanyana görmek isterseniz başka birşeyler yapmanız gerekir. 
	Bunu da aşağıdaki kısımlarda çok eşleşmeli lookup başlığı altında 
	görebilirsiniz.</p>

<h3>Özet</h3>
<ul>
<li>Vlookup her zaman sağa doğru arama yapar. Sola doğru arama için <strong>Index-Match</strong> kullanılmalı.</li>
<li>Vlookup her zaman tek kritere göre arama yapar. Çok kolon için ya <strong>Index-Match</strong> kulanılmalı, ya da yardımcı kolonda birleştirme yapılmalı(önerilmez) </li>
<li>Vlookup her zaman ilk değeri getirir, ikinci değer için yardımcı kolon kullanmak gerekir. Bu yardımcı kolonda da aşağı kayan bir COUNTIF kullanılır. 
Bu yöntem zahmetlidir. Aşağıda gösterilecek olan UDF kullanımı çok daha 
basittir.</li>
<li>Vlookup küçük büyük harf ayrımı gözetmez.</li>
<li>Vlookup'ın tam ve yaklaşık olmak üzere iki arama şekli vardır.</li>
<li>Vlookup joker karekterlerini kullanmaya izin verir.</li>
</ul>


</div>

<!--********************************************************************************************************************************************--><h2 class="baslik">
	HLOOKUP VE LOOKUP</h2>
<div class='konu'>
<h3>HLOOKUP</h3>
<p>Eğer lookup işlemini kolonda değil de satırda yapma durumu varsa o zaman HLOOKUP kullanırız. VLOOKUP'ta ilk kolona bakılırken bunda ilk satıra bakılarak lookup işlemi yapılır. 
Genel mantık Vlookup ile aynıdır.</p>
	<p><img src="/images/excellookup7.jpg"></p>

<pre class="formul">=HLOOKUP(B5;$A$1:$H$2;2;0)</pre>

<h3>LOOKUP</h3>
<p>VLOOKUP ve HLOOKUP'ın hem satırda hem sütunda çalışan bir formudur. Diğer ikisine göre eksikliği 
<strong>sadece sıralı datada çalışıyor </strong>olması. 
Artıları ise daha fazla. İlave bir parametre girmeden, bulabilyorsa tam eşleşmeyi yapar, bulamazsa aranan değerden en küçük değeri getirir. Ayrıca vlookup sadece sağa, hlookup sadece aşağı giderken LOOKUP ise her yöne doğru arama yapabilmektedir. </p>

<p>İki formu vardır; array ve vektör. Birçok yerde sadece vektör formu gösterildiği için 
ben de buna sadık kalıyorum.</p>

<pre class="formul">
=LOOKUP(C20;$B$16:$B$18) //varmı diye bakar, varsa kendisini getirir, yoksa N/A
=LOOKUP(C20;$B$16:$B$18;$C$16:$C$18) //varsa belirtilen kolondaki eşleşen veya en yakın datayı getirir
</pre>

<p>Bu arada bir dezavantaj da şudur ki, çok kolonlu lookup işlemi yapılamaz. Bunun için 
yine meşhur <strong>INDEX-MATCH</strong> kombinasyonu kullanılmalıdır.</p>

</div>



<!--********************************************************************************************************************************************--><h2 class="baslik">
	MATCH, INDEX, OFFSET</h2>
<div class="konu">
<h3>MATCH</h3>
	<p>Bir veri dizisi içinde bir elemanın kaçıncı olduğunu 
	MATCH fonksiyonu ile 
buluruz. Mesela aşağıdaki tabloda hem kolonda(F1:H1) hem de satırda(E2:E4) 
dizilmiş a,b,c değerleri var. Bunların sırasını buluyoruz.</p>
	<p><img src="/images/excelmatch.jpg"></p>
	<pre class="formul">=MATCH(A2;$E$2:$E$4;0) //Satır
=MATCH(A2;$F$1:$H$1;0) //Sütun</pre>
<p>Match, illa bir hücre dizisiyle kullanılmak zorunda değil. { } içindeki 
dizilerle de kullanılabilir. Diziler ve dizi formülleri için
<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">buraya</a> 
bakabilirsiniz.</p>
	<pre class="formul">=MATCH(2;{5;7;2;6;4};0) //3 döner</pre>
	<p>Son parametresi eşleşme tipini gösterir ve genelde tam eşleşmeyi temsilen 
	0/False girilir. 1 girilirse ve tam eşleşme yoksa aranandan küçük, -1 
	girilise aranandan büyük değeri getirir.</p>
	<p>MATCH'in dizilerle olan bu kullanımı, onu
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllerinin</a> aranan fonksiyonu yapmaktadır.</p>
	<h3>INDEX</h3>
	<p>Index'in olayı, bir veri dizesinde belirtilen indeksteki elemanı 
	getirir. Eğer tek parametre varsa, tek kolonluk bir veri vardır ve satırda 
	arama yapar, diğer durumlarda satır ve sütun kesişimine bakar. Yani bu 
	fonksiyonla aklınızda kalması gereken en önemli kelime <strong>kesişim</strong>dir.</p>
	<p><img src="/images/excelindex1.jpg"></p>
	<pre class="formul">=INDEX(B2:D9;2;3) //2.satır 3.kolon--&gt;486
=INDEX(B2:B9;2) //2.satır--&gt;1915</pre>
	<p>Formülümüzü illa data kümesini seçerek değil, seçime başlıkları da dahil 
	ederek oluşturabiliriz.</p>
	<pre class="formul">=INDEX(A1:D9;3;4) //486</pre>
	<h3>INDEX-MATCH</h3>
	<p>Index ve Match bir arada kullanıldığında etkisi müthiş olur. Bu şekilde 
	hem Vlookup'ın alternatifidir, hem de Vlookup'ın yapamadığı sola gitme ve 
	çoklu kritere göre lookup yapma imkanı verir.</p>
	<h4>Sola giden vlookup</h4>
	<p>Aşağıdaki örnekte Şube2'nin bölgesini bulalım.</p>
	<p><img src="/images/excelindex2.jpg"></p>
	<pre class="formul">=INDEX(J1:J12;MATCH("Şube2";K1:K12;0))</pre>
	<p>Formülün yaptığı iş şu:Önce MATCH ile, Şube2'nin K1:K12 içindeki sırasını 
	buluyoruz. Sonra bu sıra numarasını ,INDEX ile J1:J12 içinde arıyoruz.</p>
	<h4>Çok kriterli vlookup</h4>
	<p>Aşağıdaki örnekte Bölge2'nin Toplam ürün satışıan göre 1.şubesinin adını 
	ve Ürün1 rakamını getirelim.</p>
	<p><img src="/images/excelindex3.jpg"></p>
	<pre class="formul">=INDEX(K:K;MATCH("Bölge2"&amp;1;J:J&amp;M:M;0)) //Şube6</pre>
	<p>Çok kriterli aramalarımızda kullandığımız formül, bir dizi formülü olmak durumunda. Zira Match ile iki ayrı 
	değer(önce Bölge değerini sonra da sıra değerini) arıyoruz. O yüzden formülü 
	bitirince normal Enter yapmak yerine <strong>Ctrl+Shift+Enter</strong> yapıyoruz. 
	Bunun detaylarını
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formülleri </a>sayfasında göreceğiz. Formülü doğru girdiğinizden emin olmak 
	için formülün başında ve sonunda süslü parantezleri "{ ve }" görüp görmediğinize 
	bakabilirsiniz.</p>
	<pre class="formul">{=INDEX(K:K;MATCH("Bölge2"&amp;1;J:J&amp;M:M;0))}</pre>
	<h5>Büyük listelerde çok kriterli lookup</h5>
	<p>Yanlız veri kümemiz çok büyükse ne yardımcı kolonlar açmak ne de INDEX/MATCH kullanmak işe yarar. 
	Belki yardımcı kolonlar aracılığı ile yukarıda gördüğümüz optimizasyon 
	yöntemini kullanmak çzüm olabilir. Ancak bu çözümün de uygulanmadığı 
	durumlarda Access'e başvurmanız gerekebilir. Bununla ilgili örnek bir 
	videoyu <span class="GridAltItem">
	<a href="https://www.udemy.com/derinlemesine-excel-vlookup-ve-index-match-fonksiyonlar/learn/v4/overview">
	Udemy'deki ücretsiz eğitimimde</a></span> anlattım(10.video: Büyük veri 
	kümeleriyle çalışmak).</p>
<h3 id="cokeslesmelilookup">Çok Eşleşmeli Vlookup</h3>
<p>Bazen, lookup yaptığımız alanda birden fazla eşleşme varsa hepsinin sonucunun 
gelmesini isteriz. Normal Excel fonksiyonları ile bunu yapmak çok olası değil. 
Aslında oldukça dolambaçlı yollardan bunu yapanlar var ama ben buna girmek 
istemiyorum. Bunun yerine aşağıdaki UDF'i kullanmak çok daha kolaydır.</p>
	<pre class="brush:vb">Function çok_eşleşmeli_vlookup(aranan As Range, alan As Range, kaçıncıkolon As Integer)
Dim a As Range
Dim dict As New Scripting.Dictionary 'Reference olarak eklenmiş olmalı, ekli değilse Late binding olarak yaratılabilir
 
If alan.Columns(1).Rows.Count = 1048576 Then
    Set alan = Range(alan(1, 1), alan(1, 1).End(xlDown))
End If
 
    For Each a In alan.Resize(, 1)
        If Not dict.Exists(a.Value) Then
            dict.Add a.Value, a.Offset(0, kaçıncıkolon - 1).Value
        Else
            geçici = dict(a.Value)
            dict.Remove (a.Value)
            dict.Add a.Value, geçici &amp; ";" &amp; a.Offset(0, kaçıncıkolon - 1).Value 'x
        End If
    Next a
 
çok_eşleşmeli_vlookup = dict(aranan.Value)
End Function</pre>
	<p>NOT: Yukarıda sonunda x bulunan satırı kaldırıp aşağıdaki 2 satırı 
	eklersek, eşleşen değerleri bize distinct(benzersiz) olarak getirir.</p>
	<pre class="brush:vb"> x = a.Offset(0, kaçıncıkolon - 1).Value
dict.Add a.Value, geçici &amp; IIf(InStr(1, geçici, x, vbTextCompare) &gt; 0, "", ";" &amp; x)</pre>
	<p>Aslında kodu daha kullanışlı hale getirebiliriz, bunun için bi parametre daha ekleyelim.</p>
	<pre class="brush:vb" id="cokeslesmelilookup_pre">
Function çok_eşleşmeli_vlookup(aranan As Range, alan As Range, kaçıncıkolon As Integer, Optional distinctmi As Boolean = True)
Dim a As Range
Dim dict As New Scripting.Dictionary 'Reference olarak eklenmiş olmalı, ekli değilse Late binding olarak yaratılabilir
 
If alan.Columns(1).Rows.Count = 1048576 Then
    Set alan = Range(alan(1, 1), alan(1, 1).End(xlDown))
End If
 
    For Each a In alan.Resize(, 1)
        If Not dict.Exists(a.Value) Then
            dict.Add a.Value, a.Offset(0, kaçıncıkolon - 1).Value
        Else
            geçici = dict(a.Value)
            dict.Remove (a.Value)
            x = a.Offset(0, kaçıncıkolon - 1).Value
            If distinctmi = True Then
                dict.Add a.Value, geçici & IIf(InStr(1, geçici, x, vbTextCompare) > 0, "", ";" & x)
            Else
                dict.Add a.Value, geçici & ";" & a.Offset(0, kaçıncıkolon - 1).Value
            End If
        End If
    Next a
 
çok_eşleşmeli_vlookup = dict(aranan.Value)
End Function	
	</pre>
	<h3>OFFSET</h3>
	<p>Bir referansa(başlangıç noktasına) göre X satır altında/üstünde ve Y sütun 
	sağında/solunda hücreye başvurmak istediğimizde OFFSET fonksiyonunu 
	kullanırız. Pozitif rakamlar sağı ve aşağıyı gösterirken, negatif rakamlar 
	solu ve yukarıyı gösterir.</p>
	<p>Yükseklik ve genişlik olmak üzere iki de opsiyonel parametresi 
	vardır. Bunlar belirtilmezse, referansın yükseklik ve genişliği baz 
	alınır. Özetle syntax şöyle: <span class="keywordler">OFFSET (Referans, 
	satır, sütun, [yükseklik], [genişlik])</span></p>
	<pre class="formul">=OFFSET(A2;2;1) //B4
=OFFSET(C2;2;-1) //B4
=OFFSET(C2;-2;1) //D1</pre>
	<h4>Kayarak ilerleyen alanlardaki formüller</h4>
	<p><img src="/images/kayanlookup1.jpg"></p>
	<p>Yukardaki tabloda D ve E kolonları arasında Nisan eklense bile formülün kendisi de içeriği de aynı kalır;E2 
	için konuşacak olursak formül D2-C2'dir, yeni kolon eklenince de aynı kalır.</p>
	<p><img src="/images/kayanlookup2.jpg"></p>
	<p>Araya Nisan eklenince, istediğimiz sonucu elde etmek için formülü değiştirmemiz 
	yani E2-D2 
	yapmamız gerekir. Amma ve lakin, her ay yeni kolon eklendikçe de bununla uğraşamayız. 
	Şimdi aynı şeyi OFFSET ile yapalım ve "2 kolon soldaki rakamdan 1 kolon soldaki 
	rakamı çıkar" diyelim.</p>
	<p><img src="/images/kayanlookup3.jpg"></p>
	<p>Araya kolon girince formülün sonucu 
	farklılaşır. Bundan sonra E kolonuna Nisan rakamlarını koymak yeterli 
	olacaktır.</p>
	<p><img src="/images/kayanlookup4.jpg"></p>
	<p>Böyle, sürekli değişen satır/sütun işlemlerinde OFFSET kullanmak 
	idealdir. Son 3 ayın ortalaması, son 5 ayın maksimumu v.s gibi.</p>
	<pre class="formul">=AVERAGE(OFFSET(OFFSET(H2;0;-1);0;0;1;-3))</pre>
	<p><img src="/images/kayanlookup5.jpg"></p>
	<p>Formül açıklaması şöyle:Önce OFFSET(H2;0;-1) formülü ile bulunduğumuz 
	hücrenin bir soluna gidiyoruz. Bu sefer bunu referans verip -3 genişlik 
	diyoruz, bu şu demek oluyor. Bu hücre dahil, sola doğru toplam 3 hücre 
	genişliğinde bir hücre grubu döndür, sonra da bunların ortalamasını al.</p>
	<p>Komşu hücreye giderek referans bulmak kolay, komşu olmayan hücrede nasıl 
	yapacağımıza da bakalım. Hem bu sefer de sütunda bu işlem nasıl yapılır onu 
	görelim.	</p>
	<p><img src="/images/kayanlookup6.jpg"></p>
	<pre class="formul">=AVERAGE(OFFSET(C1;COUNT(C:C);0;-3))</pre>
	<p>Buradaki açıklama da şöyle:C1 hücresini referans al, C kolonundaki 
	numerik içerikli hücre sayısı kadar yani 6 satır aşağı in, yani C7 
	hücresine, ve sonra da C7 dahil olmak üzere önceki yukarı doğru 3 hücre seç, 
	yani C5,C6,C7, bunların da ortalamasını al.</p>
	<h4>OFFSET'e referans olarak kolon da verilebilir</h4>
	<p>Mesela belli bir kolonun hep bir sağındaki kolondaki rakamları toplamak 
	istersek aşağıdaki formülü yazarız.</p>
	<pre class="formul">=SUM(OFFSET(E:E;0;1))</pre>
	<h4 id="dnr">Dinamik Named Range	</h4>
	<p>Normalde data kaynağımızın dinamik olmasını istediğimizde
	<a href="HomeMenusu_Tablolar.aspx">Table</a> kullanmayı tercih ederiz. 
	Böylece datamıza yeni alan eklendiğinde bu data kümemize erişen başka 
	formülleri güncellemek zorunda kalmayız. Ancak bazı durumlarda (bir nedenden 
	ötürü) tablomuzu Table yapamayız. <strong>Txt</strong> dosyalarının otomatik refresh olduğu 
	durumlar gibi. Böyle durumlarda <strong>Dinamik Named Range </strong>kullanmamız gerekir. 
	Böylece data kümemiz her yeni data geldikçe otomatik genişler, data silindikçe de otomatik daralır.</p>
	<p>Bunun için <strong>Named Rangelerden, OFFSET  ve COUNTA </strong>formüllerinden faydalacağız. Genel formülümüz şöyledir</p>
	<pre class="formul">
=OFFSET(başlangıç_hücresi,0,0,COUNTA(tüm_kolon),COUNTA(teksatırlık_range))</pre>
<p>Formülün mantığı şu. Bşalnagıç hücresinden itibaren satır ve üstun olarka ilerleme, yerinde kal, ama yüksekliğin 
tüm kolonda seçtiğin dolu hücre sayısı kadar olsun, genişliğin de belirttiğin kolon sayısı kadar olsun. Bu son parametreye 
duruma göre 1 de girilebilir.</p>
	<p>Aşağıdaki örnek üzerinden bakacak olursak, listeye sürekli siciller 
	ekleniyor ve biz A kolonundaki tüm sicillerin olduğu alana (şu an için A2:A9) 
	"Siciler" ismini vermek istiyoruz. Name ekleme işlemini yaparken şu formülü 
	yazacağız.</p>
<pre class="formul">
=OFFSET('dinamik named range'!$A$2;0;0;COUNTA('dinamik named range'!$A:$A);1)</pre>

<img src="/images/exceldinamikrange.jpg">

<p>Name içine yazılmış hali de aşağıdadır.</p>
	
	<img src="/images/exceldinamikrange1.jpg">

<p>Yeni bir kayıt eklendiğinde 
	F1 hücresinin hemen 9 olduğunu görüyoruz, zira Siciler Name'i de otomatik 
	genişlemiş ve A2:A10 olmuştur.</p>

<img src="/images/exceldinamikrange3.jpg">

<p>Diğer Named Rangelerde 
	olduğunun aksine Dinamik Range'ler NameBox içinde görünmezer. Aşağıdaki ilk 
	resimden görüleceği üzere önceki örneklerde tanımladığımız diğer 3 Name var, 
	ama Siciler yok. Fakat Formulas menüsünden Name'e girip baktığımızda 
	görünmektedir ve onun formülüne tıkladığımızda 2.resimde olduğu gibi ilgili 
	alanın çevrelendiğini görebiliriz.</p>
	
	<img src="/images/exceldinamikrange2.jpg">
	<p>Bu alandaki uniqe sicillleri bir <a href="DataMenusu_VeriDogrulama.aspx">data validation</a> listesi içinde de kullanmak isteyebilirsiniz. 
	Bunun için Çeşitli Örnekler
	kısmındaki 3.örneğe bakabilirsiniz.</p>
	</div>

<!--********************************************************************************************************************************************--><h2 class="baslik">
	XLOOKUP(Office 365)</h2>
	<div class="konu">
		<p>Sadece <strong>Office 365</strong> kullanıcılarında aktif olan bu fonksiyon, yukarıdaki birçok işi bünyesinde barındırıyor. Bakalım bu fonksiyonla neler yapabiliyoruz:</p>
        <ul>
            <li>Hem sola giden vlookup yapabiliyoruz, ki bu hem UDF hem de INDEX-MATCH alternatifidir</li>
            <li>Joker karekterleri destekler(match type)</li>
            <li>Aramaya baştan veya sondan başlayabilir(search mode). </li>
            <li>Bu fonksiyon aynı zamanda bir <a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx#dinamikdizi">dinamik dizi</a> fonksiyonu olduğu için birden fazla sonuç da döndürebilir ve bunları sağa doğru &quot;döker&quot;(dök[ül]me kavramının ne olduğunu linkten okuyabilirsiniz) </li>
            <li>İçiçe iki Xlookup kullanarak kolaylıkla kesişim buldurabilirsiniz(Aşağıda bahsettiğim kesişim yöntemlerine alternatif)</li>
            <li>Çoklu kriter verebilirsiniz</li>
        </ul>
        <p>
            Ben detay örnekler yapana kadar <a href="https://exceljet.net/excel-functions/excel-xlookup-function">bu sayfadan</a> çeşitli örneklere bakabilirsiniz.</p>
	</div>


<!--********************************************************************************************************************************************--><h2 class="baslik">
	Satır ve Sütun Fonksiyonları</h2>
<div class="konu">
<p><span class="keywordler">ADDRESS(satır,sütun,[tip],[stil],[sayfa])</span>:Verilen 
satır ve sütun için A1 veya $A$1 gibi hücre adresi döndürür. Seçimli parametre 
olan "tip" için 1,2,3,4 değerleri girilebilir. Default değer 1'dir, yani mutlak adres($A$1) 
döndürür. 4 ise göreceli demektir. 2 ve 3 ise yarı mutlak yarı göreceli. Stil 
parametresini çok kullanmayacağız. Sayfa olarak da istersek bu adresi başka bir 
sayfanın adresi olarak döndürmemizi sağlayan sayfa ismi girilebilir.</p>
	<p>Bu fonksiyonu tek başına kullanmak pek birşey ifade etmez. Bunu diğer 
	fonksiyonlarla kullandığımızda ne işe yarayacağını daha iyi anlayacaksınız. 
	Bu kullanım şekillerini aşağıdaki çeşitli örnekler bölümünde görebilirsiniz.</p>
	<p><span class="keywordler">ROW/COLUMN([referans])</span>:Belirtilen 
	referansın satır/sütun numarasını döndürür. Hücre belirtilmezse formülün 
	girildiği hücrenin kendisi baz alınır. Referans olarak bir hücre grubu 
	girilirse ilk hücre baz alınır.</p>
	<pre class="formul">=ROW(C2) //2
=COLUMN(D4:F6) //4
=ROW() //8.satırda diyelim, 8</pre>
	<p>Bunlar da tek başına kullanılmak yerine bir fonksiyon kombinasyonu olarak 
	kullanılır. Aşağıda örnekleri göreceğiz.</p>
	<p><span class="keywordler">ROWS/COLUMNS([referans])</span>:Belirtilen 
	referansta kaç adet satır/sütun olduğunu döndürür.</p>
	<pre class="formul">=ROWS(C2) //1
=COLUMNS(D4:F6) //3</pre>
	<p>Bunlar da tek başına kullanılmak yerine bir fonksiyon kombinasyonu olarak 
	kullanılır. Aşağıda örnekleri göreceğiz.</p>
	<p>Yine de ROWS'un güzel bir kullanımını burada örneklendirmek istiyorum. 
	Aşağıdaki tabloda bir filtreleme işlemi yaptım ve toplam kaç kayıttan ne 
	kadarının gösterildiğini en tepeye yazdırım. (Normalde bir alana filtre 
	uyguladığınızda durum çubuğunun sol köşesinde ne kadar kaydın gösterildiği 
	yazar, ancak bu alandan çıkıp tekrar geri geldiğinizde göstermez)</p>
	<p>Formüllerimiz şöyle:</p>
	<pre class="formul">=ROWS(Table1) //tabloda kaç kayıt var
=SUBTOTAL(3;Table1[Hacim]) //B1'deki formül.Filtrelenmiş alanda kaç satırın gösterildiğini verir</pre>
<p>C1'de ise bu ikisini birleştiriyorum.(Aslında tabiki sadece C1'i göstermek 
lazım ama ben size parça parça göstermek istediğim için ayrı ayrı yazdım)</p>
	<p><img src="/images/fonksiyonsrowsfilter.jpg"></p>

</div>


<!--********************************************************************************************************************************************--><h2 class="baslik">
	Diğer Fonksiyonlar</h2>
<div class="konu">
<p><span class="keywordler">CHOOSE(indeks,değer1,değer2...değer254)</span>:1-254 
arasında verilen indeks numaralarına denk gelen değerleri döndürür. Özellikle Ay 
no veya gün noya göre ay adı/gün adı yazdırmada çok faydalıdır. Böylece içiçe if 
yapmaktan veya bir Vlookup bölgesi oluşturmaktan kurtulmuş olursunuz. Tabiki 
liste çok uzunsa işlemi Vlookup ile yapmak çok daha makul olacaktır.</p>
	<pre class="formul">=CHOOSE(A2;"Ocak";"Şubat";"Mart")</pre>
	<p>Bir başka örnek de belirli kişilere rasgele bölge kodu/doğum yılı v.s 
	atamak olabilir.</p>
	<pre class="formul">=CHOOSE(RANDBETWEEN(1;3);1979;1981;1992)</pre>
	<p><img src="/images/fonksiyonchoose.jpg"></p>
	<p>Choose'a parametre olarak belirli hücre grupları da verebilir ve bu dönen 
	hücre alanlarını başka fonksiyonlarla birlikte kullanabiliriz. Mesela 
	aşağıdaki örnekte, yapılan seçime göre SUM edilen hücre alanı değişmektedir.</p>
	<pre class="formul">=SUM(CHOOSE(A2; D2:D10; E2:E10;F2:F10)) </pre>
	<p><img src="/images/fonksiyonchoose2.jpg"></p>
	<p><strong>NOT</strong>: Yukardaki Combobox kullanımının detaylarını
	<a href="DeveloperMenusu_Kontroller.aspx">buradan</a> öğrenebilrsiniz.</p>
	<p>Choose fonksiyonunun dinamik grafik yapımında nasıl kullanıldığını görmek 
	için
	<a href="http://chandoo.org/wp/2013/04/23/interactive-chart-in-excel-tutorial/">
	bu sayfaya</a> göz atmak isteyebilirsiniz.</p>
	<p><span class="keywordler">INDIRECT</span>:Metin formundaki bir referansı 
	gerçek bir referansa dönüştümeye yarar. Aşağıdaki örnekte A4 hücresine 
	INDIRECT(A3) formülünü girdim. A3'te A1 metni var, bunu bir referans olarak 
	algılayıp, A1 hücresindeki değeri getirdi.</p>
	<p><img src="/images/excelindirect1.jpg"></p>
	<p>Bu fonksiyon bize <a href="DataMenusu_VeriDogrulama.aspx">Data Validation</a>(Veri 
	Doğrulama)'ın da desteğiyle birbirine bağımlı Comboboxlar yapmamızı sağlar. 
	Bunu bir örnekle açklayalım.</p>
	<p>Diyelim ki Data Validation ile kullancıya bir hücrede il isimlerini 
	seçtiriyoruz. Kullanıcı İstanbul seçince, alttaki comboboxta yine Data 
	Validaiton ile İstanbul ilçeleri gelsin istiyoruz. Bunu şöyle yaparız. </p>
	<ul>
		<li>İl ve ilçeleri bi yere aşağıdaki gibi altalta yazarız.</li>
		<li>İlçe isimlerini seçip bunları aşağıdaki mavi dairedeki gibi Named 
		Range yaparız.</li>
	</ul>
	<p><img src="/images/excelindirect2.jpg"></p>
	<ul>
		<li>Sonra Data Validation ile B1 hücresinin kaynağını aşağıdaki gibi 
		belirleriz</li>
	</ul>
	<p><img src="/images/excelindirect3.jpg"></p>
	<ul>
		<li>Son olarak da B2 hücresinin kaynağını aşağıdaki gibi belirleriz.</li>
	</ul>
	<p><img src="/images/excelindirect4.jpg"></p>
	<p>B1'de İstanbulu seçince B2'deki formül İstanbul Name'indeki değerleri 
	içine yükler.</p>
	<p><img src="/images/excelindirect5.jpg"></p>
	
</div>	

	<h2 id="Kesisim" class="baslik">Kesişim Bulma</h2>
	<div class="konu">
	<p>Excelde iki boyutlu bir veri kümesinde belirli satır ve sütunların 
	kesiştiği&nbsp; hücreyi bulmanın birkaç yolu vardır. Hepsine bakacağız. 
	Siz, o 
	anki ihtiyacınıza hangisi uyuyorsa onu kullanabilirsiniz. </p>
	<p>Şimdi öncelikle, aşağıdaki matrise bakalım. Burada C Grubunun 
	2.seviyesine denk gelen rakamı yani 11000i bulmak istiyoruz diyelim.</p>
	<p><img src="/images/vlookupkesisim.jpg"></p>
	<h4>1.Yöntem:Offset ve Match kombinasyonu</h4>
	<pre class="formul">=OFFSET(A3;MATCH(C9;$A$4:$A$7;0);MATCH(C10;$B$3:$D$3;0))</pre>
	<p>Yaptığımız işlemin açıklaması şöyle:Offset ile A3 hücresini referans 
	alıyoruz, ikinci parametre olarak Match ile 2'nin kaçıncı satırda olduğunu 
	bulup onu veriyoruz, üçüncü parametre olarak da yine Match ile C'nin kaçıncı 
	sütünda olduğunu öğrenip onu veriyoruz. Sonuç olarak A2'nin 2 satır aşağısı 
	ve 3 sütun sağ tarafına bak diyoruz, yani D5 hücresine.</p>
	<h4>2.Yöntem:Vlookup ve Match kombinasyonu</h4>
	<pre class="formul">=VLOOKUP(C9;$A$3:$D$7;MATCH(C10;$B$3:$D$3;0)+1;0)</pre>
	<p>Burda yaptığmız işlem ise Vlookup'a aranan değer olarak 2'yi vermek, 
	aranan alan olarak A-D kolonlarını vermek. Kaçıncı kolona bakması 
	gerektiğini ise 
	C10'daki değeri B3:D3 arasında nerde buluyorsa onun 1 fazlası olacak şekilde 
	buluyor.</p>
	<h4>3.Yöntem:Index-Match kombinasyonu</h4>
	<pre class="formul">=INDEX($B$4:$D$7;MATCH(C9;$A$4:$A$7;0);MATCH(C10;$B$3:$D$3;0))</pre>
	<p>Bu yöntem, kesişim deyince akla gelen ilk yöntemdir aslında. Yazımı 
	öncekilere göre biraz daha uzundur ama esas varoluş sebebi bakımından 
	Index-Match tam bir kesişim bulma yöntemidir. Bunun kullanım şeklini zaten 
	yukarıda gördüğümüz için ayrıca açıklamaya gerek görmüyorum.</p>
	</div>
	
	<h2 class="baslik">Çeşitli Örnekler</h2>
	<div class="konu">
		<h4 class="baslik">Bir alandaki kolon sayısına göre ortalama almak</h4>
			<div>
			<p>Ne zaman bir alandaki kolon sayısına ihtiyaç duyarız? Mesela dinamik bir alanınız var ve 
			burdaki rakamları toplayıp kolon sayısına böldürdüğünüz bir formül var. Bu, aslında ortalama aldırmak oluyor 
			ancak klasik AVERAGE formülünü kullanmak istemiyoruz, çünkü arada bazı boş hücreler var, biz boş hücreleri
			 0 olarak ele almak istiyoruz, halbuki AVERAGE formüülü boş hücreleri dikkate almaz. Ama biz 0 olmasını 
			 istiyoruz çünkü o ay ilgili kişin 0 satış yaptığını biliyoruz ve ortalaması düşsün istiyoruz, daha doğrusu
			 haksız yere ortalması yüksek çıksın istemiyoruz. Şimdi formülümüzü görelim.</p>
			 
			 <pre class="formul">
		=SUM(B2:G2)/COLUMNS(B2:G2)	 </pre>
			 	<p><img src="/images/excelcolumnsortalama.jpg"></p>
			</div>
		
		<h4 class="baslik">Bir alandaki son hücrenin adresi</h4>
			<div>
			<p>Bazen bir hücre grubundaki son hücrenin adresini elde etmek isteriz, ve sonra 
			bunu da başka bir hücredeki formülün argümanı olarak kullanırız. Dinamik bir 
			yapı olsun istediğimiz için&nbsp; de dinamik bir formülle tespit etmemiz 
			lazım.</p>
			<pre class="formul">=ADDRESS(ROW(Table1)+ROWS(Table1)-1;COLUMN(Table1)+COLUMNS(Table1)-1)</pre>
				<p>Formül şöyle işliyor. Row(Table1) ile alanın nerden başladığını 
				buluyoruz, hatırlayacak olursanız ROW fonksiyonu bir alanla 
				kullanıldığında ilk hücresini baz alıyordu. ROWS ile toplam kaç satır 
				olduğunu buluyoruz ve 1 çıkararak son hücrenin satır numarasını 
				buluyoruz. Örneğin Table1 A8:A500 arasını kapsıyorsa, 8+483-1=500.satır 
				olduğunu buluruz. Aynı mantıkla da sütun numarası elde edilir. Sonra 
				bunlar ADDRESS ile birleştirilir</p>
			</div>
		
		<h4 class="baslik">Bir alandaki uniqe değerleri listelemek</h4>
			<div>
			<p></p>
		
			 
			</div>
</div>


</asp:Content>
