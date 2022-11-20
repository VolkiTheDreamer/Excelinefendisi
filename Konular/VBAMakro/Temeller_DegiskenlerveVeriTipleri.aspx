<%@ Page Title='Temeller DegiskenlerveVeriTipleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Temeller'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Değişkenler ve Veri Tipleri</h1>

<h2 class="baslik">Değişkenler</h2>
<div class="konu">
<p>Değişkenler, adı üstünde, tüm kod çalışırken zamanla içeriği değişebilecek alanlardır.(Tabiki tüm kod boyunca sabit de kalabilirler, adları değişken diye illa değişmek zorunda değillerdir.)</p>

	<h3>Neden ve Nasıl?</h3>
	<p>Değişken tanımlamamızın iki temel nedeni var:</p>
	<ol>
		<li>Bilgisayarın belleğini daha verimli kullanmak için. Her değişken için uygun veri tipini 
		tanımlarsanız bellekte o kadar az yer tutulmuş olur. Aksi halde 
		değişkenler Variant tipte olurlar ve en yüksek hacim olan 16 Byte'lık 
		bir hafıza işgal ederler.</li>
		<li>Hatalı işlem yapılmasını engellemiş olursunuz.</li>
	</ol>
	<p>Birçoğunuzun bu aşamada ileri seviye programlamalar yapmayacağını varsayarak burada daha fazla bu konunun detaylarına şimdilik girmeyeceğim. 
	<span class="keywordler">Public/Private</span> ayrımına da bu sayfada girmeyeceğim ancak isterseniz konularda biraz ilerleyince&nbsp; <a href="Temeller_Birazdahaterminoloji.aspx">buradan</a> detaylı bilgiye ulaşabilrsiniz.</p>

	<p>Değişkenleri tanımlamak için nerdeyse her zaman <span class="keywordler">Dim</span> 
		ifadesini ve değişkenden hemen sonra <span class="keywordler">As</span> ifadesini 
		ve hemen arkasından veri tipini yazarız.</p>
	<p>Ör:</p>
<pre class="brush:vb">
Dim İsim As String
Dim Yas As Integer
Dim Agirlik As Double
Dim Gerçekleştimi As Boolean	</pre>
	<p>Bunların hiçbirini tanımlamasanız da çok büyük ihtimalle programınız 
	çalışacaktır. Bunları tanımlayarak hem daha derli toplu hem de daha hızlı 
	çalışan bir makroya sahip olmuş olursunuz.</p>

	<p>Değişken kullanmak efektiftir de, mesela i=Range("A1").Value demezseniz ve 
	Range("A1")'deki değeri birkaç yerde kullanıyorsanız ve sonradan Range("A1") yerine 
	Range("A2")'ye referans vermek isterseniz kod içinde Range("A1") geçen her 
	yerde değişiklik yapmak zorunda kalırsınız, ama i=Range("A1").Value 
	demişseniz sadece bunu değiştirmeniz ve diğer yerleri "i" olarak bırakmanız 
	yeterlidir.</p>

	<p>Değişikenin tanımlı olması ve olmaması durumunda bellek kullanımı nasıl oluyor bir ona bakalım. Bunun için basit bir örnek üzerinden gidelim.</p>
	
	<pre class="brush:vb">
Sub NoVariable()
	Range("A1").Value = Range("B2").Value
	Range("A2").Value = Range("B2").Value * 2
	Range("A3").Value = Range("B2").Value * 4
	Range("B2").Value = Range("B2").Value * 5
End Sub	
</pre>

	<p>Yukardaki kodda, VBA B2 hücresine 5 kez başvurmaktadır. Yani işlemciyi 5 kez yormuş oluyoruz. Halbuki en baştan B2'nin değerini bir değişkene atamış olsak, işlemcimiz sadece bir kez yorulmuş olacaktır. Ayrıca Eğer B2 hücresindeki değeri silip B4'e taşımamız gerekseydi, kodumuzun 5 yerinde değişiklik yapmamız gerekecekti. Bir diğer detay da bu kodda kullanılan karakter sayısı fazla ve uzun gözüküyor.</p>
	<p>Şimdi bi de değişken kullanarak yapalım:</p>

	<pre class="brush:vb">
Sub WithVariable()
  Dim i as Integer
  i = Range("B2").Value
  Range("A1").Value = i
  Range("A2").Value = i * 2
  Range("A3").Value = i * 4
  Range("B2").Value = i * 5
End Sub	
</pre>
	<p>Şimdi ne olur? İşlemcimiz, B2 hücresine 5 kez değil 1 kez başvuracak, bu da 
	daha az bellek kullanımı demektir. Ayrıca B2 yerine B4 kullanmak istersek 
	sadece i değişkenini tanımladığımız satırda yani bir kez değiştirmemiz 
	yeterli ve okunuşu daha kolay bir kod.</p>

	<p>Bunun gibi 5-6 satırlık bir kodda çok büyük bir fark göremezsiniz, ama 
	yüzlerce satırdan oluşan bir kodunuz olursa farkı o zaman hissedebilirsiniz.</p>
	<p>Peki değişkenleri tanımlamadığımızda ne oluyor, program nasıl çalışıyor? VBA, bu tür değişkenleri 
	<span class="keywordler">Variant</span> tipinde depolar ve her 
	defasında bu değişkenin ne tür bir değişken olması 
	gerektiğine karar vermeye çalışır, bu da zaman kaybıdır, yani ağır çalışan 
	bir kod demektir. Bu nedenle sitenin anasayfasında dediğim gibi işimizi doğru yapmakla 
	kalmayalım, zerafet içinde doğru yapalım, hızlı yapalım, o yüzden değişkenlerimizi mutlaka 
	tanımlayalım.</p>

	<p>Aşağıda tipi belirtilerek değişken tanımladığımızda ve tanımlamadığımızda 
	neler olduğunun süreyle ölçülmüş halini gösteren bir örneğimiz daha var. Fark oldukça açık ve net.</p>
	<p>Bu kod, k değişkeni tanımlanmadan çalıştırıldığında 5,88 saniye 
	sürmekteyken k'nın önündeki ' işareti kaldırılıp tanımlama yapıldığında ise 
	2,13 saniye sürmekte. İşte bellek yönetimi budur!</p>

<pre class="brush:vb">
Sub timerkontrol()
Dim başlangıç As Single
Dim bitiş As Single
Dim i As Long 
'Dim k As Long

başlangıç = Timer 'bu fonksiyon, kodunuzun ne kadar sürede çalıştığını tespit etmek için kullanılır

For i = 1 To 100000000 'Bu yapı For-Next döngüsüdür. Şimdilik bu döngünün nasıl kullanıldığını bilmiyor olabilirsiniz, buna takılmayın. Sonraki bölümlerde detaylıca incelenecek.
  k = k + 1
Next i

bitiş = Timer

MsgBox ("İşlem süresi:" &amp; vbNewLine &amp; Round(bitiş - başlangıç, 2) &amp; " saniyedir.")
End Sub</pre>

	<p>Bu arada bazı zamanlar olacaktır ki değişkenin içeriği sürekli 
	değişebiliyordur, bazen sayısal bazen karekter bir içeriğe sahip 
	olabiliyordur, hatta sadece sayısaldır ama bazen integer, bazen long integer 
	veya bazen double olabiliyordur veya bir sebepten dolayı tipi bilinmiyordur, 
	bu durumda değişkeni <strong>Variant</strong> olarak tanımlamaktan başka bir 
	çare yoktur, sadece tipini Variant olarak belirtiyoruz, tip belirtmezsek de 
	Excel onu Variant olarak algılar. Yani şu ikisi de aynı şekilde algılanır</p>
		<pre class="brush:vb">
Dim deger As Variant
Dim deger</pre>


	<h3>Değişken tanımlama kuralları</h3>
	<p>Değişkenleri tanımlamanın bi faydası da şudur: Değişkenlerinizde mutlaka 
	bir tane Büyük harf bulunursa ve bundan sonra değişkenleriniz hep küçük 
	harfle yazıp space'e veya Enter'a bastığınızda, Excel ilgili harfi 
	otomatikman büyük harfe çevirir, eğer büyük harf olmuyorsa anlarsınız ki 
	değişkeninizi yanlış yazmışsınız. Bu nedenle, burda bu uyarıyı yapmakta da 
	fayda görüyorum.</p>

	<p><span><strong>UYARI!</strong></span>:Değişken isimlerinizde iki kelime 
	veya daha çok varsa mutlaka bir harf, mümkünse 
	ortadan bir harf, büyük olsun. Kod içinde diğer heryerde küçük harfle 
	yazıp, VBA'in büyütmesini bekleyin. Buna programcılık dilinde camelCase notasyonu denmektedir. Genel olarak tüm programlama dillerinde en çok önerilen değişken 
	tanımlama geleneği camelCase olarak bilinen yöntemdir. Ör:</p>
<pre class="brush:vb">
Dim okulNo
Dim ayAdı	
enAltSatırNo 
</pre>
<p>Böylece siz programınızın başka bir yerinde <strong>okulno</strong> yazdığınızda otomatikman 
	<strong>okulNo</strong> olacaktır, aynı şekilde <strong>ayadı</strong> da 
	otomatikman <strong>ayAdı</strong> olacaktır.</p>

	<p>Değişken tanımlamada dikkat edilecek diğer hususlar şöyledir:</p>
	<ul>
		<li>İlk karakter bir harf olmalı</li>
		<li>boşluk, nokta(.), ünlem(!), ve şu karakterler kullanılmaz (@, &amp;, $, 
		# )</li>
		<li>Çok geçeğini sanmam ama karakter uzunluğu 255i geçmemeli</li>
		<li>VB rezerv kelimeleri kullanılmamalı (for, next, String vs.)</li>
		<li>Türkçe karakter serbest ancak başka birçok dilde geçersiz olduğu için kullanmamaya çalışın</li>
	</ul>
	
	<p>Program boyunca sabit kalacak bir değişkeniniz varsa bunu <span class="keywordler">Const</span> ifadesi 
	ile tanımlayabilirsiniz. Örneğin, mağaza sayısı 15 olan bir firma için bunu 
	sabit olarak tanımlayabilir ve döngülerde bu sabiti kullanabilirsiniz.</p>
		<pre class="brush:vb">
Const magaza As Integer = 15 'aynı satırda tanımlanmak zorunda
For i = 1 to magaza
'Kodlar
Next i
</pre>
	
	<p><strong class="dikkat">Dikkat:</strong>Sık yapılan hatalardan biri de şudur. Şimdi iki String değişken 
	tanımlamak istediğimizi düşünelim. Eğer kelimeden tasarruf edeyim deyip şu şekilde tanımlarsak hata yaparız:</p>
		<pre class="brush:vb">
Dim metin1, metin2 as String</pre>
	
	<p>Çünkü bu şekilde aslında ilk değişkenin tipi belirtilmemiş oldu ve bu 
	yüzden Variant oldu, String değil. Bu nedenle şu şekilde tanımlama yapmalıyız.</p>
		<pre class="brush:vb">
Dim metin1 as String, metin2 as String
'veya daha güvenli olsun isterseniz
Dim metin1 as String
Dim metin2 as String
</pre>
	
	<h3>Option Explicit</h3>
	<p>Eğer buraya kadar okuduklarınızdan değişkenleri tanımlamanın gerçekten iyi bir fikir olduğunu 
	düşünüyorsanız, bütün modüllerinizin başında <span class="keywordler">Option Explicit</span> 
	ifadesi bulunsun, bu sizi değişkenleri tanımlamaya zorlayacaktır.</p>
	<pre class="brush:vb">
Option Explicit
Sub zorunlu()
  mesaj="Merhaba"
  MsgBox mesaj
End sub	
</pre>
	<p>Yukardaki kodu çalıştırdığınızda hata verecektir, çünkü "mesaj" değişkeni tanımlanmamıştır.</p>
	<p>Her modülün başına tek tek bu ifadeyi yazmak istemiyorsanız, şu ayarlamayı yapın. VBE içinde <strong>Tools>Options</strong> düğmesine basın ve aşağıdaki seçeneği işaretleyin.</p>
	<p><img src="/images/Vbaoptionexplistmenu.jpg"></p>
	<h3>Değer atama ve Default(Varsayılan) değerler</h3>
	<p>Değişkenleri tanımladıktan sonra onlara bir de değer atamak gerekir.<span>Değişkenler, 
	değer atanana kadar varsayılan değerlere sahip olurlar</span>. Buna göre;</p>
	<ul>
		<li>Sayısal tipte bir değişken için varsayılan değer <strong>0'dır.</strong></li>
		<li>Karekter/String tipinde bir değişken için varsayılan değer <strong>""</strong> yani sıfır uzunluklu stringtir.</li>
		<li>Nesne tipinde bir değişken için varsayılan değer ise <strong>Nothing'</strong>dir.</li>
	</ul>
	<p>Bu konuda daha detaylı ve karşılaştırmalı bilgiye <a href="../Fasulye/NeNeredeNasil_NullNothingEmptyveIlkdegeratama.aspx">buradan ulaşabilirsiniz</a>.</p>
	<p>Şimdiye kadarki örneklerde gördüğünüz üzere sayısal veya karakter tipte bir değişkene değer ataması için = işaretini kullanırız. Ör:</p>
	<pre class="brush:vb">
i=0
ay="Ocak"
</pre>
	<p>Bir nesneye(Range, sheet, collection v.s) değer atamak içinse <span class="keywordler">Set</span> 
	ifadesini kullanırız. 
	Ör:</p>
<pre class="brush:vb">
Dim hucre As Range
Dim ws As Worksheet

Set hucre=Range("A1")
Set ws=Activesheet
</pre>
	<p>Nesne tanımlamalarında, kod bitiminde bu nesneleri tekrar 
	<span class="keywordler">Nothing</span>&nbsp;olarak 
	atamak bellek yönetimi açısından faydalıdır, böylece Excel bu nesneler için 
	bellekte gereksiz yer ayırmayacaktır.</p>
<pre class="brush:vb">
Sub Ornek()
	Dim hucre as Range
	Set hucre = Range("A1")
	'Diğer Kodlar
	Set hucre = Nothing
End Sub</pre>

	<h3 id="static">Static deyimi ile değişken tanımı</h3>
	<p>Static kavramı biraz daha ileri seviye konularındandır, ancak tek 
	başına ileri seviye konularının arasında sırıtacağı ve anlam bütünlüğü 
	açısından da buraya daha uyduğu için buraya almak durumunda kaldım. Ayrıca 
	ileri seviye terminolojik konuların ele alındığı
	<a href="Temeller_Birazdahaterminoloji.aspx#scopelifetime">şu sayfada</a> 
	Global değişkenlerle kıyaslaması da bulunamaktadır.</p>
	<p>Bu deyimle tanımlanmış değişkenlere ben zombi değişken diyorum, zira 
	tanımlandıkları prosedür çalışmayı tamamlasa bile yaşamaya devam ederler, ta 
		ki içinde bulundukları workbook kapanana kadar. Bu yüzden bunlara 
	hafızalı değişkenler de denmektedir.</p>
	<h4>Neden tanımlanır?</h4>
	<p>Dim ifadesi ile tanımladığımız tüm değişkenler ilgili prosedür 
	çalıştırıldıktan sonra bellekten silinir. Ancak bazı durumlarda, 
	tanımladığımız değişkenin prosedür çalıştıktan sonra bile bir önceki değerini tutmasını bekleriz.</p>
	<p>Aşağıdaki örnek kodu 3 kez çalıştırdığımızda sırayla şunu görürüz: i:1, 
	j:1, sonra i:1, j:2 ve en son i:1,j:3.</p>
	<pre class="brush:vb">
Sub statictest()
    Dim i As Integer
    Static j As Integer
    
    i = i + 1
    Debug.Print "i:" &amp; i
    
    j = j + 1
    Debug.Print "j:" &amp; j
End Sub	</pre>
	<p>
	Aşağıda ise günlük hayat içinden bir örnek var.</p>
	<p>
	<span class="dikkat">UYARI</span>:Bundan sonrasına devam etmeden önce
	<strong>Application.OnTime</strong> metodunun öğrenilmesinde veya genel bir fikir edinilmesinde fayda var.</p>
	<p>
	Mesela her 5 dakikada çalışacak şekilde bir ayarlanmış bir prosedür düşünün. Sabah işe 
	geldiğinizde 9:00 gibi çalışmaya başlatıyorsunuz, akşam 18:00'e kadar da 5 dk aralıklarla kendisi Refresh 
	olup çalışıyor. Saat 17:00 olduğunda belli alıcılara bir mail atsın istiyoruz 
	diyelim. Sabah çalıştırma saatimiz her zaman tam net 9:00 olmayacağı için "saat=17:00 
	mı" diye kontrol edemeyiz. Bunun yerine saat 16'dan büyük mü diye kontrol 
	etmeliyiz. Diyelim ki bir önceki çalıştırma işlemi 16:58:23'te oldu, bir 
	sonraki 17:03:23te olacak, tam bu anda yakalarz, ama bi sonraki de 
	17:08:23te olacak, yine kontrole takılır ve bu şekilde 12 kez mail gitmiş 
	olur, ki böyle birşey istemeyiz. </p>
	<p>
	İşte böyle bir durumda statik bir değişken tanımlayabiliriz. Aşağıdaki 
	örneğe bakalım. Saat 16'dan büyükse i'yi her defasında 1 artırıyoruz ancak 
	sadece i=1 ise yani ilk kez gidecekse mail gönderimi yapıyoruz.</p>
	<pre class="brush:vb">
Sub statikornek()
Static i As Integer

If Hour(Now) &gt; 16 Then 'Bu yapı koşullu karşılaştırma yapmamazı sağlayan IF yapısıdır. Şimdilik bu yapının nasıl kullanıldığını bilmiyor olabilirsiniz, buna takılmayın. Sonraki bölümlerde detaylıca incelenecek.
    i = i + 1
    If i = 1 Then Call mailproseduru
End If

Application.OnTime Now + TimeValue("00:05:00"), "statikornek" ' 5 dk sonra kendisini tekrar çalıştırıyor

End Sub

Sub mailproseduru()
  MsgBox "mail gönderimi"
  'diğer kodlar
End Sub</pre>

	<p>Bu işlemi pek tabiki saatin 17:05:00'ten küçük olup olmadığına bakarak da 
	yapabilirdik ama Static konusunu anlamak adına uygun bir örnek olacağını 
	düşündüm.</p>
	<p>
	Bu arada bir diğer alternatif de i'nin değerini sayfada boş ve görünmeyen 
	bir hücreye yazıp bunun değerini kontrol etmek olabilirdi ama static değişken kullanmak daha şık bir yöntemdir.<br>
</p>
	</div>
	
<h2 class="baslik">Veri Tipleri ve Boyutlar</h2>
<div class="konu">
<p>Exceldeki temel veri tiplerini iki gruba ayırabiliriz, sayısal ve sayısal olmayan. 
Toplamda 14 çeşit temel veri tipi bulunur. Neden bu kadar veri grubu var dersek, birinci sebep kullandıkları hafıza 
miktarı, ikincisi ise belli işlemlerin sadece belli veri tipleriyle yapılmasını 
sağlayarak hatalı işlemlerin olmasını engellemek diyebiliriz. Şimdi bunlara yakından bakalım.</p>

<h3>Sayısal veri tipleri</h3>
<table class="alterantelitable">
<tr><th>Tip</th><th>Alabileceği değerler</th><th style="text-align: center">Hafıza Kullanımı (Byte)</th></tr>
<tr><td>Byte </td><td>0 ile 255  </td><td style="text-align: center">1</td></tr>
<tr><td>Integer </td><td>-32.768 ile 32.767   </td>
	<td style="text-align: center">2</td></tr>
<tr><td>Long  </td><td>-2.147.483.648 ile 2.147.483.648   </td>
	<td style="text-align: center">4</td></tr>
<tr><td>Single </td><td>-3.4*10^38 ile 3.4*10^38 (Küsuratlar ihmal edilmiştir)</td>
	<td style="text-align: center">
	4</td></tr>
<tr><td>Double </td><td>-1.7*10^308 ile 1.7*10^308 (Küsuratlar ihmal edilmiştir)</td>
	<td style="text-align: center">
	8</td></tr>
</table>
<p><strong>Currency</strong> diye bir tip daha var ama <span>kariyer hayatım 
boyunca</span> hiç kullanmadım, o yüzden buraya koymuyorum.</p>
	<p>Bunlardan en çok küsurata sahip olanı Double'dır. İhtiyacınıza göre 
	birini kullanırsınız. Aşağıda küsuratları gösteren bir örnek var.</p>
	<pre class="brush:vb">Sub tipler()

Dim d As Double, s As Single, c As Currency, l As Long, i As Integer, t As Byte

a = 9
b = 7
d = b / a
s = b / a
c = b / a
l = b / a
i = b / a
t = b / a

Debug.Print d, s, c, l, i, t 'sırasıyla 0,777777777777778 0,7777778 0,7778 1 1 1

End Sub</pre>

<p>Bir de <strong>Decimal</strong> diye bir tip var, bu en yüksek küsurat 
duyarlığına sahip veri tipidir. Ancak bunun kullanımı biraz daha alengirli, o 
yüzden burada detaya girmiycem, zaten bunu da Currency tipi gibi <span>kariyer 
hayatım boyunca </span>kullanma ihtiyacım hiç olmadı.</p>
	<h3>Sayısal olmayan tipler</h3>
<table class="alterantelitable">
<tr><th>Tip</th><th>Alabileceği değerler</th><th>Hafıza Kullanımı (Byte)</th></tr>
<tr><td>String(sabit boyutlu) </td><td>Max 65400 karakter</td><td>sabit boyut</td></tr>
<tr><td>String(değişken boyutlu) </td><td>Max 2 milyar karakter</td><td>10+karakter sayısı</td></tr>
<tr><td>Date  </td><td>Makul bir tarih girin yeter</td><td>8</td></tr>
<tr><td>Boolean </td><td>True ve False  </td><td>2</td></tr>
<tr><td>Object </td><td>Çok çeşitli(Ör:Range, Workbook, Collection)</td><td>4</td></tr>
<tr><td>Variant(Varsayılan veri tipi)</td><td>Herşey olabilir</td><td>16</td></tr>
</table>

<p>Burada String biraz enteresan görünüyor. Bununla ilgili bir örnek yapalım:</p>

<pre class="brush:vb">
Dim isim As String 'Boyut belirtmediğimiz için değişken boyutludur
isim="Volkan" 'şuan 6 karakter içerir+10=16 Byte
isim="Mustafa Kemal" 'şuan 13 karakter içerir+10=23 Byte</pre>
	<p>
	Şimdi diyebilirsiniz ki, Variant tipi 16 Byte tutuyor, değişken boyutlu 
	String'i neden kullanalım ki, yukardaki örnekte 23 Byte oldu. Variant'ın 16 
	Byte tutma olayı, ilk tanımlama anındadır. Sonrasında, içersindeki 
	değişkenin tipine göre boyutu artabilir.</p>
	<p>Devam edelim, şimdi de sabit boyutlu String nasıl tanımlanır ona bakalım.</p>
	<pre class="brush:vb">Dim isim as String*15
isim="Volkan" 'Hafızada şöyle tutulur "Volkan         " '15i tamamlayacak kadar boşluk eklenir</pre>
<h3>Temel olmayan veri tipleri</h3>
<p>Bu yukarıdaki temel veri tipleri dışında Range, Collection, Worksheet gibi çeşitli class(sınıf)lardan üretilen nesneler var, o yüzden bunlar da veri tipi olarak düşünülebilir. <strong>Dim hucre As Range</strong> ifadesinde olduğu gibi. Bunun dışında ileri VBA sayfalarında göreceğimiz gibi kendi tanımladığımız sınıfları da veri tipi olarak düşünebiliriz. Mesela Student diye bir class tanımladıysanız <strong>Dim st As Student şeklinde</strong> bir değişken tanımlayabilirsiniz.</p>
	<h3>Enumeration/Constant</h3>
	<p>Bazı sabitler vardır ki bunlar Excel açılır açılmaz yüklenirler ve 
	kullanıma hazırdırlar. Bunların kimisinin kullanımı faydalı iken kimisi 
	zorunludur. Mesela <strong>vbNullString </strong>boş string amaçlı olarak "" 
	yerine kullanılabilir. Böylece boşuna bellekte yer ayrılmamış olur, zira bu 
	vbNullString zaten o anda bellektedir, yani ilave bellek tüketmez. Ancak 
	mesaj kutularına verilen cevabın evet mi olduğunu anlamak için <strong>vbYes</strong> 
	sabitini kullanmak zorunludur.</p>
	<p>Biz bu sabitlerden 3 çeşitini kullanıyor olacağız.</p>
	<ul>
		<li>vb ile başlayanlar: VBA <span>(kütüphanesinin)</span>library'sinin sabitleri</li>
		<li>xl ile başlyanlar: Excel library'sinin saibtleri</li>
		<li>mso ile başlayanlar: Office library'sinin sabitleri</li>
	</ul>
	<p>Bunların bir de numerik karşılıkları vardır, bunlara Enumeration 
	değerleri denir. Örneğin vbYes'in değeri 6'dır. bunların ikisi de 
	kullanılabilir.(NOT:İlerde kodlarınızı VSTO ortamına geçirmek istediğinzde 
	numerik karşılıklarını kullanmanız gerekecektir)</p>
	<p>Mesela aşağıdaki kodu yazıp çalışıtığrıdığınızda size bu sabitin sayısal 
	karşılığı olan 65280'i verir.</p>
	<p><img src="/images/vbaenum1.jpg"></p>
	<p>Ayrıca kendi Enumeration tiplerinizi de oluşturabilirsiniz.</p>
	<pre class="brush:vb">Private Enum Bolgeler
  Akdeniz '0
  Karadeniz '1
  Marmara '2
  Ege '3
  Anadolu '4
End Enum</pre>
	<p><img src="/images/vbaenum2.jpg"></p>

</div>
</asp:Content>
