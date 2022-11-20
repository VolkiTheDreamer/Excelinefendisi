<%@ Page Title='DizilerveDizimsiYapilar DizilerArray' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Diziler ve Dizimsi Yapılar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Diziler(Arrayler)</h1>
	<h2 class="baslik">Giriş</h2>
	<div class="konu">
	<p>Kodumuzda aynı türden bir veya iki farklı eleman kullanacaksak standart 
	değişkenler yeterli olacaktır. Ancak bu sayı artarsa değişken kullanımı pek pratik 
	olmamaya başlar. İşte bu nokta, dizilerin devreye girdiği yerdir.</p>
	<p>Mesela 20 bölgesi olan bir bankada çalışıyorsanız her bölge için bir 
	değişken tanımlamak yerine <strong>bölge</strong> adında bir dizi 
	tanımlanıp, her bölge bu diziye elaman olarak atanabilir.&nbsp; 
	Karşılaştırmaya bakalım:</p>
	<pre class="brush:vb">'pratik olmayan yöntem
Dim bölge1 As String
Dim bölge2 As String
.....
Dim bölge20 As String</pre>
	<p>Bir de bunun şube versiyonu var, ki şube sayısı 1000 civarındaysa 
	değişken tanımlarken ömrünüzden 1 yıl gider herhalde.</p>
	<p>Peki şimdi nedir bu diziler, nasıl tanımlanır, başka neler yapılır, 
	bunlara bakalım.</p>
	<p>Genel bir bilgi edindikten sonra
	<a href="../Fasulye/NeNeredeNasil_Diziler.aspx">buraya</a> da bakmanızı 
	tavsiye ederim. Özellikle başka programlama dillerine aşinaysanız aradaki farkları ve benzerlikleri görmek için faydalı olacağını düşünüyorum.</p>
</div>
<h2 class='baslik'>Temel bilgiler</h2>
<div class="konu">
	<h3>Tanımlama(Declaration)</h3>
	<p>Diziler, normal değişkenler gibi <strong>Dim</strong> ifadesi ile 
	tanımlanırlar, ancak değişken adının yanında fazladan yuvarlak parantezlere 
	sahiptirler. Parantezler; dizi boyutu baştan belliyse boyut numarasını 
	içerirler, belli değilse boş bırakılır ve sonradan tanımlanır. Dizi boyutu 
	baştan belirlenen dizilere <strong>Statik</strong> dizi, boyutu baştan 
	belirtilmeyen dizilere <strong>Dinamik</strong> dizi denir.</p>
	<pre class="brush:vb">Dim diziadı(boyut) As Tip </pre>
	<p>Tip belirtilmezse tıpkı normal değişkenlerde olduğu gibi dizimiz <strong>
	Variant </strong>tipli bir dizi olur ve her tür değişkeni karışık olarak 
	depolayabilir.</p>
	<pre class="brush:vb">Dim bölgeler(19) As String 'boyut belirtildi
Dim şubekod() As String 'boyut belirtilmedi</pre>
	<p>Bu örnekte bölge sayısı sabit olduğu için boyut baştan belirtildi, yani 
	Statik tanımlandı ancak 
	bir bankada şube sayısı çok sık değişebilir; yeni şubeler açılır, mevcut 
	şubeler kapanır, şubeler birleşir;o yüzden ona baştan bir boyut belirtmesek 
	de olur, yani Dinamik tanımlandı. Bu dinamik dizilere aşağıda ayrıca bakıyor olacağız.</p>
	<h3>Boyut(elaman sayısı), index, alt/üst limitler</h3>
	<p>Dizi için tanımlanan eleman sayısına boyut denir. Boyut belirtmenin de iki yolu vardır.</p>
	
<ul>
<li><strong>Dim dizi(x to y) As Tip</strong></li>
<li><strong>Dim dizi(y) As Tip</strong></li>
</ul>

	<p>İlk yöntemde dizinin kaçıncı indexten başlayıp kaçta biteceği 
	belirtilirken ikinci yöntemde ise doğrudan kaçta biteceği belirtilir, kaçta 
	başlayacağı ise <span class="keywordler">Option Base</span> ifadesinin 
	kullanılıp kullanılmadığına göre değişir. İndeksin biteceği son yere üst 
	sınır denir ve bu sınır <span class="keywordler">Ubound </span>fonksiyonu 
	ile elde edilir.</p>
<pre class="brush:vb">Dim segmentler(4) As String
segmentler(0)="Bireysel"
segmentler(1)="Birebir"
segmentler(2)="Kobi"
segmentler(3)="Ticari"
'segmentler(4)="Özel"

For i=0 to UBound(segmentler)
    MsgBox segmentler(i)
Next i</pre>
	<p>Yukardaki örnekte gördüğünüz üzere son indexli elamana değer atanmadı, bu 
	yüzden de boş olarak göründü. Anlayacağınız üzere, dizideki tüm elemanlara 
	değer atanmak zorunda değildir. Yani, bir dizide eleman sayısı ile dolu(değer 
	atanmış) eleman sayısı aynı olmayabilir.</p>
	<p>Varsayılan olarak dizilerin ilk eleman indeksi 0'dır. Ancak bu index
	<span class="keywordler">Option Base 1</span> ifadesi ile 1 yapılabilir, ki 
	ben bunu zorunda olmadığınız sürece çok kullanmanızı tavsiye etmiyorum. Veya 
	yukarda 1. yöntemdeki gibi dizi boyutu 
	baştan <strong>1 to y</strong> şeklinde belirtilerek de başlangıç indexi 1 
	yapılabilir. Yani şu iki ifade tamamen özdeştir.</p>
	<pre class="brush:vb">Option Base 1
Dim bolgeler(20) As String</pre>
	<p>ve</p>
	<pre class="brush:vb">Dim bolgeler(1 to 20) As String</pre>
	<p>İndex numarası açısından olmasa da eleman sayısı açısından şu dizi de 
	yukardakilerle aynı kapasitededir, yani hepsi de 20 eleman içerir.</p>
	<pre class="brush:vb">Dim bolgeler(19) As String</pre>
	<p>Dizi tanımının yönteminden ve <span class="keywordler">Option Base</span><strong>
	</strong>kullanımından bağımsız olarak alt indexin ne olduğu ise
	<span class="keywordler">LBound </span>fonksiyonu ile elde edilir.</p>
	<p><strong>Option Base 1</strong> kullanımını tavsiye etmiyoruz dedik, zira 
	ilgili modüldeki tüm prosedürler için indexi 1den başlatır. Bununla beraber 
	bazı dizileri bilinçli olarak 1 nolu indeksten başlatmak gerekebilir. Mesela 
	ayno isimli bir dizimiz olduğunu düşünün, ayno(1) diyince Ocak ayını ele 
	almak, anlaşılırlık açısından daha makbuldür, pek tabiki ayno(1) içinde 
	Şubat ayı da depolanabilir ancak bu yol, konuşma diline biraz aykırılık teşkil 
	edeceği için bunu ayno(1 to 12) şeklinde tanımlamayı tercih etmek daha 
	akıllıca olacaktır.</p>
	<p>Bu arada <strong>x to y</strong> yönteminde x olarak 1'den büyük değerler 
	de belirtilebilir ama bunun pratikte çok kullanıldığı görülmez.</p>
	<p>Çok boyutlu dizilerde eleman sayısı, boyutlardaki elemanların çarpımına 
	eşittir.(Bunlar aşağıda ayrıca detaylı incelenecek)</p>
	<p><strong>Özetleyecek olursak;</strong></p>
	<p>Dim b(10) As String şeklinde tanımlanan 1 boyutlu bir dizide;</p>
	<table class="alterantelitable">
			<th>Aranan</th>
			<th>Yöntem</th>
			<th>Sonuç</th>
		<tr>
			<td>Alt limit</td>
			<td>LBound(b)</td>
			<td>0</td>
		</tr>
		<tr>
			<td>Üst limit</td>
			<td>UBound(b)</td>
			<td>10</td>
		</tr>
		<tr>
			<td>Eleman sayısı</td>
			<td>Ubound(b)-LBound(b)+1</td>
			<td>11</td>
		</tr>
	</table>
	
		<p>Dim b(1 to 10) As String şeklinde tanımlanan 1 boyutlu bir başka dizide;</p>
	<table class="alterantelitable">
			<th>Aranan</th>
			<th>Yöntem</th>
			<th>Sonuç</th>
		<tr>
			<td>Alt limit</td>
			<td>LBound(b)</td>
			<td>1</td>
		</tr>
		<tr>
			<td>Üst limit</td>
			<td>UBound(b)</td>
			<td>10</td>
		</tr>
		<tr>
			<td>Eleman sayısı</td>
			<td>Ubound(b)-LBound(b)+1</td>
			<td>10</td>
		</tr>
	</table>


	<p>Dim b(10,5) As Integer şeklinde tanımlanan 2 boyutlu bir dizide;</p>
	<table class="alterantelitable">
			<th>Aranan</th>
			<th>Yöntem</th>
			<th>Sonuç</th>
		<tr>
			<td>1.boyutun Alt limiti</td>
			<td>LBound(b,1)</td>
			<td>0</td>
		</tr>
		<tr>
			<td>1.boyutun Üst limiti</td>
			<td>UBound(b,1)</td>
			<td>10</td>
		</tr>
		<tr>
			<td>2.boyutun Alt limiti</td>
			<td>LBound(b,2)</td>
			<td>0</td>
		</tr>
		
		<tr>
			<td>2.boyutun Üst limiti</td>
			<td>UBound(b,2)</td>
			<td>5</td>
		</tr>
		<tr>
			<td>1.boyutun Eleman sayısı</td>
			<td>Ubound(b,1)-LBound(b,1)+1</td>
			<td>11</td>
		</tr>
		<tr>
			<td>2.boyutun Eleman sayısı</td>
			<td>Ubound(b,2)-LBound(b,2)+1</td>
			<td>6</td>
		</tr>
		<tr>
			<td>Dizideki Eleman sayısı</td>
			<td>1. boyut elaman sayısı*2.boyut eleman sayısı</td>
			<td>66</td>
		</tr>

	</table>


	<p><span class="dikkat">Dikkat:</span>Bir dizinin boyutu baştan birkez 
	belirtilirse bir daha asla değişmez. Boyutu değişen dizilere ise aşağıda 
	değineceğiz.</p>
	<h3>Elemanlara değer atama</h3>
	<h4>İlk değer atama</h4>
	<p>Dizi elamanlarına index numaralarıyla ulaşırız. Normalde bunlara tek tek 
	değer atamak genelde pratikte karşılaşılan bir durum değildir, belki küçük 
	boyutlu dizilerde olabilir ancak genelde bir döngüsel yapı ile bir hücre 
	grubundan değer okuyup onları atamak şeklinde olmaktadır. Aksi halde kodumuz oldukça 
	uzayacaktır. </p>
	<pre class="brush:vb">Dim Segment(3) As String
Segment(0) = "Bireysel"
Segment(1) = "Birebir"
Segment(2) = "Kobi"
Segment(3) = "Ticari"</pre>
	<p>Döngüyle atamaya örnek olarak da şunu verebiliriz;A1:A20 arasındaki bölge 
	kodlarını bölge dizisine atıyoruz.</p>
	<pre class="brush:vb">For i=1 to 20
   bölge(i)=Cells(i,1).Value2
Next i</pre>
	<h4>Değer atanmamış elemanlar ve Erase fonksiyonu</h4>
	<p>Henüz değer atanmamış dizi elemanları, dizinin tipine göre default 
	değerlerini alırlar. Bunlar;</p>
	<ul>
		<li>String diziler için sıfır uzunluklu metin, yani ""</li>
		<li>Nümerik diziler için 0</li>
		<li>Variant diziler için Empty</li>
		<li>Object diziler için Nothing</li>
	</ul>
	<p>Statik dizilerde elemanların hepsini tek seferde varsayılan değerlerine 
	döndürmek için <span class="keywordler">Erase</span> fonksiyonu kullanılır. 
	Yani dizi elemanlarının taşıdığı değerler boşaltılır. Ancak bunlar hala 
	hafızada yer kaplamaya devam ederler.</p>
	<pre class="brush:vb">Sub erasestatik()
  Dim segment(2) As String
  segment(0) = "bireysel"
  segment(1) = "birebirt"
  segment(2) = "kobi"

  Erase segment()
  
  Debug.Print segment(1) '"" döndürür
End Sub</pre>
	<h4>Array fonksiyonu</h4>
	<p>Elle tek tek tanımlama yapılması gereken durumlarda, bir diğer eleman 
	tanımlama yöntemi de <span class="keywordler">Array</span> fonksiyonunu 
	kullanmaktır, ki bu sadece elaman değeri atama değil aynı zamanda 
	diziyi tanımlama yöntemidir de. Zira bu şekilde tanımlanan diziler <strong>Variant</strong> 
	tipte olurlar ve bildiğiniz gibi <strong>Dim </strong>ile tanımlanmayan 
	tüm değişkenler Variant kabul edilirler. Bu sayede elemanlar tanımlanırken dizi de 
	yaratılmış olur. Ama biz yine de iyi bir programcı olup dizimizi Dim ile tanımlayalım. <strong>Bu yöntemle tanımlanan dizilerde başlangıç 
	indeksi her zaman 0 olur. Yine, bu yöntemle yaratılan diziler 1 boyutlu olur</strong></p>
	<pre class="brush:vb">Dim Segment As Variant
Segment = Array("Bireysel","Birebir","Kobi","Ticari") 'tek satırda tanımlama imkanı

Debug.Print Segment(2) 'Kobi yazar</pre>
	<p>Bu yöntemin pratik bir kullanımı "ayisim" gibi bir fonksiyon şeklinde 
	kendini gösterebilir. Fonksiyonlarla kafanızı karıştırmamak için şimdilik Sub prosedür şeklinde örneğe görelim.</p>
	<pre class="brush:vb">Sub ayisimornek()
  Dim ayisim As Variant
  ayisim = Array("", "Ocak", "Şubat") 'ilk ayı boş geçtim, index 0 olduğu için
  Debug.Print ayisim(2)
End Sub</pre>
	<h3>Dizi içinde dolaşma ve elemanlara erişim</h3>
	<p>Dizilere eleman atama işini döngülerle yaptığımız gibi, dizi elamanlarını 
	okumayı da yine genelde döngülerle yaparız. Ender olarak kod içinde bir yerde bir 
	index numarası temin edip onu doğrudan da kullandığımız da olur.</p>
	<pre class="brush:vb">For i=1 to 20
   Cells(i,1).Value2=bölge(i)
Next i</pre>
	<p>Daha şık şekli aşağıdaki gibi olabilir, hatta sadece şık değil aynı 
	zamanda güvenlidir de. </p>
	<pre class="brush:vb">For i=LBound(bölge) to UBound(bölge)
   Cells(i,1).Value2=bölge(i)
Next i</pre>
	<p>Evet, biraz daha fazla kod yazmış&nbsp; olduk ama 
	güvenlik ön planda olacaksa hard-code(rakamı doğrudan belirterek) alt-üst 
	limit vermek yerine bu fonksiyonlarla vermek daha verimlidir. Hard-code 
	yazıldığında, bölge sayısı 1 arttığında 20 
	yerine 21 yapmayı unutursanız kodunuz yine hata almadan çalışır ancak eksik çalışmış olur, ve 
	belki akabinde çalışan kodlarla bölgelere otomatik mail gidiyorsa, son 
	bölgeye hiç mail gitmemiş olur. Bir de içiçe bir dizi varsa ve her seviye 
	için üst limit değişiyorsa Ubound'ı kullanmak zaten zaruri olacaktır.</p>

<p> Üstelik bu yazım şekli alt limitin 0 mı 1 mi olduğunu da pek önemsemez . 
Böylece <strong>Option Base</strong> var mıydı yok muydu, diziyi (1 to 20) şeklinde mi yoksa (20) 
şeklinde mi tanımladığınızı hatırlamak zorunda kalmazsınız.</p>
	<p> <span class="dikkat">dikkat</span>:Dizi içinde For döngüsü ile 
	dolaşırken diziye eleman atama işlemi sadece For i=1 to 10 şeklindeki basit 
	For döngüsü ile yapılabilirken, eleman değerini okuma işlemi ise hem basit 
	For ile hem de For Each ile yapılabilir.</p>

<p>Mesela aşağıdaki gibi bir kullanım sorunsuz çalışırken,</p>

<pre class="brush:vb">
Dim Bölge(1 to 20) As String
For i=1 to 20
   bölge(i)=Cells(i,1).Value2
Next i
</pre>

<p>bu kod da hata vermeden çalışır ancak eleman değerlerinin değişmediğini görebilirsiniz. (Eğer daha önce bir değer atanmadıysa içleri boş görünecektir.)</p>

<pre class="brush:vb">
Dim Bölge(1 To 20) As String

For Each b In Bölge
   b = ActiveCell.Value2
   ActiveCell.Offset(1, 0).Select
Next b
</pre>

</div>


<h2 class='baslik'>İleri seviye işlemler</h2>
<div class="konu">
	<h3>Dinamik diziler</h3>
	<p>Yukarıdaki örnekler hep statik dizi örnekleriydi(Array fonksiyonu ile tanımlananlar 
	hariç)</p>
	<p>Dizimizin boyutunu baştan bilmiyorsak ve kodun gidişatına göre değişken bir 
	şekilde karşımıza çıkma durumu varsa, diziyi boyutsuz yani dinamik tanımlarız 
	ve zamanı geldiğinde boyutunu belirtiriz. Bunu da <span class="keywordler">ReDim</span><strong>
	</strong>ifadesi ile yaparız. 
	Aslında yaptığımız şey, yeni bir statik dizi yaratmaktır. Zira 
	<strong>ReDim</strong> kullanıp da yeni boyutunu verdiğimiz diziyi yine yeterli bulmazsak 
	tekrar boyut değiştirebiliriz.</p>
	<pre class="brush:vb">Dim şubeler() As String
....
....
ReDim şubeler(800)</pre>
	<p>Statik dizilerde, <span style="text-decoration: underline">genelde</span> kaynak 
	israfı sözkonusudur, zira baştan olası en 
	yüksek değere göre boyut belirlenir ve bu boyutların hepsi çoğu zaman 
	kullanılmaz. Tabiki bölge sayısı gibi kesin olarak bilinen ve hepsi 
	kullanılan durumlar istisnadır. Bu nedenle genel olarak Statik değil Dinamik 
	dizi kullanılması önerilir.</p>
	<h4>Erase fonksiyonu</h4>
	<p><span>Dinamik dizilerde <strong>Erase </strong>fonksiyonu statik 
	dizilerden farklı çalışır. Statik dizilerde Erase, elemanları default 
	değerlerine atarken dinamik dizide diziyi boyutsuzlaştırır, yani bir nevi 
	siler. Diziyi tekrar kullanmak isterseniz ReDim ile yeniden boyutlandırmanız 
	gerekir. </span></p>
	<h4>İlave boyut artışı</h4>
	<p>Diyelim ki ReDim ile 10 elemanlık bir boyutlandırma yaptınız ancak öyle 
	bir nokta geldi ki, 10 eleman yetmiyor, yani boyut artırmanız lazım, böyle 
	bir durumda <span class="keywordler">ReDim Preserve</span> ifadesini 
	kullanırız. Preserve demezsek dizi yine yeni boyuta göre boyutlanır ancak 
	önceki elemanlar silinmiş olur. Preserve ile ilk 10 elemanın değerini de 
	korumuş oluruz.</p>
	<p>Boyut artırma durumların ReDim'in bir veya iki kez kullanılması önerilir. Birkaç kez boyut artırma ihtiyacı oluyorsa belki dizi değil de Collection kullanmak faydalı olabilir.&nbsp;</p>
	<p><strong>NOT</strong>:Variant olarak tanımlanan <strong>değişkenler</strong> doğası gereği dinamiktirler.(Variant 
	tanımlanan <strong>diziler</strong> ise statiktir)</p>
	<pre class="brush:vb">Dim statikVar As Variant 'standart Variant değişken
Dim dinamikVar(5) As Variant 'Variant tipli dizi</pre>
	<h4>Variant değişken vs Variant dizi</h4>
	<p>Variantlarla ilgili detaylı bir örnek aşağıda bulunmaktadır. Bunun<span> 
	oldukça aydınlatıcı olduğunu düşünüyorum.</span></p>
	<pre class="brush:vb">Sub variantlı()
Dim aylar1 As Variant 'içine istediğiniz tipte değer atayabileceğiniz bir DEĞİŞKENDİR, buna dizi de dahildir,
'ama dizi atayana kadar dizi değildir ve IsArray testi false döner. Ubound da kullanamazsınız
Dim aylar2() As Variant 'Parantez olduğu için bu kesinlikle DİZİDİR, IsArray true döner

'***öncelikle parantezsiz olan yani değişken olan Aylara bakalım*****
aylar1 = 1 'number atandı
Debug.Print aylar1
'Debug.Print aylar1(1) 'hata verir, çünkü içeriği henüz bir dizi değil.
Debug.Print IsArray(aylar1) 'false
'Debug.Print UBound(aylar1) 'çalışmaz çünkü şuan için dizi değil
aylar1 = "volki" 'string
Debug.Print IsArray(aylar1) 'yine false
Debug.Print aylar1
aylar1 = Now 'tarih
Debug.Print IsArray(aylar1) 'hala false
Debug.Print aylar1
aylar1 = Array("hi", "world", "naber") 'dizi
Debug.Print IsArray(aylar1) 'artık true, çünkü Array function ile array yaptık
Debug.Print UBound(aylar1)
Debug.Print aylar1(2) 'artık hata vermez, çünkü içeriği bir dizi
aylar1 = Array(Array("hi", "world", "naber"), Array(1, "ss", 3)) 'dizi dizisi
Debug.Print UBound(aylar1)
Debug.Print aylar1(1)(1)
aylar1 = Range("a1:a5").Value ' range atadık, 2 boyutlu dizi
Debug.Print UBound(aylar1) '5 döner
Debug.Print aylar1(2, 1) '2.satırdaki değer döner

'***şmdi de Dizi olan Aylara bakalım
Debug.Print IsArray(aylar2)
'Debug.Print UBound(aylar2) hata verir, zira henüz boyut belli değil, ya ReDim yapılmalı ya da Array ile değer atanmalı
ReDim aylar2(3)
Debug.Print UBound(aylar2)
aylar2 = Array("OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK") '
Debug.Print UBound(aylar2)
Debug.Print aylar2(2)
Debug.Print UBound(aylar2)
'şimdi de farklı data tiplerinde atama yapıyoruz, Variant olduğu için sorun olmuyor
aylar2 = Array("OCAK", 2, 3, "30.04.2015", Now + 60, "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK") ' variant olduğu için içine karışık tipli verilen atayabilirm
Debug.Print aylar2(4)

End Sub</pre>
	<h3>Çok boyutlu diziler</h3>
	<p>Şimdiye kadar gördüğümüz dizilerin çoğu tek boyutlu idi, ama ihtiyacımıza 
	göre birden fazla 
	boyutlu diziler de oluşturabiliyoruz. Pratikte üçten fazla boyutun 
	kullanıldığını ben açıkçası ne duydum ne de gördüm. Şahsen kendi kodlarımda 
	kullandığım dizilerin büyük kısmı tek boyutludur, birkaç tane 2 boyutlu bir 
	tane de 3 boyutlu dizim bulunmakta. 3 boyutlu dizi örneğinde aynı zamanda 
	Dictionary de kullandığım için bu örneği o bölümde ele alacağız.</p>
	<p>Şimdi 2 boyutlu bir dizi kullanım örneğine bakalım. Excel sayfalarının 
	kendileri zaten satır ve sütunlardan oluşmakta olup 2 boyutludurlar ve bu 
	durumu 2 boyutlu dizilerle birlikte çok sık kullanıyor olacağız.</p>
	<p>Mesela 10 bölgesi olan bir bankada her bölgenin bir Bireysel bir de 
	Ticari müdürü olduğunu düşünelim. Bunları bir diziye atama işlemi nasıl 
	oluyor ona bakalım.</p>
	
	<p><img alt="" src="/images/vbadizi1.jpg"></p>
	<pre class="brush:vb">
Sub ikiboyutludizi()
Dim mudur(1 To 10, 1 To 2) As Long 'ilk boyut bölge için, ikinci boyut segment tipi için
Dim i As Integer, j As Integer

For i = LBound(mudur, 1) To UBound(mudur, 1)
    For j = LBound(mudur, 2) To UBound(mudur, 2)
        mudur(i, j) = Cells(i + 1, j + 1).Value
    Next j
Next i

Debug.Print mudur(3, 2)
End Sub</pre>
	<p>
	Tabi bunun için <a href="#rangevariantdizi">aşağıda</a> başka bir yöntem var 
	ki, bundan çok daha basittir.</p>
	<h3>İçiçe diziler(Dizi dizileri)</h3>
	<p>İngilizcede Array of Array ve Jagged Array olarak kullanılan bu dizi 
	türünü Variant tipte dizi tanımlayarak elde ederiz. Bilindiği gibi Variant 
	veri tipi içinde herşeyi tutabilir, buna diziler de dahildir. Yanlız burdaki 
	ayrıma dikkat etmek lazım, klasik Variant bir değişken tanımlamıyoruz, 
	Variant tipli bir dizi tanımlıyoruz. Aradaki ayrımı görmek için şu iki 
	satıra bakalım:</p>
	<pre class="brush:vb">Dim d As Variant 'bu Variant tipli klasik bir değişkendir
Dim d() As Variant 'bu ise Variant tipli bir dizidir</pre>
	<p>İçiçe dizilerin tanımlaması iki ayrı dizi şeklinde olur, ama kullanımı <strong>dizi(x)(y)</strong> 
	şeklindedir. Yani aslında bu dizi türü tek boyutludur ama içerdiği her 
	elaman da bir başka dizidir.</p>
	<pre class="brush:vb">Dim dışdizi() As Variant 'ana dizi Variant olmalıdır
Dim içdizi() As String ' Alt dizinin Variant olması gerekmez

'iç diziyi dış diziye eleman olarak atama
dışdizi(n)=içdizi
'nihai kullanım şekli
Debug.Print dışdizi(x)(y)</pre>
	<p>Bu dizilerin kullanımı çok boyutlu dizilere benzer ancak, kullanım 
	ihtiyacı ve şekli küçük farklar göstermektedir. Şöyle ki, çok boyutlu 
	dizilerde tümboyutlar için hafızada yer ayrılır, belki o boyutlardan bazısı 
	hiç kullanılmayacak bile. Diyelim 20 bölgeli bir banka var, bazı bölgelerde 
	40 şube varken bazısında 15 şube bulunabilir. Şimdi <strong>Şubeler(1 to 20, 1 
	to 40)</strong> şeklinde 2 boyutlu bir dizi tanımlarsak bazı bölgelerde bazı boyutlar israf olacak. İşte içiçe 
	dizilerde bu israf olmuyor. Örneğin 2.bölgenin şube sayısı 18 ise, burdaki 
	18.şubeyi tanımlayacak son elaman bölge(2)(18) oluyor, bölge(2)(19) şeklinde 
	bir tanımlama yapılmıyor.</p>
	<p>Mesela aşağıdaki tabloya göre içiçe dizimizi oluşturalım.</p>
	<p><img src="/images/vbadizi2.jpg"></p>
	<pre class="brush:vb">
Sub içiçedizi()
Dim bolge(1 To 3) As Variant 'dış dizi 
Dim subeler() As String 'iç dizi
Dim i As Integer, s As Integer, k As Integer

For i = 1 To 3
    ReDim subeler(1 To Cells(i + 1, 2)) 'iç dizinin boyutunu ayarlıyoruz, her bölgede bu boyut değişecek
    
    'iç diziyi oluşturalım
    For s = 1 To UBound(subeler)
        subeler(s) = Cells(s + k + 1, 6)        
    Next s

    k = k + s - 1 'satır sayısı resetlenmesin diye geçici bir değişkene atıyorum    
    
    'iç diziyi dış dizinin ilgili elemanına atayalım
    bolge(i) = subeler

Next i

Debug.Print bolge(1)(5) 'Şube5
Debug.Print bolge(2)(2) 'Şube8
Debug.Print bolge(3)(4) 'hata verir, çünkü 3.bölgenin 4.şubesi yok
End Sub</pre>
	<p>
	Dizi dizisiyle yapılabilen işlemler "Collection of Collection" veya 
	"Dictionary of Dictionary" yapılarıyla da yapılabilir. Bunları bir sonraki 
	bölümde görüyor olacağız. Bu arada hepsinin avantaj ve dezavantajını anlatan 
	güzel bir sayfa var, ona
	<a href="http://stackoverflow.com/questions/19633937/can-you-declare-jagged-arrays-in-excel-vba-directly">
	buradan</a> ulaşabilirsiniz.</p>
		<h3>Rangeler ve Diziler</h3>
		<h4 id="rangevariantdizi">
		Bir hücre grubunu diziye atama(Sayfadan okuma)</span></h4>
		<p> Dizileri, bir hücre grubundan hızlı veri okuma ve yazma amacıyla da 
		kullanırız. Bu kullanım şekli, özellikle ilgili veri kümesi üzerinde bir güncelleme(belli 
		bi rakamla çarpmak gibi) yapmak 
		istediğinizde idealdir.</p>
		<p> Tabi bu amaçla kullanmak istediğimizde standart dizi tanımı yerine 
		<strong>Variant</strong> olarak tanımlarız, ve tanımladığımız dizi 
		<strong>2 boyutlu bir dizi</strong> 
		olur. Zira bir sayfaya baktığınızda gördüğünüz şey satır ve sütunlardan 
		oluşan iki boyutlu bir dizidir. Böyle bir dizide de ilk boyut satır 
		ikinci boyut sütun olur. Bu noktada karışıklığa neden olan bir konu vardır 
		ki o da şudur: 
		<span style="text-decoration: underline">Sözkonusu hücre grubu tek kolonluk bir alandan oluşuyorsa bile iki 
		boyutludur</span>. İlk boyut satır sayısı kadar elemanlı, ikinci grup da sütun sayısı olan 
		1 elamanlıdır. <strong>Bu tür dizilerde indeks 0dan değil 1den 
		başlar.(Tek boyutlu variant dizilerde ise 0'dan başlıyordu)</strong></p>
		<p> Mesela aşağaıdaki kod ile A1:A1000 hücrelerindeki değerler 100 ile 
		çarpılır.</p>
		<pre class="brush:vb">Sub rakamguncelle()

Dim rakamlar As Variant 'Dizi tanımı
rakamlar = Range("A1:A10000").Value 'burada hücreden diziye okuma yaptık

For i = LBound(rakamlar) To UBound(rakamlar) 'i'ler satır boyutudur
  rakamlar(i, 1) = rakamlar(i, 1) / 100 '1 ise sütun boyutu
Next i

Range("A1:A10000").Value = rakamlar 'üzerine de yazabilriz başka bir yere de yazabilirdik

End Sub</pre>
	<p>Alanımız tek kolon değil de birden çok kolondan oluşuyorsa sütun için de 
	bir iç döngü daha eklememiz gerekir. Bu sefer hücre grubumuzu dinamik 
	düşünelim, buna göre kodumuzu aşağıdaki gibi olacaktır.</p>
	<pre class="brush:vb">
Dim rakamlar As Variant 'Dizi tanımı
rakamlar = Selection.Value 'burada hücreden diziye okuma yaptık
 
For i = LBound(rakamlar) To UBound(rakamlar) 'i'ler satır boyutudur
    For j = LBound(rakamlar, 2) To UBound(rakamlar, 2) 'j'ler satır boyutudur
        rakamlar(i, j) = rakamlar(i, j) * 1000
    Next j
Next i
 
Selection.Value = rakamlar</pre>
	<h4>Tek boyutlu bir diziyi sayfaya yazdırma</h4>
	<p>Yazdırma işlemi yapmak için <strong>Resize</strong> özelliği ile hücreleri yeniden 
	boyutlandırmayı bilmeniz gerekiyor. Resize hakkında bilgi sahibi değilseniz
	<a href="DortTemelNesne_Range.aspx#resize">buradan</a> bakıp önbilgi 
	edinebilirsiniz.</p>
	<p>Bu işlem için bir <strong>Range</strong> nesnesi yaratırız ve sonra hedef 
	hücreyi belirleriz. Sonra bu hücreyi <strong>Resize</strong> ile dizi boyutu kadar 
	genişletiriz.</p>
	<pre class="brush:vb">' Yatay yazma(Tek satırlı çok kolonlu)
Dim dizi(1 To 3) As Integer
Dim Hedef As Range

dizi(1) = 10
dizi(2) = 20
dizi(3) = 30

Set Hedef = Range("A1").Resize(1, UBound(dizi))
Hedef.Value = dizi
</pre>
	<p>Eğer dizimizi dikey şekilde yani tek kolon çok satırda yazmak isteseydik, 
	Resize argümanlarını ters çevirmemiz ve Transpose metodunu kullanmamız 
	gerekirdi.</p>
	<pre class="brush:vb">
' Dikey yazma(Tek kolonk çok satır)
Dim dizi(1 To 3) As Integer
Dim Hedef As Range

dizi(1) = 10
dizi(2) = 20
dizi(3) = 30

Set Hedef = Range("A1").Resize(UBound(dizi), 1)
Hedef.Value = WorksheetFunction.Transpose(dizi)
</pre>
	<h4>İki boyutlu dizileri sayfaya yazdırma</h4>
	<p>İki boyutlu bir dizimizi sayfaya yazdırmak istiyorsak yine Resize 
	metodunu kullanırız, ilk boyutu satır olarak, ikinci boyutu da kolon olarak 
	kullanabiliriz, veya tam tersini de istersek yerlerini değiştirip Transpoze 
	yaparız.</p>
	<pre class="brush:vb">Dim dizi(1 To 3, 1 To 2) As Integer
Dim alan As Range

dizi(1, 1) = 10
dizi(2, 1) = 20
dizi(3, 1) = 30
dizi(1, 2) = 100
dizi(2, 2) = 200
dizi(3, 2) = 300

Set alan = Range("A1").Resize(UBound(dizi, 1), UBound(dizi, 2))
alan.Value = dizi</pre>
	<p>Transpoze hali de şöyle:</p>
	<pre class="brush:vb">Dim dizi(1 To 3, 1 To 2) As Integer
Dim alan As Range

dizi(1, 1) = 10
dizi(2, 1) = 20
dizi(3, 1) = 30
dizi(1, 2) = 100
dizi(2, 2) = 200
dizi(3, 2) = 300

Set alan = Range("A1").Resize(UBound(dizi, 2), UBound(dizi, 1))
alan.Value = WorksheetFunction.Transpose(dizi)</pre>
	<h3>Dizilerin diğer kullanma şekilleri</h3>
	<h4>Dizileri dizilere atama</h4>
	<p>Gelişmiş programlama dilllerinde olduğunun aksine VBA'de bir diziyi tek 
	seferde başka bir diziye kopyalama/atama yöntemi bulunmamaktadır. Yani şöyle bir kod çalışmaz.</p>
<pre class="brush:vb">
Dim A(1 To 3) As Long
Dim B(1 To 3) As Long

A(1)=100000
A(2)=200000
A(3)=300000
B=A
</pre>

<p>Peki nasıl yapılır? Tabiki döngülerle;</p>
<pre class="brush:vb">
Dim A(1 To 3) As Long
Dim B(1 To 3) As Long

A(1)=100000
A(2)=200000
A(3)=300000

For i=1 to 10
   B(i)=A(i)
Next i

Debug.Print B(2)
</pre>
	<h4>Variant dizileri Variant dizilere atama</h4>
	<p>Dizileri başka dizilere atayamıyoruz ancak içi dizi olan bir Variantı 
	başka bir Varianta atayabiliyoruz. Zira, bu işlem aslında dizi atama değil 
	değişken atamadır.</p>
	<pre class="brush:vb">
Dim A As Variant
Dim B As Variant
A = Array(100000, 200000, 30000)
B=A 'atama başarılıdır
Debug.Print B(2) '30000 yazar</pre>
	<h4>
	Range atanmış Variant dizileri normal dizilere atama</h4>
	<p>
	Bir hücre grubunu doğrudan bir variant diziye atama şeklini yukarda 
	görmüştük, ama bir şekilde bunları tek boyutlu bir diziye atamanız gerekirse 
	bunu yapmanın yolu döngü kullanmaktır. Tabi istenirse arada hiç variant dizi 
	olmadan da doğrudan ilgili alanlar üzerinden de geçilebilir ancak, eğer bir 
	şekilde ilgili alan kod içinde silindiyse, veya büyük bir alan ise peformans 
	sorunu yaşamamak için Variant diziden normal diziye atamak daha verimli olacaktır, 
	zira <strong>diziler üzerindeki işlemler range'lere göre çok daha hızlıdır</strong>.</p>
	<p>
	Şimdi
	aşağıdaki gibi bir hücre grubu olsun, 
	bunu önce Varianta sonra da normal bir diziye atayalım.</p>
	<p>
	<img src="/images/vbadizi4.jpg"></p>
	<pre class="brush:vb">
Sub birkolonlu_variant_to_dizi()
Dim var As Variant
Dim dizi() As String
Dim i As Integer

var = Range("B2:B10").Value2
ReDim dizi(1 To UBound(var, 1))

For i = 1 To UBound(var, 1)
    dizi(i) = var(i, 1)
Next i

MsgBox dizi(3)
End Sub</pre>
	<p>
	Şimdi bunu bir de çok boyutlu dizide yapalım		
	</p>
	<p>
	<img src="/images/vbadizi3.jpg"></p>
	<pre class="brush:vb">
Sub çokkolonlu_variant_to_dizi()
Dim var As Variant
Dim dizi() As String
Dim s As Integer, k As Integer, i As Integer

var = Range("B2:D10").Value2
ReDim dizi(1 To UBound(var, 1) * UBound(var, 2))

s = 0
For k = 1 To UBound(var, 2)
   For i = 1 To UBound(var, 1)
     dizi(i + s) = var(i, k)
   Next i
   s = s + UBound(var, 1) 'her boyut geçişinde tekrar başa dönmesin diye aratoplam alınır
Next k

MsgBox dizi(15)
End Sub	</pre>
	<p>
	Aynı diziyi iki boyutlu bir standart diziye atamak da aşağıdaki gibi 
	olurdur.</p>
	<pre class="brush:vb">
Sub çokkolonlu_variant_to_çokboyutludizi()
Dim var As Variant
Dim dizi() As String
Dim s As Integer, k As Integer, i As Integer

var = Range("B2:D10").Value2
ReDim dizi(1 To UBound(var, 1), 1 To UBound(var, 2))

For k = 1 To UBound(var, 2)
  For i = 1 To UBound(var, 1)
    dizi(i, k) = var(i, k)
  Next i
Next k

MsgBox dizi(2, 3)
End Sub</pre>
	<h4>Dizileri collectiona atama</h4>
	<p>Bazı durumlarda elimizdeki diziyi daha fazla esnekliği olan ve bazı 
	yönlerden dizilere benzeyen collectionlara atamak isteriz. Ancak colletion 
	konusunu henüz işlemediğimiz için(siteyi sırayla okuduğunuz varsayımıyla 
	hareket ederek) bu durumu collection 
	sayfasında ele alacağız.</p>
	<h4>Dizileri bir prosedüre parametre olarak gönderme</h4>
	<p>Dizileri bir Sub veya Function prosedüre parametre olarak 
		gönderebiliriz. Unutulmaması gereken nokta, diziler her zaman <strong>ByRef</strong> 
		anahtar kelimesi ile gönderilir. Aşağıda basit bir örneğimiz bulunuyor.<pre class="brush:vb">
Sub dizigonder()
Dim d(5) As String
'sadece iki elemanı dolduralım
d(0) = "volkan"
d(1) = "meltem"

Call mesajver(d)
End Sub


Sub mesajver(ByRef dz() As String)
    For Each a In dz
        If a <> "" Then i = i + 1
    Next a
eleman = UBound(dz) - LBound(dz) + 1
MsgBox eleman & " adet elemanın " & i & " tanesi doludur"
End Sub	</pre>
		<h4 id="dizisonuc"> Prosedürden sonuç olarak dizi döndürme</h4>
		<p>Fonksiyonlar konusunda göreceğimiz gibi, 
		<span style="text-decoration: underline">genelde</span> bir fonksiyondan 
		sadece tek bir değer döndürülür. Aslında bu birçok yerde doğru diye 
		anlatılan yanlış bir bilgidir. Genelde durum böyledir ancak fonksiyonlar tek bir değer döndürmek 
		zorunda değildir, sadece dönüş değeri tek olmak zorundadır. Zira gerek ByRef kullanımı, gerek 
		dönüş değerinin dizi/collection olması sayesinde 
		fonksiyonlar pekala çoklu 
		değer döndürebilirler. Aşağıdaki örnekte çoklu değer olarak dizi 
		döndürüyoruz.</p>
		<pre class="brush:vb">
SSub diziata()
    'boyutsuz dizi yaratılır
    Dim d() As Integer
    
    d = dizidondur '6 değerli bir dizi döner, istediğimizi kullanırız
    Debug.Print d(3)
End Sub

Function dizidondur() As Integer() 'dönen değer tipi () ile biter ki dizi döndürdüğü anlaşılsın    
    Dim dizim(0 To 5) As Integer
    
    For x = 0 To 5
        dizim(x) = 10 * x
    Next x

    dizidondur = dizim
End Function				
</pre>
		<h3> Split ve join</h3>
		<h4> Split</h4>
		<p> Dizilerle kullandığımız çok faydalı iki fonksiyon vardır.</p>
		<p> Bunlardan ilki <span class="keywordler">Split</span> fonksiyonudur. 
		Bununla, belirli bir ayraçla birleşmiş olan bir <strong>String</strong> ifade parçalara 
		ayrılarak bir Variant içine veya klasik bir dizi içine aktarılır. 
		Gördüğünüz gibi bu da aslında bir nevi dizi tanımlama yöntemi 
		olmaktadır.</p>
		<p> Bu fonksiyon, genelde bir hücreden ";" veya "," ile ayrılmış 
		elemanları tek tek elde etmek için kullanılır.</p>
		<pre class="brush:vb">
Sub splitornek()
    Dim s As String
    s = "35516;37770;34234"

    Dim dizi As Variant 'veya Dim dizi() as String şeklinde boyutsuz dizi
    dizi = Split(s, ";")

    Debug.Print dizi(0) '35516 döner
End Sub			</pre>
		<p> Bu Split fonksiyonunu kullanarak bir UDF yazmıştım, ki 
		bu sayfayı yazarkenki en yüksek Excel sürümü olan Excel 2016 içinde bile bu fonksiyon 
		hala yerel olarak bulunmuyor. Bunu anlamak gerçekten zor, zira çok 
		ihtiyaç duyulan bir fonksiyon olduğunu düşünüyorum. Neyseki VBA ve UDF 
		var da başımızın çaresine bakabiliriyoruz. </p>
		<p> Bu 
		fonksiyon, bir metindeki belirtilen indeksteki kelimeyi seçiyor. Ayracı 
		default olarak " " yani boşluk belirledim, ancak isteyen farklı bir ayraç 
		da belirleyebiliyor. İşte bu fonksiyonum:</p>
		<pre class="brush:vb">
Function kelimesec(hucre As Range, kaçıncı As Byte, Optional ayrac As String = " ")
'normal bir cümlede ayrac boşluk olacağı için ayracı girmeye gerek yok, zaten default olarak " " atadım.
'ama içeriği mesela / ile ayrılmış bir hücre varsa 3.parametreyi / olarak girersiniz
  Dim kelimeler As Variant
  kelimeler = Split(hucre.Value2, ayrac)
  kelimesec = kelimeler(kaçıncı - 1)
End Function</pre>
		<h4> Join</h4>
		<p> Diğer faydalı fonksiyon da <span class="keywordler">
		Join</span> olup bir dizi içindeki elemanları, belirtilen ayraçla 
		birleştirip bir String ifade döndürür, yani Splitin tersi gibi çalışır.</p>
		<pre class="brush:vb">Sub joinornek()
Dim b As String
Dim v As Variant
v = Array(35516, 34433, 32335)

b = Join(v, ";")
Debug.Print b

End Sub</pre>
		<p>Bu fonksiyon sadece bir boyutlu dizilerde çalışır. 
		O yüzden bir 
		Range'den okunan ve 2 boyutlu olan bir <strong>Variant</strong> dizide işe yaramayacaktır. Ama 
		bunu yapmanın da bazı yolları var. Yukarda Variant dizileri klasik 
		dizilere çevirmeyi görmüştük, yine böyle yaparız ve gerisi kolay. </p>
		<h5>1.Yöntem:Variant diziyi normal diziye dönüştürmek</h5>
		<p>1 kolonlularda</p>
	<pre class="brush:vb">
Sub birkolonlu_variant_to_dizi()
Dim var As Variant
Dim dizi() As String
Dim kombine As String
Dim i As Integer

var = Range("B2:B10").Value2
ReDim dizi(1 To UBound(var, 1))

For i = 1 To UBound(var, 1)
    dizi(i) = var(i, 1)
Next i

kombine = Join(dizi, ", ")

MsgBox kombine

End Sub</pre>
		<p>2 kolonlularda</p>
	<pre class="brush:vb">
Sub çokkolonlu_variant_to_dizi()
Dim var As Variant
Dim dizi() As String
Dim kombine As String
Dim s As Integer, k As Integer, i As Integer

var = Range("B2:D10").Value2
ReDim dizi(1 To UBound(var, 1) * UBound(var, 2))

s = 0
For k = 1 To UBound(var, 2)
   For i = 1 To UBound(var, 1)
     dizi(i + s) = var(i, k)
   Next i
   s = s + UBound(var, 1) 'her boyut geçişinde tekrar başa dönmesin diye aratoplam alınır
Next k

kombine = Join(dizi, ", ")
MsgBox kombine
End Sub	</pre>
		<h5>2.yöntem:Transpoze(sadece tek kolonsa)</h5>
		<pre class=brush:vb>kombine = Join(WorksheetFunction.Transpose(Range("B2:B10").Value), ", ")</pre>
		<h3> Varmı kontrolü ve Index bulma</h3>
		<h4> Filtreleme</h4>
	<p> Dizilerle ilgili sık yapılan işlemlerden biri de, bir elemanın o 
		dizide bulunup bulunmadığı ve /veya bulunuyorsa da index numarasını 
		öğrenmek olacaktır.<p> Bu işlemlere geçmeden önce
		<span class="keywordler">Filter</span> fonksiyonundan bahsedelim. Bu 
		metod ile <strong>String </strong>tipli bir dizinin içinde belirli bir 
		metni arar ve bunları filtreler, eşleşmeyenleri hariç tutarız. (Burada 
		filtrelemekten kastım, elemek değil seçmektir.)<p> <strong>
		Syntax:Filter(dizi,aranan,[Dahilmihariçmi],[aramatipi])</strong><p> 
		<strong>dizi </strong>olarak tek boyutlu bir diziyi parametre olarak 
		veririz, bu nedenle range atanmış bir Variant dizi parametre olarak 
		verilemez.<p> 
		<strong>Dahilmihariçmi</strong> parametresi girilmez veya True girilirse 
		aranan kelimeyi içeren kelimeler aranır, False girilirse aranan kelimeyi içer<span style="text-decoration: underline">me</span>yen 
		elemanlar filtrelenir.<p> <strong>aramatipi</strong> girilmez veya <span>
		<strong>vbBinaryCompare</strong> girilirse küçük büyük harf duyarlılığı 
		gözetilerek arama yapılır, <strong>vbTextCompare</strong> 
		girilirse küçükbüyük harf duyarlılığı olmadan arama yapılır. İstisna:ş 
		ve Ş, ç ve Ç v.s aynı kabul edilirken bir tek İ ve i farklı algılanır.</span><p> 
		Bununla beraber bu metodu, aranan metnin tam eşleşme durumu için 
		kullanamazsınız. Yani "is" metnini aradığınızda "is, istanbul, mistik" 
		kelimelerinin tamamı gelir.<p> 
		Şimdi basit bir örnekle konuyu anlamaya çalışalım. Aşağıdaki dizide 
		içinde "kan" geçen isimleri filtrelemeye çalışıyoruz.<pre class="brush:vb">Sub basitfiltre()
Dim isimler As Variant
Dim filtreli As Variant

isimler = Array("volkan", "erhan", "hakan", "özkan", "meltem", "serkan")
filtreli = Filter(isimler, "kan", True, vbTextCompare)

For Each f In filtreli
   Debug.Print f
Next f

End Sub</pre>
	<p> Aşağıdaki örnekte 
	ise içinde "ankara" geçen tüm kelimeler aranır ve filtrelenir. Yanlız bu 
	sefer kaynak dizimiz 2 boyutlu ve önce onu tek boyutlu diziye çevirmemiz 
	gerekiyor.<pre class="brush:vb">Sub filtreornek()
Dim şubeler As Variant
Dim filtreli As Variant
Dim geçici() As String
Dim i As Integer

'çeşitli şubelerin olduğu alandan okuma yapılır(sırayla ankara, antalya, ankara, istanbul, sivas yazsın)
şubeler = Range("A1:A5").Value

'Variant ve 2 boyutlu olan şubeler dizisi klasik bir boyutlu bir diziye atanır
ReDim geçici(1 To UBound(şubeler, 1))
For i = 1 To UBound(şubeler, 1)
  geçici(i) = şubeler(i, 1)
Next i

filtreli = Filter(geçici, "ankara", True, vbTextCompare)

Set Hedef = Range("C1").Resize(UBound(filtreli) + 1, 1)
Hedef.Value = filtreli

End Sub</pre>
		<h4> Bir metin bir dizide var mı kontrolü?</h4>
		<p> Filter fonksiyonunu öğrendiğimize göre şimdi bunu daha spesifik bir 
		amaçla kullanabiliriz: <strong>Bir metnin bir dizide olup olmadığını öğrenmek 
		için</strong>. Bunun için birkaç yöntem bulunuyor.</p>
		<p> İlk yöntemde diziye filtre uygular ve elde ettiğimiz yeni dizinin 
		üst limitine bakarız. Eğer üst limit 0 veya daha üstü bir sayı ise aradığımız 
		metin dizide var demektir, -1 ise dizide bulunmuyor demektir.</p>
		<pre class="brush:vb">
Sub varmı1sub()
'1.Yöntem:Filter, tam eşleşme aranmıyor
Dim şubeler As Variant
Dim filtreli As Variant
Dim geçici() As String
Dim i As Integer

'çeşitli şubelerin olduğu alandan okuma yapılır(sırayla ankara, antalya, ankara, istanbul,sivas yazsın)
şubeler = Range("A1:A5").Value

'Variant ve 2 boyutlu olan şubeler dizisi klasik bir boyutlu bir diziye atanır
ReDim geçici(1 To UBound(şubeler, 1))
For i = 1 To UBound(şubeler, 1)
   geçici(i) = şubeler(i, 1)
Next i

filtreli = Filter(geçici, "ankara", True, vbTextCompare)

If UBound(filtreli)=-1 Then 
  MsgBox "aradığınız kelime dizide yok"
Else
  MsgBox "aradığınız kelime dizide var"
End If

End Sub</pre>
	<p>
	Bunu bir fonksiyon haline getirmek daha kullanışlı yapacaktır.</p>
	<pre class="brush:vb">
'1'in function hali
Function Varmı1(aranan As String, dizi As Variant) As Boolean
    Varmı1 = (UBound(Filter(dizi, aranan, True, vbTextCompare)) > -1)
End Function
'--------
    Sub test1()
        isimler = Array("volkan", "erhan", "hakan", "özkan", "meltem", "serkan")
        Debug.Print Varmı1("volk", isimler) 'True
        Debug.Print Varmı1("volkan", isimler) 'True
    End Sub
</pre>

<p>Tabi bu yöntemde tam eşleşme durumu sağlanmıyor, zira ilgili metni içeren 
		bir kelime var mı diye kontrol ediyor. Eğer ki tam eşleşme istiyorsak, 
		<strong>Match</strong> fonksiyonunu kullanmalıyız veya Filter dışında 
		çözüm sağlayan 2.yöntemi denemeliyiz. <strong>Bu yöntemde 1.yöntemin 
aksine sayısal değerleri de arayabiliyoruz</strong>. Bu yüzden hem sayıları hem 
stringleri kapsayacak şekilde Variant tipte bir aranan parametresi belirtiyoruz.</p>

<pre class="brush:vb">
Function Varmı2(aranan As Variant, dizi As Variant) As Boolean
'2.yöntem)Match, tam eşleşme sağlanıyor
On Error GoTo hata
    Varmı2 = Not IsError(WorksheetFunction.Match(aranan, dizi, 0))
Exit Function
hata:
    Varmı2 = False
End Function
'-----------
    Sub test2()
        isimler = Array("volkan", "erhan", "hakan", "özkan", "meltem", "serkan")
        Debug.Print Varmı2("volk", isimler) 'False
        Debug.Print Varmı2("volkan", isimler) 'True

        sayılar = Array(1, 2, 3)
        Debug.Print varmı3(1, sayılar) 'true
        Debug.Print varmı3(10, sayılar) 'false
    End Sub
</pre>
<p>3. yöntem, dizi elemanlarını <strong>Join</strong> ile birleştirip, yine bir String 
		fonksiyonu olan ve bir string içinde bir alt metin arayan <strong>InStr</strong> 
		fonksiyonu ile arama yapmak şeklinde olabilir, üstelik bu yöntem az önce 
		belirtildiği gibi tam eşleşme sağlar. Bunda da 2.yöntemde olduğu gibi 
sayısal değerleri de arayabiliyoruz, o yüzden fonknsiyonumuzu Variant 
tanımlıyoruz. </p>
	<p>Şimdi bu örneği direkt Function olarak yapalım.</p>
<pre class="brush:vb">
Function varmı3(aranan As Variant, dizi As Variant) As Boolean
ayraç = "-"
joinli = ayraç & Join(dizi, ayraç) 'tam eşleşmeyi sağlamak için başına da ilgili ayracı koyuyoruz
If InStr(1, joinli, ayraç & aranan & ayraç, vbTextCompare) Then
    varmı3 = True
Else
    varmı3 = False
End If
End Function
    Sub test3()
        isimler = Array("volkan", "erhan", "hakan", "özkan", "meltem", "serkan")
        Debug.Print varmı3("volk", isimler) 'False
        Debug.Print varmı3("volkan", isimler) 'True
        
        sayılar = Array(1, 2, 3)
        Debug.Print varmı3(1, sayılar) 'true
        Debug.Print varmı3(10, sayılar) 'false
    End Sub
</pre>
		<p>
		4. yöntem olarak, daha basit bir yol olan <strong>Dictionary</strong> kullanılabilir, 
		tabi kodun genel yapısı buna uygun ise. Bunu sonraki <a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">sayfalarda</a> göreceğiz.</p>
	<p>
		5.yöntem olarak
		<a href="https://www.geeksforgeeks.org/searching-algorithms/">arama 
		algoritmalarından</a> birini kullanarak kendi fonksiyonunuzu yazmayı 
		önerebilirim ama çok profesyonelleşmek istemediğiniz sürece bu sulara 
		girmenize gerek yoktur.</p>
		<h4>
		Var mı, varsa kaçta?</h4>
		<p>
		Yukarıdaki Match fonksiyonunu kullandığımız 2. yöntemde aslında ilgili 
		metnin kaçıncı sırada olduğunu buldurmaya yarayan bir fonksiyon kullanmış olduk. Eğer 
		bulamazsa hata döndürür, zaten biz de hata döndürüp döndürmediğine 
		bakıyorduk, döndürmüyorsa elemanı içeriyor diyorduk. Peki ya hata 
		döndürmüyorsa? Döndürdüğü şey sıra numarası olmaktadır.</p>
		<h3> Dizileri sıralama</h3>
		<p>
		Algoritma kavramına detaylı girdiğimizde karşımıza çıkan en temel 
		algoritmalardan biri de sıralama algoritmalarıdır.(Diğeri de hemen az 
		önce bahsettiğim arama 
		algoritmasıdır). </p>
	<p>
		Bu sıralama algoritmalarını anlamak önemli, zira dizilerle ilgili olarak 
		sıralama yapmayı sağlayan ne yerleşik bir fonksiyon/metod, ne de arama 
		yaparken kullandığımız gibi yardımcı yöntemler(Match, Join) 
		bulunmaktadır. Kendi fonksiyonumuz kendimizin yazması gerekmektedir. 
		</p>
	<p>
		Bunların da kendi içinde türleri vardır. Bubble Sort, Merge Sort, 
		Quick Sort gibi. Burada bunlardan ikisine bakacağız.
		<a href="https://www.mrexcel.com/forum/excel-questions/690718-vba-sort-array-numbers.html">
		Şu</a> ve
		<a href="https://stackoverflow.com/questions/152319/vba-array-sort-function">
		şu</a> sitelerde daha bol miktarda örnek bulunmakta,
		<a href="https://www.geeksforgeeks.org/sorting-algorithms/">şu</a> ve
		<a href="http://bilgisayarkavramlari.sadievrenseker.com/2008/08/09/siralama-algoritmalari-sorting-algorithms/">
		şu</a>(Türkçe) sitede 
		ise muazzam bir kaynak bulunmaktadır. Bu son sitelerden farklı yönteleri 
		öğrenip VBA kodlamasını yapmayı deneyebilirsiniz.</p>
	<h4>
		Bubble Sort</h4>
	<p>
		Bu yöntem en basit yöntem olup, aynı zamanda en yavaş çalışan yöntemdir 
		de. Gerçi küçük dizilerde bu yavaşlık problemi çok farkedilmez. Çalışma 
		yöntemi, tekrarlı bir şeklide komşu değerlerin sıralamasını yapmaktan 
		ibarettir. Aşağıdaki örnekte tüm aşamaların üzerinden tek tek geçeceğiz.</p>
	<p>
		Elimizdeki dizi 5,1,4,2,8 rakamlarından oluşsun.</p>
	<h5>
	1.aşama</h5>
	<p>
	<span style="text-decoration: underline">Önce&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
	Sonra&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>
	<br style="text-decoration: underline">(&nbsp;<b>5</b>&nbsp;<b>1</b>&nbsp;4 
	2 8 ) –&gt; (&nbsp;<span style="color: red"><b>1</b></span>&nbsp;<span style="color: red"><b>5</b></span>&nbsp;4 
	2 8 )<br>( 1&nbsp;<b>5</b>&nbsp;<b>4</b>&nbsp;2 
	8 ) –&gt;&nbsp; ( 1&nbsp;<span style="color: red"><b>4</b></span>&nbsp;<span style="color: red"><b>5</b></span>&nbsp;2 
	8 )<br>( 1 4&nbsp;<b>5</b>&nbsp;<b>2</b>&nbsp;8 
	) –&gt;&nbsp; ( 1 4&nbsp;<span style="color: red"><b>2</b></span>&nbsp;<span style="color: red"><b>5</b></span>&nbsp;8 
	)<br>( 1 4 2&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;) 
	–&gt; ( 1 4 2&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;) 'burada bi değişiklik olmaz, 
	sırası doğru</p>
	<h5>
	2.aşama</h5>
	<p>
	(&nbsp;<b>1</b>&nbsp;<b>4</b>&nbsp;2 
	5 8 ) –&gt; (&nbsp;<b>1</b>&nbsp;<b>4</b>&nbsp;2 
	5 8 )<br>( 1&nbsp;<b>4</b>&nbsp;<b>2</b>&nbsp;5 
	8 ) –&gt; ( 1&nbsp;<span style="color: red"><b>2</b></span>&nbsp;<span style="color: red"><b>4</b></span>&nbsp;5 
	8 )<br>( 1 2&nbsp;<b>4</b>&nbsp;<b>5</b>&nbsp;8 
	) –&gt; ( 1 2&nbsp;<b>4</b>&nbsp;<b>5</b>&nbsp;8 
	)<br>( 1 2 4&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;) 
	–&gt;&nbsp; ( 1 2 4&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;)<br>
	Şu anda sıralama tamam ancak bu aşamada 1 tane de olsa bir değişim olduğu 
	için algoritmanın devam etmesi lazım, ta ki hiç değişim olmayana kadar.</p>
	<h5>
	3.aşama</h5>
	<p>
	(&nbsp;<b>1</b>&nbsp;<b>2</b>&nbsp;4 
	5 8 ) –&gt; (&nbsp;<b>1</b>&nbsp;<b>2</b>&nbsp;4 
	5 8 )<br>( 1&nbsp;<b>2</b>&nbsp;<b>4</b>&nbsp;5 
	8 ) –&gt; ( 1&nbsp;<b>2</b>&nbsp;<b>4</b>&nbsp;5 
	8 )<br>( 1 2&nbsp;<b>4</b>&nbsp;<b>5</b>&nbsp;8 
	) –&gt; ( 1 2&nbsp;<b>4</b>&nbsp;<b>5</b>&nbsp;8 
	)<br>( 1 2 4&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;) 
	–&gt; ( 1 2 4&nbsp;<b>5</b>&nbsp;<b>8</b>&nbsp;)</p>
	<p>
		Evet 3.aşamada artık bir değişiklik yok ve algoritmamız durur. Şimdi bunun 
		kodunu görelim.</p>
	<pre class="brush:vb">
Function BubbleSort(arr) As Variant
  Dim geçici As Variant
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) &gt; arr(j) Then
        geçici = arr(i)
        arr(i) = arr(j)
        arr(j) = geçici
      End If
    Next j
  Next i
  
BubbleSort = arr
End Function
'-------
    Sub bubble_sort_test()
    Dim dizi As Variant
    Dim sıralı As Variant
    
    dizi = Array(5,1,4,2,8)
    sıralı = BubbleSort(dizi)

    Call Diziyazdır(sıralı)
    End Sub</pre>
	<h4>
	Quicksort</h4>
	<p>
	QuickSort(Hızlı sıralama) algoritması, böl ve fethet felsefesine dayanan bir yöntem 
	uygular. Bu anlamda Merge Sort algoritmasına benzer. </p>
	<p>
	Bu yöntemde, bir 
	referans seçilir(bu referans en sondaki, en baştaki olabilir, ortdaki veya 
	rasgele olabilir. Buna pivot veya dayanak noktası da denir. Bu referanstan 
	küçük olanlar sol tarafa, büyük olanlar sağ tarafa konulur ve&nbsp; daha 
	fazla yer değiştirme işlemi yapılamayana kadar bu işlem recursive olarak devam eder. 
	Bu konuda yukarıda bahsettiğim linklerde de bolca açıklama ve örnek 
	bulunmaktadır. </p>
	<p>
	Şimdi, algoritmanın icabettiği şekilde ben de bir kod hazırladım. Referans 
	noktasını "sondaki eleman" olarak belirledim. İşleyişin 
	aşama aşama nasıl olduğunu kod içinde bulabilirsiniz. Verdiğim linklerde 
	gerekli açıklamalar, videolarla birlikte sunulduğu için ben daha fazla 
	uzatmamak adına ilave açıklamada bulunmayacağım. Detayla ilgilenmiyorsanız 
	kodu doğrudan da kullanabilirsiniz.</p>
	<p>
	Kodumuz iki kısımdan oluşmakta. Önce parçalama kısmı sonra sıralama kısmı. Kodu 
	tek seferde çalıştırıp Immediate Windowda algoritmanın işleyişini aşama 
	aşama&nbsp; takip edebilirsiniz. </p>
	<pre class="brush:vb">
Function Parçala(ByRef dizi As Variant, ByVal low As Integer, ByVal high As Integer) As Integer
    referans = dizi(high)
    Debug.Print "Referansımız:" & referans & ".Bundan küçükler sola, büyükler sağa. referans ortalarına. Soldakiler önce ele alınacka, sağdakiler daha sonra."
    i = low - 1
    For J = low To high - 1
        If dizi(J) &lt;= referans Then
            i = i + 1
            Debug.Print i, J, dizi(J) & "&lt;= referans olduğu için " & dizi(i) & " ile " & dizi(J) & " yer değiştirecek." & dizi(i) & " yavaşça sağa kayaacak: " & Join(dizi, "-")
            geçici = dizi(i) '
            dizi(i) = dizi(J)
            dizi(J) = geçici
            'Call DiziBitişikyazdır(dizi, "yer değişim sonucunda:")
        Else
            Debug.Print i, J, dizi(J) & "&gt;" & "referans(" & referans & ") olduğu için yer değişme yok"
        End If
    Next J
    Call DiziBitişikyazdır(dizi, "for çıkışında dizinin durumu:")
    
    Debug.Print dizi(i + 1) & " ile " & dizi(high) & " yer değiştirecek, yani referans orataya girecek."
    geçici2 = dizi(i + 1)
    dizi(i + 1) = dizi(high) 'referansla. yani referansı ortaya koyuyuoruz
    dizi(high) = geçici2
    Call DiziBitişikyazdır(dizi, "Bu aşamanın sonunda dizimiz:")
    Debug.Print "-----------------------------------------" & vbNewLine
    
    Parçala = i + 1
End Function
'------------------------------------
Sub hızlısırala(ByRef dizi As Variant, ByVal low As Integer, ByVal high As Integer)
    If low &lt; high Then
        Pi = Parçala(dizi, low, high)
        hızlısırala dizi, low, Pi - 1
        hızlısırala dizi, Pi + 1, high
    End If
End Sub
'------------------------------------
'bu kodu çalıştırarak test ediyoruz
    Sub quick_sort_test() 
         karışıkdizi = Array(5, 1, 4, 2, 8)
        'karışıkdizi = Array(52, 3, 45, 32, 8, 19, 47, 26, 1, 100, 68, 24, 12, 9, 72, 55, 36)
        Call DiziBitişikyazdır(karışıkdizi, "İlk giriş:")
        hızlısırala karışıkdizi, 0, UBound(karışıkdizi) 'byref gittiği için sıralanıp gelir
        Call DiziBitişikyazdır(karışıkdizi, "Nihai sıralama sonucu:")
    End Sub
</pre>
	<h3>
	Diziler için Utility Modülü</h3>
	<p>
	Dizilerle ilgili olarak sık yapılan işlmleri, özellikle test/kontrol için 
	yapılan işlemleri bir modül içinde toparlalak iyi bir fikir olacaktır. 
	Bunlardan sıralamayı yukarıda görmüştük, diğerlerine birkaç örneği ise 
	aşağıda bulabilirsiniz.</p>
	<pre class="brush:vb">
Sub Diziyazdır(dizi As Variant)
    For Each d In dizi
        Debug.Print d
    Next d
End Sub
'------------------
Sub DiziBitişikyazdır(dizi As Variant, Optional başlık As String = "Birleşen dizi:")
bitişik = Join(dizi, "-")
Debug.Print başlık & bitişik & vbNewLine

End Sub
'-------------
Function ikidizibirleştir(dizi1 As Variant, dizi2 As Variant)
Dim geçici As Variant

adet = UBound(dizi1) + UBound(dizi2) + 2
ReDim geçici(adet - 1)

For i = LBound(dizi1) To UBound(dizi1)
    geçici(i) = dizi1(i)
Next i

For i = UBound(dizi1) + 1 To adet - 1
    geçici(i) = dizi2(i - 1 - UBound(dizi1))
Next i
'control için
DiziBitişikyazdır (geçici)

ikidizibirleştir = geçici
End Function
	
	</pre>
		<h3> Hatırlatma</h3>
		<p> Sayfanın başında belirttiğim gibi
		<a href="../Fasulye/NeNeredeNasil_Diziler.aspx">buraya</a> da bakmanızı 
		tavsiye ediyorum. </p>

	</div>
</asp:Content>
