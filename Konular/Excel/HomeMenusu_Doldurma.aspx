<%@ Page Title='HomeMenusu Filling' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>

<div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Home Menüsü'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Filling (Otomatik doldurma işlemleri)</h1>

	<p>Doldurma işlemleri üzerinde çok konuşmaya değer konular değildir ancak Excel 2013'le gelen efsanevi <strong>Flashfill</strong> özelliği ile konuşmaya değer hale geldi. O yüzden kısaca klasik doldurma işlemlerinden bahsedelim ve sonrasında Flashfill'e geçelim.</p>
    <h2 class="baslik">Klasik Doldurma işlemleri</h2>
    <div class="konu">
        <p>Aslında burda Fill menüsünden ziyade <span class="keywordler">
		Options&gt;Advanced</span> sekmesindeki <strong>Edit Custom List</strong>'ten 
		bahsetmek istiyorum. Zira Fill menüsünde çok kayda değer birşey yok, en 
		azından ben hiç kullanmıyorum.</p>
		<p>Bildiğiniz gibi bir hücreye Ocak yazıp sağa veya aşağı doğru 
		kaydırdığınızda bu Şubat, Mart diye devam etmektedir. Bunun sebebi, 
		ayların Excel'e bir <strong>Liste/Seri </strong>olarak tanıtılmış 
		olmasıdır. <strong>Edit Custom List </strong>(Advanced sekmesinin en 
		altına gitmeniz gerekiyor) butonuna bastığımızda bunu görebiliriz.</p>
		<p>
		<img src="/images/excelfill1.jpg"></p>
		<p>Biz de buraya kendi listemizi ekleyebiliriz. Bunu istersek manuel 
		girişle veya hazırda bulunan bir listeyi import ederek yapabiliriz. Biz 
		hazır listeyi import edelim. Diyelim ki kurumumuzun bölge isimlerine sık 
		sık ihtiyaç duyuyoruz. Liste tanımlayarak bunları ikide bir bölge 
		dosyasından almaktan kurtulmuş olacağız.</p>
		<p>
		<img src="/images/excelfill2.jpg"></p>
		<p>Bundan sonra bir hücreye Akdeniz yazıp aşağı/sağa sürüklersem sırayla 
		Batı Karadeniz, Doğu Anadolu diye yazmaya başlayacak.</p>
		<p><span class="dikkat">Dikkat</span>: Buraya sadece sayılardan oluşan 
		bir liste girilemez. Mesela bu bölgelerin bölge kodlarını giremeyiz. 
		Ayrıca Listeye girilen metinlerin toplam uzunluğu 255 karakter 
		olmalıdır. Yani burya bölge isimleri belki sığacaktır ama şube 
		isimlerinden bir liste yapmak pek olası görünmüyor.</p>
    </div>
    <h2 class="baslik">FlashFill</h2>
    <div class="konu">
    <p>Uzuuuunca bir listeniz var diyelim. Mesela Bir Ad Soyad listesi. Bunu Ad ve Soyad olarak ikiye bölmek istiyorsunuz.</p>
		<p>		<img alt="" height="186" src="/images/home_flashfill1.jpg" width="342"></p>
		<p>Önünüzde birkaç seçenek var(Tabiki tek tek elle yazmayı bir seçenek 
		olarak düşünmüyoruz :))</p>
		<ol>
			<li>Bir formül yazıp bunu aşağı doğru çekmek.Bu yöntemin sakıncası 2 isimli kişilerde formülün karmaşıklaşmasıdır. Karmaşık metin formülleri için 
			<a href="FormulasMenusuFonksiyonlar_MetinselFonksiyonlar.aspx">buraya</a> tıklayınız. <ul>
				<li>İsim için:<span class="keywordler">=LEFT(A2;FIND(" ";A2))</span></li>
				<li>Soyisim için:<span class="keywordler">=RIGHT(A2;LEN(A2)-FIND(" 
				";A2))</span></li>
			</ul>
			</li>
			<li>Bir UDF kullanmak(Bu en hızlı yöntemdir, ancak VBA bilmeyi 
			gerektirir)<ul>
<pre class="brush:vb">
Function kelimesec(hucre As Range, kaçıncı As Byte, Optional ayrac As String = " ")
    'normal bir cümlede ayrac boşluk olacğaı için ayracı girmeye gerek yok, zaten default olarak " " atadım.
    'ama mesela içeriği / ile ayrılmış bir hücre varsa 3.parametre / olarak girilir
    Dim kelimeler As Variant
    kelimeler = Split(hucre.Value2, ayrac)
    kelimesec = kelimeler(kaçıncı - 1)
End Function				
</pre>
			</ul>
			</li>
			<li>Benim burada anlatacağım ise çok daha basit ve pratik bir 
			yöntem:<strong>FlashFill</strong>(Excel 2013le birlikte aramıza 
			katıldı)
			<p>Hemen baştan belirteyim, Flashfill'in otomatik çalışması için
			<span class="keywordler">File&gt;Options&gt;Advanced&gt;Editing options</span> 
			altında <strong>Automatically Flash Fill</strong> seçeneğinin 
			işaretli olması lazım. Aksi halde manuel Flash Fill yapmanız 
			gerkeir, ki bu da oldukça kolaydır.(Klavyeden
			<span class="keywordler">Ctrl+E</span> kısayolu 
			ile veya <span class="keywordler">Home Menüsü&gt;Fill&gt;Flash Fill
			</span>komutu ile)</p>
				<p>Şimdi yukardaki resimde B2 hücresine "Ali" yazalım. B2 
				hücresine gelip "V" harfine basar basmaz, otomatik flash fill 
				yapılacağına dair aşağıdaki görüntü ortaya çıkar:</p>

				<img alt="" src="/images/home_flashfill2.jpg">
				<p>Enter'a basar basmaz da otomatik tamamlanır. Sonuç aşağıdaki gibi olacaktır:</p>

				<img alt="" src="/images/home_flashfill3.jpg">
				
				<p>Bu işlemin tersi için de yani farklı kolonlardaki isim ve soyisimi birleştirme işlemleri de aynı mantıkla yapılabilir.</p>
			<p>FlashFill'in çalışma şekli şöyledir:Siz B2'ye Ali yazdığınız 
			zaman, bunun etrafında içinde Ali olan bir hücre var mı diye 
			bakıyor, bulursa bir desen(pattern) oluşturuyor ve bu deseni diğer 
			hücrelere de uygular. Burda oluşturduğu desen şu: &quot;Ali Korkmaz metni içinde Ali, metnin ilk kelimesidir, o yüzden uygulanmak istenen şey, her metnin ilk kelimesini almaktır.&quot;</p>
<p>Bazı karmaşık işlemler Flashfill ile yapılamıyor. Böyle durumlarda yine metin formüllerini uygulamanız gerekebilir.</p>


			</li>
		</ol>
	</div>

</asp:Content>
