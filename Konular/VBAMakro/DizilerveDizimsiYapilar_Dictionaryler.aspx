<%@ Page Title='DizilerveDizimsiYapilar Dictionary' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Diziler ve Dizimsi Yapılar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Dictionary</h1>

<p> Dizilerden sonra en çok kullanılan toplu değer saklama aracı 
Dictionary'lerdir. Collectionlara benzerler, onlar gibi değerleri <strong>Key/Value</strong> 
ikilisi şeklinde tutarlar. Gerçi Collectionlarda ikili yapı yerine tekli yapı 
kullanımı daha yaygındır; ikili yapı gerektiğinde çoğunlukla Dictionary kullanılır. 
Bu kadar çok benzerlik gösterdikleri için bu sayfa boyunca yeri geldikçe 
Collectionlarla olan benzerlik ve farklılıklara değinilecek, en aşağıda da genel 
bir karşılaştırma yapılacaktır.</p>
	<h2 class="baslik">Genel Bakış</h2>
	<div class="konu">
	<h3> Ne, nasıl, nerede</h3>
	<p> Adı üzerinde, bu yapıları bir sözlük gibi kullanırız. Mesela çeşitli 
	kategorideki kelimeler 
	için Türkçe-İngilizce karşılıkları şeklinde bir liste oluşturabiliriz, veya 
	bölge kodu-bölge adı gibi lookup listeleri oluşturabiliriz.</p>
	<pre class="brush:vb">Key, Item
---- ----
bir, one
iki, two
.......

veya
5001, Akdeniz
5002, Başkent
.....</pre>
		<p>Diğer örnekleri şöyle sıralayabiliriz:</p>
		<ul>
			<li>Key:Ürün kodu, Item:ürün açıklaması(ürün yerine şube, bölge, 
			stok v.s birçok şey konulabilir)</li>
			<li>Key:Ürün kodu, Item:ürün alış/satış fiyatı</li>
			<li>Key:Ürün kodu, Item:ürün cirosu/stok adedi v.s</li>
			<li>Key:Müşteri kodu, Item:Telefon no</li>
			<li>Key:Kısaltma, Item:Kısaltmanın uzun/açık hali(CHP, Cumhuriyet 
			Halk Partisi gibi)</li>
			<li>vb.</li>
		</ul>
		<h3>Tanımlama şekilleri</h3>

	<p> Collectionlardan farklı olarak Dictionaryleri kullanmak için projemize <strong>
	Scripting.Runtime</strong> kütüphanesini(Library) eklemek gerekir.
	(<strong>Tools&gt;References&gt;Microsoft Scripting Runtime</strong>)</p>

		<p>Dictionaryler, birçok object türü gibi hem Late hem Early binding 
	şeklinde tanımlanabilirler. (Early ve Late binding hakkında detay bilgi ve 
	birbirlerine göre avantaj/dezavantajları için 
		<a href="Ileriseviyekonular_ObjelerDunyasi.aspx#Binding">buraya</a> tıklayınız)</p>
	<pre class="brush:vb">
'Early binding olarak tanımlanırsa
Dim sayılar As New Scripting.Dictionary
'veya
Dim sayılar As Scripting.Dictionary
Set sayılar = New Scripting.Dictionary 

'Late binding olarak tanımlanırsa
Dim sayılar As Object
Set sayılar = CreateObject("Scripting.Dictionary")
</pre>
		<p>Early Binding'in tek satır ve iki satır versiyonu arasındaki farkı 
		Collectionlarda anlattığımız için burada tekrar aynı detaya girmiyorum.</p>
		<h3>Property ve Metodları</h3>
	<h4>Eleman ekleme</h4>
	<p>Collectionlarda olduğu gibi eleman eklemek için <strong>Add</strong> 
	metodu kullanılır. Dictionarylerde zorunlu olan iki parametre vardır, bunlar 
	sırasıyla şöyledir:<strong> Key </strong>ve <strong>Item</strong>. 
	(Collectionlarda Key opsiyoneldir)</p>
		<p>Bunların ikisi de Variant tiptedirler, yani her değeri taşıyabilirler, 
		buna Dictionary dahil, yani Dictionary Dictionary'si diye bir kavram 
		teknik olarak mümkündür ve aşağıda da bir örneğini yapıyor olacağız. (Collectionlarda 
		Key'ler string olmak zorundaydı)</p>
	<p> Şimdi iki basit örnek yapalım, ilki sayıların Türkçe-İngilizce karşılığı, 
	ikincisi de lookup liste olarak, bölge kodları ve bölge adları olsun.</p>
	<pre class="brush:vb">Dim sayılar As New Scripting.Dictionary

sayılar.Add "bir", "one"
sayılar.Add "iki", "two"

Debug.Print sayılar("bir") 'one yazar</pre>
	<p>İkinci örneğimiz de şöyle:</p>
	<pre class="brush:vb">Dim bölgeler As New Scripting.Dictionary

bölgeler.Add 5001, "Akdeniz"
bölgeler.Add 5002, "Başkent"

Debug.Print bölgeler(5001)</pre>
		<p><strong>Key</strong> parametresi Collectionlarda olduğu gibi benzersiz bir parametre 
		olmalıdır, yani sadece bir kere kullanılmalıdır. Collectionlarda vurgu 
		Item'dadır, yani<strong> </strong>sözel ifade etmek gerekirse<strong>
		</strong>biz <strong>Item'</strong>ı depolarız ve ona istenirse Key adında bir isim veririz. Dictionarylerde ise vurgu 
		Key'dedir, Key ile ona karşılık gelen yani onun lookup'ı olan Item 
		birlikte depolanır. </p>
		<p>Key de Item de Variant tipli parametrelerdir demiştik ancak Key'in 
		bir istisnası var, dizi(array) olamaz. Item ise array dahil herşey olabilir.</p>
		<p><strong>Item</strong> <strong>propertysi/özelliği</strong>:Collectionlardan 
		farklı olarak Dictionarylerde eleman eklemenin farklı bir yolu daha 
		vardır: Doğrudan atama yöntemi(Implicit adding). Eğer Dictionary içinde olmayan bir 
		anahtara atama yapılmaya çalışılırsa onu doğrudan eklemiş oluruz. Bunun 
		için Item property'sini kullanırız, veya bu property düşürülerek de 
		yazılabilir.</p>
		<pre class="brush:vb">'madenler isimli dict içinde şuan sadece altın ve gümüş var olsun
madenler.Item("diamond")="elmas" 'madenler.Add "diamond", "elmas" ile aynı</pre>
		<p>Bu özellik, parametre olarak Key'i alır. Yani Key'i "diamond", Item'ı 
		"elmas"tır. Mesela bu elemanda Item'ın baş harfini büyük olarak değiştirmek için şu kodu 
		yazarız.</p>
<span>
		<pre class="brush:vb">madenler.Item("diamond")="Elmas"
'veya kısaca
madenler("diamond")="Elmas"</pre>
		<p><span class="dikkat">ÖNEMLİ</span>: Collectionlardaki aynı mantıkla, 
		elemanlara ters yönden erişim mümkün değildir. Yani nasıl ki 
		collectionlarda Key'ler uniqe'tir ve Item belirtilerek Key'e 
		ulaşamıyorduk, dictionarylerde de Item belirterek Key'lere ulaşamayız.<br>
</p>
		</span>
		<h4>Exists metodu ile "Varmı" kontrolü</h4>
		<p>Bir Key'in Dictionary'de olup olmadığını <strong>Exists</strong> metodu ile 
		kontrol ederiz. Varsa True, yoksa False döndürür. Genel kullanımı 
		"Eleman yoksa onu ekle" şeklindedir.</p>
	<pre class="brush:vb">
If Not dict.Exists("elma") Then dict.Add "elma", "apple"</pre>
		<p>Daha sadece olarak şöyle de diyebilirdik. Zira bu yöntemle, dictionary içinde "olmayan"
		bir elemanı doğrudan dictionary'ye ekliyorduk. </p>
		<pre class="brush:vb">dict("elma")= "apple"</pre>
		<p>
		<span>Ama "Varmı" kontrolünü yaptığımız sırada başka işlemler de 
		yapmamız gerekirse Exists kullanmalıyız.</span></p>
		<p>
		Collectionlarda doğrudan böyle bir metod yoktur. Bu işlem, hata kontrolü ile dolaylı 
		olarak yapılmaktadır.</p>
		<h4>
		Items metodu</h4>
		<p>
		Bu metod, Dictionary içindeki tüm itemları döndürür ve bunu 0 tabanlı bir dizi 
		olarak depolar.</p>
	<pre class="brush:vb">
meyveler=dict.Items 'elemanları diziye atadık
Debug.Print Join(meyveler,"-") 'dizideki elemanları - ile birleştirdik</pre>
		<p>
		Döngüyle tamamına erişebiliriz.</p>
<span>
		<pre class="brush:vb">For Each i In dict.Items
  Debug.Print i
Next i</pre>
		<p>Herhangi bir indeksteki Item'a ulaşmak için de kullanılır.</p>
		<pre class="brush:vb">Debug.Print dict.Items(0)</pre>
		</span>
		<h4>
		Key propertysi</h4>
		<p>
		Dictionary içindeki belli bir Key'in değerini değiştirmek için 
		kullanılır.(Key'in karşılık geldiği Item'ı değil. Bunun için Item 
		özelliğini kullanıyoruz)</p>
	<pre class="brush:vb">
dict.Add "elma", "apple"
dict.Key("elma") = "Elma"</pre>
		<p>
		Bu yukardaki gibi bir örnekteki tekil bir elemanı değiştirmekten ziyade 
		tüm Key'lerin önüne sabit bir ifade 
		eklemek gibi çoklu değişiklikler yapılması daha olasıdır, ki bunu da döngü ile yaparız.</p>
		<pre class="brush:vb">For Each k In dict.Keys
  dict.Key(k) = "M_" &amp; k 'tüm meyvererin önünde Meyvenin M'si ve _ kodyduk
Next k</pre>
		<p>Elemanlara erişimi ve döngü detaylarını aşağıda göreceğiz.</p>
		<p><span class="dikkat">Dikkat</span>:Bu özellik Write-Only olup sadece 
		Key'in değerini değiştirmek için&nbsp;kullanılır, Item'ı elde etmek için 
		kullanılmaz. Item'ı elde etmek için nasıl kullanacağımıza elemanlara 
		erişim bölümünden bakabilirsiniz.</p>
		<h4>
		Keys metodu</h4>
		<p>
		<span>Dictionary içindeki tüm Keyleri döndürür ve bunu 0 tabanlı bir dizi 
		olarak depolar. </span></p>
	<pre class="brush:vb">
meyveler=dict.Keys
Debug.Print Join(meyveler,"-")</pre>

		<p>Herhangi bir indeksteki Key'e ulaşmak için de kullanılır.</p>
		<pre class="brush:vb">Debug.Print dict.Keys(0) 'kısaca dict(0) da yazılabilir</pre>

		<h4>
		Remove ve RemoveAll ile eleman çıkarma</h4>
		<p>
		Belirtilen key'deki elemanı çıkarmak için <strong>Remove</strong>, tüm elamanları 
		çıkarmak için 
		yani Dictionary'yi boşaltmak için <strong>RemoveAll</strong> metodunu kullanırız.</p>
		<pre class="brush:vb">
meyveler.Remove "elma"
meyveler.RemoveAll</pre>
		<p>
		Collectionlarda RemoveAll yoktu, bunun yerine tüm Collection içinde dolaşıp 
		elemanları tek tek kaldırmak gerekiyordu veya yeni Collection atama veya 
		Nothing ataması yapmak gibi 
		dolaylı yollara başvuruluyordu.</p>
		<p>
		<strong>NOT</strong>:Nothing ataması veya New Dictionary ataması da 
		ilgili dictionary'nin içini boşaltır.</p>
<span>
		<pre class="brush:vb">
Set meyveler = Nothing
Set meyveler = New Dictionary</pre>
		<p>
		RemoveAll ve New Dictionary arasındaki ayrımı görmek için
		<a href="Ileriseviyekonular_PivotTableChartveSlicernesneleri.aspx">şu 
		sayfada</a> Pivot Tablolarla ilgili kısımda "<span>Birden çok fieldda 
		filtre uygulama" başlığı altındaki örneği inceleyin.</span></p>
		</span>
		<h4>
		CompareMode ile küçük/büyük harf duyarlılığı</h4>
		<p>
		Dictionaryler default olarak küçük/büyük harf ayrımına duyarlıdır(case-sensitive'dir), 
		yani "elma" ve "Elma" farklı olarak algılanır. Bu duyarlılığı kaldırmak için aşağıdaki kod yazılır.</p>
		<pre class="brush:vb">Dim dict As New Dictionary
dict.CompareMode = vbTextCompare 'veya numerik değer olarak 1

'tekrar case sensitive yapmak için şöyle yazılır
dict.CompareMode = vbBinaryCompare 'veya 0
</pre>

		<h4>
		Count ile eleman saymak</h4>
		<p>
		Collectiondaki gibi içerdeki toplam eleman sayısını verir.</p>

		<pre class="brush:vb">Debug.Print dict.Count</pre>
		<h3>
		Elemenlara tekil erişim 
		ve Dictionary içinde dolaşma</h3>
		<p>
		Itemlara erişim ile Key'lere erişim bazen kafa karışıtırıcı 
		olabilmektedir. Bu kısımda bunların detaylarına değinmeye çalışacağım.</p>
		<p>
		Collectionlarda olduğunun aksine Dictionary'lere doğrudan Index numarası 
		ile ulaşılmaz, zira Index diye birşey yoktur. Hatırlayacak olursak Collectionlar sıralı bir yapıya 
		sahipken Dictionary'lerde sıra yoktur. Sonuç olarak, Dictionary'lerde 
		elemanlara erişim onun Key'i aracalığı ile olur, bu da ya <strong>Indexli Key 
		</strong>belirterek 
		ya da doğrudan <strong>Key'in kendisi</strong>(Stringse) yazılarak olmaktadır. 
		Indexli Key ile 
		ulaştığımız şey Key'in kendisi iken, Key'in kendisini yazarak eriştiğimiz şey 
		ise bu 
		Key'in lookup değeridir. Collectionlarda Index ile Item'a ulaşıyorduk, ki 
		buna Key adı verilerek de ulaşılabilir demiştik, Dictionarylerde ise 
		Indexli key 
		vererek Item'a ulaşıyoruz.</p>
		<p>
		<span class="dikkat">Dİkkat</span>:Key ismini <span class="keywordler">Key</span> propertysi ile 
		kullanamayız. Zira bu property write-only'dir, yani sadece Key'in 
		değerini değiştirmek için kullanılır.</p>
		<p>
		Farkındayım, bütün bunlar şuan çok karışık geliyor olabilir. Aşağıdaki 
		örnekler biraz daha aydınlanmanıza yarayacaktır. Biraz aşağıda tüm 
		bunları derleyip toplayan bir örnek ve bir tablo daha göreceksiniz. 
		Ondan sonra kendi 
		örneklerinizi de yapınca konu iyice pekişecektir. </p>
		<pre class="brush:vb">Dim dict As New Dictionary
dict.Add "elma", "apple"

'Tekil elemana read erişimi
Debug.Print dict("elma") 'Key'in kendisi ile Itema erişim. apple değerini verir
Debug.Print dict.Keys(0) 'Indeksli Key ile Key'e erişim. elma değerini verir 
Debug.Print dict.Key("elma") 'Hata verir. Çünkü Key propertysi write-onlydir.

'Tekil elemanlara write erişimi
dict.Key("elma")="Elma" 'Key'in değeri değişti, onun lookupı olan Itemın değil. Yani elme Elme oldu.
dict("elma")="apple" 'Itemın değerini değişti. Eleman yoksa Implicit ekleme olur. Yani elma, apple ikilisi eklenir
dict.Item("apple")="Apple" 'Item'ın değerini değiştirdik.</pre>
		<p>Aşağıda ise erişim yöntemleriyle ilgili küçük bir collection/dictionary 
		karşılaştırması bulunuyor.</p>
		<pre class="brush:vb">Dim col As New Collection
Dim dict As New Scripting.Dictionary

'Collection örneği
col.Add "Elma" '1.index
col.Add "Armut" '2.index
col.Add "Erik" '3.index

Debug.Print col(2) 'veya col.Item(2). Armut yazar

'Key'li Coll örneği. ilk yazılan Item, ikincisi Key'dir
col.Add "Elma", "Apple"
col.Add "Armut", "Pear"
col.Add "Erik", "Plum"

Debug.Print col("Plum") 'Erik yazar
Debug.Print col("Erik") 'Hata verir. Erişim Item'la olmaz,

'Dictionary örneği. ilk yazılan Key, ikincisi Item'dır
dict.Add "Elma", "Apple"
dict.Add "Armut", "Pear"
dict.Add "Erik", "Plum"

Debug.Print dict("Armut") 'Pear yazar
Debug.Print dict.Key("Armut") 'Hata verir. Çünkü Key özelliği Write-Only</pre>
		<h4>Tüm elemanları dolaşma</h4>
		<p>Dictionary'lerin tüm elemanları dolaşmak için tüm dizimsi yapılarda 
		olduğu gibi For Next döngülerini kullanıyoruz.</p>
		<p>For Each döngüsünü hem Early Binding hem Late Binding için kullanabilirken, 
		For Next döngüsünü sadece Early Binding tanımlama yapıldığında kullanabilirz.(Binding 
		çeşitleri hakkında bilgi için
		<a href="Ileriseviyekonular_ObjelerDunyasi.aspx#Binding">tıklayınız</a>)</p>
		<pre class="brush:vb">Dim k As Variant
For Each k In dict.Keys
   Debug.Print k, dict(k)
Next k</pre>
		<p>
<strong>For Each k In dict.Keys</strong> satırını <strong>For Each k In dict</strong> şeklinde de 
		yazabilirdik. Yani sadece Dictionary'nin adını yazmak onun Keys propertysine 
		bakacağız diye algılanır. Collectionlarda böyle bi yazımla Items 
		kastedilir.</p>
		<p>Klasik For döngüsünde ise dikkat edilecek husus, başlangıcın 0'dan 
		başlaması, bitiş indeksinin de eleman sayısı - 1 olmasıdır.</p>
	<pre class="brush:vb">
For i=0 To dict.Count-1
   Debug.Print dict.Keys(i), dict.Items(i)
Next i</pre>
		<p>
		Şimdi başka bir örneğe bakalım</p>
		<pre class="brush:vb">
Dim dict As New Scripting.Dictionary

dict.Add Key:="Apple", Item:="Elma"
dict.Add Key:="Orange", Item:="Portakal"
dict.Add Key:="Plum", Item:="Erik"

'Keylerin değerini değiştirebiliyoruz, yukardaki örnekte tek elemanın değerini 
'değiştirmiştik, şimdi tüm elemaların başına "A_" koyuyoruz.
For Each k In dict.Keys
   dict.Key(k) = "A_" &amp; k
Next k

'tüm keyler ve bunların lookup değeri olan itemları yazdırıyoruz
For Each k In dict.Keys
   Debug.Print k, dict(k)
Next k</pre>
		<p>
		Şimdi bi de 
		iki kolonlu bir listeyi döngüsel olarak Dictionary'ye ekleyelim, ama 
		eklerken Varmı diye de kontrol edelim. Bu liste, Şube kodu ve şube 
		adından oluşan bir liste olabilir.</p>
		<pre class="brush:vb">
Dim dict As Object 'Late binding ile yaratıyoruz
Set dict = CreateObject("Scripting.Dictionary")
 
Dim anahtar, deger 'tip belirtmeye gerek yok, Varianttırlar

Do
    anahtar = ActiveCell.Value
    deger = ActiveCell.Offset(0, 1).Value

    If Not dict.Exists(anahtar) Then
        dict.Add anahtar, deger
    End If
    
    ActiveCell.Offset(1, 0).Select
Loop Until IsEmpty(ActiveCell)

Debug.Print dict.Count</pre>

		<h3>Key, Keys, Item ve Items birlikte kullanımı</h3>
		<p>Aşağıdaki iki örnek ile tüm bu öğrendiklerimizi pekiştirelim.</p>
		<pre class="brush:vb">
Sub foreach_in_dict()
Dim madenler As New Scripting.Dictionary

madenler.Add "gold", "altın"
madenler.Add "iron", "demir"
madenler.Item("diamond") = "elmas" 'madenler.Add "diamond", "elmas" ile aynı
madenler("cupper") = "bakır" 'üsttekinin kısa yöntemi

Debug.Print "--------sadece key---------"
For Each K In madenler.Keys 'Keys yazmasak da olur
    Debug.Print K
Next K

Debug.Print "--------sadece item---------"
For Each i In madenler.Items
    Debug.Print i
Next i

Debug.Print "------Keysden giderek ikili yazım----------"
For Each K In madenler.Keys
    Debug.Print K, madenler(K)
Next K

Debug.Print "--------Items'dan giderek ikili yazım olmaz---------"
For Each i In madenler.Items
    Debug.Print i, madenler.Item(i) 'Item belirterek key'lere ulaşılamaz
Next i

End Sub		
		</pre>
		<p>Bu örnek ise pekiştirme için çok daha kuvvetli bir örnek.</p>
		<pre class="brush:vb">
Sub key_değiştirme()
Dim iller As New Dictionary

iller.Add "01", "adana"
iller.Add "02", "adıyaman"
iller.Add "03", "afyon"

iller.Item(0) = "Adana" 'Key'i 0, Item'ı Adana olan yeni bir eleman ekler
iller.Item(0) = "İstanbul" 'Var olan bir kayıt olduğu için update yapar
iller.Items(0) = "hey" 'Items ile atama değil okuma yapılır, bu satır etkisizdir
Debug.Print iller.Item(0) 'Key'i 0 olan item okunur, yani İstanbul
Debug.Print iller.Items(0) 'Item indeksi 0 olan kayıt okunur

'iller.Key("06") = "Ankara" 'hata: olmayan bir key'e erişemeyiz
iller.Key("01") = "001" 'Key'in kendisini değiştirdik, write-only
iller.Keys(0) = "hey" 'item gibi etkisiz. keys sadece okumada kullanılr
Debug.Print iller.Keys(0) 'hey değil 001 yazar
'Debug.Print iller.Key(0) 'hata alrız, çünkü write-only

'şimdi baştan bir dolaşaım
Debug.Print "--------değişm öncesi------"
For Each K In iller.Keys
    Debug.Print K, iller(K)
Next K

'şimdi de hem keyleri hem itemları aynı anda değiştirelim
For Each K In iller.Keys
    YK = Val(K)
    iller.Key(K) = YK
    iller.Item(YK) = UCase(iller.Item(YK))
Next K

'şimdi dğeişklikleri kontrol edelim
Debug.Print "--------değişm sonrası------"
For Each K In iller.Keys
    Debug.Print K, iller(K)
Next K
End Sub	</pre>
		<h4>
		Nihai Özet Tablo</h4>
		<p>
		Aşağıdaki tablo ile de yukarıdaki örnekleri bir tablo şeklinde 
		görüyoruz.</p>
		<p>
		İller isimli bir dictionary'miz olduğunu ve 01-Adana ile 
		33-İçel arasındaki kayıtların eklendiğini düşünün. Buna göre;</p>
		<p>
		<table class="alterantelitable">
			<tr >
				<th>Üye</th>
				<th>Ekleme</th>
				<th>Erişim</th>
				<th>Update</th>
			</tr>
			<tr>
				<td>Item() propertysi</td>
				<td>Olmayan kayıt key ile eklenir.Ör: Item(«34»)=«istanbul»</td>
				<td>Dict içinde bulunan bir Item, Key kullanılarak okunur.
				<strong>Erişilen şey Item'dır.</strong> Ör: Item(«34»)<font>--&gt;istanbul</font></td>
				<td>Dict içinde bulunan bir Item, Key kullanılarak değiştirilir. Ör:Item(«34»)=«<strong>İ</strong>stanbul»</td>
			</tr>
			<tr>
				<td>Item<span style="color: red"><strong>s</strong></span>() metod</td>
				<td>Ekleme yapılamaz</td>
				<td><span style="color: red"><strong>Indeks</strong></span> no ile erişilir.
				<strong>Erişilen şey Item'dır</strong>. Ör: Items(0)--&gt;<font>adana</font></td>
				<td>Etkisizdir. Update yapılamaz.</td>
			</tr>
			<tr height="20">
				<td height="20">Key() propertysi<br><strong>(write-only)</strong></td>
				<td>Ekleme yapılamaz. Kullanılırsa hata alınır</td>
				<td>Olmayan key’e erişemeyiz. Olanın ise içeriğini değiştiririz.
				<strong>Erişilen şey Key'dir.</strong></td>
				<td>Update yapılamaz</td>
			</tr>
			<tr>
				<td>Key<span style="color: red"><strong>s</strong></span>() metod</td>
				<td>Ekleme yapılamaz</td>
				<td><span style="color: red"><strong>Indeks</strong></span> no ile erişilir.
				<strong>Erişilen şey Key'dir.</strong> Ör: Keys(0)--&gt;<font> «01»</font></td>
				<td>Update yapılamaz</td>
			</tr>
		</table>
		<br>
</p>
	</div>
		<h2 class="baslik">Kıyaslamalar ve ileri örnekler</h2>
		<div class="konu">
	<h3>Collectionların ve Dictionarylerin karşılaştırılması</h3>
			<p>Bu iki yapının benzerlikleri, farklılıkları ve birbirine göre 
			avantaj/dezavantajları bulunmaktadır. "Şu daha iyidir" diye 
			doğrudan bir söylem doğru değildir. Her araçta olduğu gibi, o an 
			ihtiyacımızı en iyi hangisi görüyorsa onu kullanmamız gerekmektedir. 
			Ben burada bir karşılaştırma vereceğim, kararı siz verin. Tabiki benim 
			de naçizane bazı tavsiye ve yönlendirmelerim olmayacak değil.</p>
			<ul>
				<li>Collectionlar VBA içinde yerel olarak bulunurken, Dictionary'leri 
				kullanmak için bunu ya reference olarak eklemeli ya da 
				CreateObject ile Late Binding şekilde yaratmalıyız.</li>
				<li>Dictioanry'lerde Key de Item da hem okunabilir hem yazılabilirdir. 
				Ancak Collectionlarda Item'ın değerini değiştiremezsiniz, yani read-onlydir. Bunu yapmak için önce onu kaldırmalı sonra yeni değerle 
				tekrar eklemeniz gerekir. Ayrıca Collectionlarda Keyler ne read-only ne write-onyldir, 
				yani değer atanamadıkları gibi elde edilemezler de; sadece ilgili 
				Item'a ulaşmada kullanılırlar.</li>
				<li>Collectionlar sıralıdır, Dictionarylerde sıra yoktur. Bu 
				yüzden Collectionlar'a indeks numarası ile erişebilirken 
				Dictionary'lerde indeks ile erişilemez. Bununla beraber 
				Dictionary'lerin Keys ve Items metodları bunları 0 tabanlı bir 
				dizi olarak döndürür, yani Key ve Item'lardan oluşan bir dizi 0 
				indekslidir. Ancak Collectionlarda indeks 1'den başlar.</li>
                <li>Her iki yapıda da elemanlarda dolaşmak için For 
				Each yapısı kullanılır. Direkt ilgili nesnenin adı verilerek 
				dolaşılmaya çalışıldığında, Collectionda itemlarda 
				dolaşılırken Dictionarylerde Keylerde dolaşılır.</li>
				<li>Collectionları diziye atamak için döngüsel bir yapıyı içeren 
				birkaç satırlık koda 
				ihtiyaç duyulurken Dictionarylerde <span class="keywordler">Items</span> ve 
				<span class="keywordler">Keys</span> metodları bize doğrudan 
				dizi verirler.</li>
				<li>Dictionarylerde eleman eklemek için <span class="keywordler">Add
				</span>metoduna ek olarak implicit(üstü 
				kapalı) ekleme yöntemi de varken, Collectionlarda 
				yanlızca Add metodu kullanılır.</li>
				<li>Dictionarylerde tüm elemanlar tek seferde(<span class="keywordler">RemoveAll</span> ile) 
				silinebilirken Collectionlarda dolaylı yöntemler izlenir.</li>
				<li>Keyler her ikisinde de benzersiz olmalıdır.</li>
				<li>Keyler Dictionaryde zorunlu iken Collectionda seçimliktir.</li>
				<li>Keyler Collectionlarda String olmalıyken, Dictionarylerde ise dizi 
				dışında herşey olabilirler.</li>
				<li>Dictionarylerde bir elamanın var olup olmadığı 
				<span class="keywordler">Exists</span> metodu ile kontrol 
				edilebilirken Collectionlarda bu kontrol için birkaç satırlık 
				kod yazmak gerekir.</li>
				<li>Collectionlar küçük/büyük harf ayrımına duyarlı değilken
<span>Dictionaryler</span> duyarlıdır, istenirse duyarsız hale getirilebilir.</li>
				<li>Genel olarak bakıldığında, Dictionaryler Collectionlara göre daha hızlıdır</li>
			</ul>
			<table class="alterantelitable">
				<tr>
					<th> &nbsp;</th>
					<th>Dictionary</th>
					<th>Collection</th>
				</tr>
				<tr>
					<td>Parametreler</td>
					<td>
					<p>İkisi de zorunlu</p>
					</td>
					<td>
					<p>Sadece Item zorunlu</p>
					</td>
				</tr>
				<tr>
					<td>Vurgu</td>
					<td>
					<p><strong>Key</strong>, Item</p>
					</td>
					<td>
					<p><strong>Item</strong>, (Key)</p>
					</td>
				</tr>
				<tr>
					<td>Erişim</td>
					<td>
					<p>Item(«key adı»)--&gt;Item<br>Keys(indeks)--&gt;Key</p>
					</td>
					<td>
					<p>İndeks, Key--&gt;Sadece item elde edilir. Key elde 
					edilemez</p>
					</td>
				</tr>
			</table>
			<p>Bu farkları bir kısmını küçük bir kodda inceleyelim:</p>
			<pre class="brush:vb">
Sub coll_vs_dict()
Dim dict As New Dictionary 'library
Dim coll As New Collection 'no library

dict.Add "yüz", "hundred"
coll.Add "hundred", "yüz"

Debug.Print dict.Keys(0), dict.Items(0) 'iki değer de elde edilebilir
Debug.Print coll(1), coll("yüz") 'Key yani "yüz" değeri elde edilemez

dict.Item("yüz") = 100 'key'in lookup değeri olan item'ı değiştirebiliriz
Debug.Print dict.Keys(0), dict.Items(0)
coll(1) = 100 'hata. collectionlar readonlydir, itemlar dğeiştirilemez, keylere zaten ulaşamıyoruz bile

Debug.Print dict(0) 'dictionaryde indeks yoktur
Debug.Print coll(1) 'collectionlarda indeks var ve 1'den başlar
End Sub</pre>
			<p>Büyük üstat Cpearson der ki:</p>

			<blockquote><em>Her iki obje de benzer datayı gruplamak için çok faydalıdır ancak 
			herşey eşit olduğunda ben Dictionary kullanmayı tercih ederim. </em> </blockquote>
			<p>Gerekçe olarak da yukarda belirttiğim maddelerden bazılarını dile 
			getirmiş. </p>
			<p>Benim de naçizane bir tavsiyem var: Eğer 
			sadece arka arkaya birşeyler eklemek 
istiyorsanız ve kümeden birşeyler çıkarma veya Varmı kontrolü gibi şeyler 
			yapmayacaksanız Collection kullanın, hem de Key'siz haliyle. Ama bir 
			lookup değeri de olacaksa Key ve Item ikilisine yani bir 
			Dictionary'ye ihtiyacınız var demektir.</p>

<h3>İçiçe Dictionary(Dictionary of Dictionary)</h3>
			<p>Dizi Dizisi ve Collection Collectionı gibi Dictionarylerin de 
			içiçe geçmiş formları vardır. Hatta yeri gelir, Dictionary of 
			Collection, Collection of Dictionary veye Array of Collection gibi 
			çapraz formlar da kullanmamız gerekebilir.</p>
			<p>Benim şahsen çok ihtiyacım olmadı ancak internette bol miktarda 
			örnek bulunmaktadır. Mesela şu
			<a href="https://www.experts-exchange.com/articles/3391/Using-the-Dictionary-Class-in-VBA.html">
			sayfada</a> hem Dictionaryler hakkında çok faydalı bilgiler hem de 
			içiçe Dictionary dahil birçok örneği bulabilirsiniz.</p>
			<p>Bununla birlikte gözünüzde canlanması için aşağıdaki gibi bir 
			örnek kod yazabiliriz.</p>
			<pre class="brush:vb">Sub dictofdict()
Dim m As New Scripting.Dictionary
Dim s As New Scripting.Dictionary
Dim dict As New Scripting.Dictionary

m.Add "elma", "apple"
m.Add "erik", "plum"

s.Add "salatalık", "cucumber"
s.Add "domates", "tomato"

dict.Add "meyveler", m
dict.Add "sebzeler", s

Debug.Print dict("meyveler")("elma") 'veya dict("meyveler").Item("elma")

End Sub</pre>
			<p>Bu arada&nbsp; ilk başta kulağa sanki Dictionary of Dictionary ile 
			çözülebilirmiş gibi gelen bir problemi ben aşağıdaki gibi Dicitonary 
			ve 3 boyutlu dizi ile hallettim. Biraz üzerine düşününce farklı 
			çözümler de üretilebiliyor. Mesela listenizi farklı bir 
			formata getirip arkasından Dictionary tipli bir dizi ile de sorun 
			çözülebilir.</p>
<span>

<h4>Dictionary ve Dizi bir arada</h4>
			<p>Aşağıdaki gibi bir listemiz var ve bunu bir alttaki resimdeki 
			hale getirmek istiyoruz. Bu liste hergün güncellenen bir personel 
			dosyası ve bir alttaki hale gelmeli, yoksa burdan beslenen 
			formüllerde hatalar olacağı gibi bölgelere giden otomatik maillerde 
			yanlış kişilere yanlış mailler gidebilir.
<span>

			Böyle bir listenin hergün manuel bir şekilde işlenmesi de 
			oldukça zahmetli olurdu. O yüzden aşağıdaki gibi bir kod yazmalıyız.
			</span>
			</p>
			<p>
			<img src="/images/vbadict5.jpg"></p>
			<p>
			Getirmek istediğim hal ise şu. Bu hale geldikten sonra burdan 
			beslenen birçok lookup formülü var.</p>
			<p>
			<img src="/images/vbadict6.jpg"></p>
			<p>
			<strong>İzlenecek yol</strong>:Blg ve Sgm isminde iki Dictionary 
			tanımlarız. Bir de sicilleri atayacağımız bir dizi. Dizimiz çok 
			boyutlu olabileceği gibi içiçe dizi şeklinde de olabilir. Ben burada 
			çok boyutlu dizi yöntemini seçtim.</p>
			</span>
			<pre class="brush:vb">
Sub dictvedizi()
Dim Blg As New Scripting.Dictionary
Dim Sgm As New Scripting.Dictionary
Dim Siciller() As String

ReDim Siciller(0 To 2, 0 To 3, 0 To 3) 'boyutlar sırayla şöyle:bölge sayısı, segment  sayısı, kişi sayısı

Set alanBolge = Range(Range("a2"), Range("a2").End(xlDown))
Set alanSegment = Range(Range("c2"), Range("c2").End(xlDown))

i = 0
For Each d In alanBolge
    If Not Blg.Exists(d.Value) Then
        Blg.Add Key:=d.Value, Item:=i
        i = i + 1
    End If
Next d

k = 0
For Each d In alanSegment
    If Not Sgm.Exists(d.Value) Then
        Sgm.Add Key:=d.Value, Item:=k
        k = k + 1
    End If
Next d

'data okuma
For Each d In alanBolge
    Siciller(Blg(d.Value), Sgm(d.Offset(0, 2).Value), dolusay(Siciller, Blg(d.Value), Sgm(d.Offset(0, 2).Value)) + 1) = d.Offset(0, 1).Value
Next d

'hedefe yazma
For i = 0 To Blg.Count - 1
    Cells(i + 2, 5).Value = Blg.Keys(i)
Next i

For i = 0 To Sgm.Count - 1
    Cells(1, i + 6).Value = Sgm.Keys(i)
Next i


For x = 1 To 4
    For y = 1 To 3
        Set h = Cells(1 + y, 5 + x)
        h.Select
        h.Value = sonucgetir(Siciller, Blg(h.Offset(0, -x).Value), Sgm(h.Offset(-y, 0).Value))
    Next y
Next x


End Sub
Public Function dolusay(ByVal Data As Variant, ByVal i1 As Integer, ByVal i2 As Integer) As Integer
    Dim Count As Integer
    Count = 0
    
    For j = 0 To UBound(Data, 3) - 1
        If Len(Data(i1, i2, j)) > 0 Then
            Count = Count + 1
        End If
    Next j
    dolusay = Count
End Function
Public Function sonucgetir(ByVal Data As Variant, ByVal i1 As Integer, ByVal i2 As Integer) As String
    sonucgetir = ""
    For i = 0 To UBound(Data, 3)
        If Len(Data(i1, i2, i)) > 0 Then
            x = Data(i1, i2, i) & ";" & x
            sonucgetir = Left(x, Len(x) - 1)
        End If
    Next i
End Function</pre>
			<p>
			Kodu biraz inceleyin, başka türlü nasıl yapılabilrdi, onu düşünün. 
			Şükür ki, gerek Excel'de gerek VBA'de bir işi yapmanın birden çok 
			yolu olabilmektedir. Artık aklınıza hangisi gelirse onun üzerine 
			yoğunlaşın.	</p>

</div>

<h2 class="baslik">Çeşitli örnekler</h2>

	<div class="konu">
	
	<h4 class="baslik">Her değer için min/maks değeri alıp depolama</h4>
			<div>
			<p>Kişi isimlerinin birkaç kez geçtiği yerde herkesin en büyük satışını aldırma, herkese 
			ait en küçük tarih buldurma gibi örnekler de dictionarylerin pratik uygulamaları 
			arasındadır. Burdaki mantık şu şekilde işler. For döngüsü ile tüm değerler 
			taranır ve sırayla eklenir, ancak ekeme yaplırken Exists ile "daha önce eklenmiş 
			mi kontrolü" yapılır. Tabiki datanın ya manuel ya da kod içinde sıralanmış olması gerekir.</p>

				<p>Aşağıdaki örnekte, belirli sicil numaralı kişlerin belirli şubelere 
				başlama tarihleri var. Bir kişi zaman içinde bir şubeden başka şubeye tayin olabilmekte, o yüzden bazı 
				kişilerin birden fazla satırda geçtiğini anlamak zor olmayacaktır.
				Biz burada bir kişinin bankaya en erken giriş tarihin bulmaya çalışacağız. 
				Ör:35516 için 14.09.2014 tarihini bulmalıyız, 38541 için de 28.05.2015.</p>
				<p><img src="/images/vbadict3.jpg"></p>
				<p>Kodu uzatmamak adına listeyi manuel olarak sıralayalım</p>
				<p><img src="/images/vbadict4.jpg"></p>
				<p>Kodumuz ise aşağıdaki gibi olacaktır.</p>
				<pre class="brush:vb">
			Sub enbuyuktarih()
			Dim st As New Scripting.Dictionary
			Dim alan As Range, a As Range
			
			Set alan = Range("A2:A7")
			For Each a In alan
			    If Not st.Exists(a.Value2) Then
			        st.Add a.Value2, a.Offset(0, 2).Value
			    End If
			Next a
			
			For Each s In st
			    Debug.Print s, st(s)
			Next s
			End Sub	</pre>
				<p>Eğer ki bu ilk şubenin hangisi olduğunu da öğrenmek isteseydik farklı 
				bişeyler daha yazmamız gerekirdi. Ben şube kodunu Collection'a ekleyerek bulma 
				yöntemini denedim. Dizi kullanarak da çözülebilirdi.</p>
				<pre class="brush:vb">
			Sub enbuyuktarih2()
			Dim st As New Scripting.Dictionary
			Dim col As New Collection
			Dim alan As Range, a As Range
			
			Set alan = Range("A2:A7")
			For Each a In alan
			    If Not st.Exists(a.Value2) Then
			        st.Add a.Value2, a.Offset(0, 2).Value
			        col.Add a.Offset(0, 1).Value, CStr(a.Value) 'Keyler collectionlarda  string olmalı
			    End If
			Next a
			
			For Each s In st
			    Debug.Print s, st(s), col(CStr(s))
			Next s
			End Sub	
				
				</pre>
			</div>
	
	
	<h4 class="baslik">Dictionary tipli bir dizi</h4>
				<div>
				<p>Gerçek bir sözlük uygulaması(İngilizce dışındakiler uydurmadır :))</p>
					<p><img src="/images/vbadictionarysozluk.jpg"></p>
				<pre class="brush:vb">
				Sub arraydict()
				Dim dict(2) As New Scripting.Dictionary
				dict(0).Add "ekmek", "bread" 'ingilizce
				dict(1).Add "ekmek", "brot" 'almanca
				dict(2).Add "ekmek", "brotti" 'italyanca
				'1000 satır boyunca 1000 kelimeyi okuyup atadık diyelim
				
				'almancada ekmek ne demek
				Debug.Print dict(1)("ekmek")
				End Sub</pre>
					<p>
					Datamızı bu şekile getirdikten sonra bu yukardaki örnekteki ve benzer 
					örneklerdeki sorunuda bu yöntemle çözebiliriz</p>
					<ul>
						<li>Akdeniz bölgesinin Ticari müdürü kim(diğer 3 segment müdürü de var)</li>
						<li>Ankarada Patates fiyatı(diğer 10 sebze fiyatı da var)</li>
						<li>Almanyada erkeklerin yaş ortalaması kaç(kadınların ortalaması da var)</li>
						<li>v.s v.s</li>
					</ul>
					<p>Mesela bir üstteki sicil-tarih örneğini de bu yöntemle yazabiliriz.</p>

				<pre class="brush:vb">
				Sub arraydict2()
				Dim dict(1) As New Scripting.Dictionary
				Dim alan As Range, a As Range
				
				Set alan = Range("A2:A7")
				For Each a In alan
				  If Not dict(0).Exists(a.Value2) Then 'herhangi birine bakılabilir
				    dict(0).Add a.Value, a.Offset(0, 1).Value 'şb
				    dict(1).Add a.Value, a.Offset(0, 2).Value 'tarih
				  End If
				Next a
				
				For Each s In dict(0)
				    Debug.Print s, dict(0)(s), dict(1)(s)
				Next s
				
				End Sub
				</pre>
 	     </div>
 	     
 	     <h4 class="baslik">Dictionary, Collection ve Collection Dizisi bir arada</h4>
			<div>
			<p>Bu sitenin iletişim sayfasından ara ara benden destek isteyenler oluyor. Yine bir gün
			gelen bir mailde, aşağıdaki gibi bir talep vardı.
			</p>

			<blockquote><em>Parekende elektrik malzemesi satıyoruz. Satışı yapmamız için alış yapmış olmamız lazım. 
			Bir malı x müşterisine y miktar satış yapmış isek o malın yine y miktar girişinin olması gerekiyor. 
			Yani aynı müşteri adına aynı miktarda satış ve alış miktarı olması lazım. İşi sağlama bağlamak için 
			E sütunundaki malzeme adını da ekledim. Fakat alınan malın satışı bir alt sütunda olacak diye bir 
			kaide yok, herhangi bir yerde olabilir. Çünkü bu bir muhasebe çıktısının excel'e çevrilmiş hali. 
			Normalde biz bu işi gelişmiş süzmelerle yapıyor ve tek tek satırları boyuyoruz. Bu da zaman alıyor 
			tabii ki. Haftada birkaç çıktı aldığınızda problem oluyor.
			Aynı olan verileri boyamak istiyoruz. Veri yapısı şöyle: K sütünunda müşteri adı, G sütununda alış 
			miktarı, I sütununda satış miktarı E sütununda stok ismi var. K sütununun satırlarında aynı müşteri 
			adı varsa ve söz konusu satırlardaki alış miktarları(G sütunu satırları) satış miktarlarına(I 
			sütunu satırları) eşitse ve E sütunundaki satır değerleri de eşitse A'dan K dahil o satır boyansın 
			istiyorum. Koşullu biçimlendirme ile yapamadım. Buradaki amaç aynı miktarda olan mallar aynı müşteri 
			için giriş çıkış yapmışsa o satırı boyayıp listeden elimine etmek......
			</em></blockquote>
			<p><img src="/images/dictcolldizi.jpg"></p>
				<p>İlk başta Conditional Formatting ile yapmayı denedim ancak alım ve satımın aynı olma 
				durumu ardışık olmayan satırlarda da olabileceği için bundan vazgeçip aşağıdaki kodu hazırladım.</p>
			<pre class="brush:vb">
Sub mukerrerbul()

Dim dict As Object
Dim koll As New Collection
Dim babaKol() As Collection
Dim a As Range, s As Range
Dim i As Integer
Dim c As Variant

'babaKol collection dizisinin boyutunu buluyoruz
For Each a In Range([e2], [e2].End(xlDown))
    If Not ColdaVarmı(koll, a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value) Then
        koll.Add a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value
    End If
Next a

Set dict = CreateObject("Scripting.Dictionary")
'önce g'nin boyutunu berlieyelim
ReDim babaKol(koll.Count - 1)

'sonra dolu olan alımları dictionary'ye ekliyoruz
For Each a In Range([e2], [e2].End(xlDown))
    If Not dict.Exists(a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value) And Not IsEmpty(a.Offset(0, 2)) Then
        Set babaKol(i) = New Collection
        dict.Add a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value, babaKol(i)
        babaKol(i).Add a.Row
        i = i + 1
    ElseIf dict.Exists(a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value) And Not IsEmpty(a.Offset(0, 2)) Then
        dict(a.Value & a.Offset(0, 2).Value & a.Offset(0, 6).Value).Add a.Row
    End If
Next a

'şimdi satımları kontrol ediyruz. varsa, hem bunu hem de bunun alım karşılığını işaretliyoruz
For Each s In Range([e2], [e2].End(xlDown))
    If dict.Exists(s.Value & s.Offset(0, 4).Value & s.Offset(0, 6).Value) Then
        'önce satımı boyayalım
        s.EntireRow.Font.Color = vbRed
        'şimdi de alımları boyayalım
        For Each c In dict(s.Value & s.Offset(0, 4).Value & s.Offset(0, 6).Value)
            Rows(c).Font.Color = vbRed
        Next c
    End If
Next s

End Sub
Function ColdaVarmı(col As Collection, kontrol As Variant) As Boolean
    On Error Resume Next
    ColdaVarmı = False
    Dim x As Variant
    For Each x In col
        If x = kontrol Then
            ColdaVarmı = True
            Exit Function
        End If
    Next x
End Function			
</pre>
<p>Şimdi elimizde neler var bi bakalım:</p>
				<p><strong>koll</strong>:Toplam kaç değişik benzersiz kaydımız 
				var, bunu tutacağımız koleksiyon.<br><strong>dict</strong>:Firma-Stok-alım 
				adedinden oluşan benzersiz kayıtları tutacak dictionary. Bunun 
				value parametresinde ise ilk başta satır numarasını gösteren bir değişken kullanmıştım ancak sonradan farkettim ki, aynı kayda ait başka mükerrer kayıtlar da olabiliyor, o yüzden tek değer içeren bir değişken yerien bi collection kullanmak gerekiyor, ancak her kayıt için de farklı bir collection kullanmam gerektiği için bunu normal bir collection yerine 
				"collection dizisi" (<strong>babaKol</strong>) şeklinde yarattım.
				<br>Sonra dolu olan alımları dictionary'ye ekledim.<br>Sonrasında, satım miktarlarını bu dictionary içinde var mı diye kontrol ettim, varsa hem kendisinin olduğu satırı hem de bunun dictionaryde karşılık gelen collectiondaki satırları yani alımların olduğu satırları boyattım.</p>
				<p>Dosyanın kendisine de 				<a href="../../../Ornek_dosyalar/Makrolar/dictcolldizi.xlsx">
				buradan</a> ulaşabilirsiniz.</p>
			</div>
	</div>

</asp:Content>
