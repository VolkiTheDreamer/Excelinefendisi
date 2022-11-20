<%@ Page Title='DizilerveDizimsiYapilar Collectionlar' Language='C#' MasterPageFile='~/MasterPage.master' 
AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
	<div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Diziler ve Dizimsi Yapılar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Collectionlar</h1>

	<p>Collection'lar, Diziler gibi benzer özellikli öğeleri bir arada tutan 
	yapılardır. Dizilerden farklı olarak, tek bir veri tipleri yoktur, yani 
	içlerinde aynı anda hem sayı hem metin tutabilirler. Ayrıca dizilerdeki 
	boyut yerine sayısal bir <strong>Index</strong> ile metinsel özelliği olan 
	opsiyonel bir <strong>Key</strong> 
	unsurlarına sahiptirler. Dizilerle daha genel bir kıyasalamayı aşağıda 
	bulabilirsiniz.</p>
	
<h2 class="baslik">Genel Bakış</h2>	
<div class="konu">
	<h3>İki tür Collection</h3>
	<p>Daha derinlere dalmadan, birkaç tanıdık Collection'dan bahsetmekte 
	fayda var. Workbook<span style="color: red"><strong>s</strong></span>, Worksheet<span><span style="color: red"><strong>s</strong></span></span>, Sheet<span><span style="color: red"><strong>s</strong></span></span> 
	v.s. Bunlar Excel'in <strong>yerel</strong>(built-in) collectionlarıdır. Bunlar VB6 ve VB.Net'teki Collection class'ını baz alırlar ve çok daha fazla 
	üyeye sahiptirler, o yüzden daha kullanışlıdırlar. Bizim burada işleyeceğimiz ise 
	Collection classının VBA uyarlaması 
	olup, VB6/VB.Net versiyonuna göre biraz daha çelimsizdir.</p>
	<p>Yerel collection'lar VB'nin Collection sınıfından geldikleri için 
	yereldirler, o yüzden ayrıca tanımlanmazlar, Excel 
	açıldığı anda otomatikman kullanılır haldedirler(Hepsi olmasa da büyük 
	çoğunluğu).</p>
	<p>O yüzden şöyle bir kod yazmak çok anlamsızdır.</p>
	<pre class="brush:vb">Dim wsc As New Worksheets
For Each ws in wsc
      ......
Next ws
</pre>
	<p>Onun yerine Worksheets'i doğrudan kullanırız.</p>
	<pre class="brush:vb">
For Each ws in Worksheets
      ......
Next ws</pre>
	<h3>Tanımlama</h3>
	<p>Şimdi, konumuz olan Collectionlara gelecek olursak bunların tanımlaması 
	ve kullanımı iki şekilde olabilir.</p>
	<pre class="brush:vb">Dim col As Collection 'Tanımlandı
Set col = New Collection 'Hafızada yer ayrıldı</pre>
	<p>Başka bir tanımlama şekli de tanımlama ve yaratma işleminin 
	tek satırda olduğu şekildir.</p>
<span>
	<pre class="brush:vb">Dim col As New Collection </pre>
	<p>
<span>
	İlk yöntem her zaman en güvenli yöntem olmakla birlikte e</span>ski bir bilgisayarınız yoksa performans sorunu yaşamayacağınızı 
	düşünerek ikinci yöntemi de güvenle kullanabilirsiniz. 
	</span>Bir diğer fark da şudur: Eğer yarattığınız collection üzerinde <strong>
	"Nothing mi?"</strong> kontrolü yapmanız gerekiyorsa ilk yöntemde(o an 
	sadece ilk satırı yazılmışsa) yaratılan Collection bir değer döndürmezken, 
	yani Nothing iken, ikinci yöntemle yaratılan False döndürür. Çünkü ikinci yöntemde New diyerek atamayı da yapmış oluyoruz, ilk yöntemde ise sadece hafızada yer ayırmış olduk, henüz Collection nesnesini yaratmadık. Aşağıda bunu açıklayan bir örnek görebilirsiniz. 
	Konuyla ilgili daha detaylı bilgiye
	<a href="Ileriseviyekonular_ObjelerDunyasi.aspx">buradan</a> 
	ulaşabilirsiniz.</p>

<pre class="brush:vb">
Sub newnothingcol()
    Dim col1 As Collection    
    Dim col2 As New Collection
   
    MsgBox "col1'in türü " &amp; TypeName(col1)
    MsgBox "col2'nin türü: " &amp; TypeName(col2)
    
    If col1 Is Nothing Then 'buraya girer
        MsgBox "col1 hiçbirşeydir"
    End If
        
    If col2 Is Nothing Then 'buraya girmez
        MsgBox "col2 hiçbirşeydir"
    End If

    Set col1 = New Collection
    MsgBox "col1'in türü: " &amp; TypeName(col1)

End Sub
</pre>


	<h3>Sınıfın Üyeleri(Property ve Metodlar)</h3>
	<p>Bu sınıfın yanlızca 4 üyesi vardır. <strong>Add, Remove, Item, Count.</strong></p>
	<p>Üye sayısı az olduğu için bunlara tek tek bakacağız.</p>
	<h4>Add metodu ile yeni eleman eklemek</h4>
	<p>Üyelerden en sık kullanılanıdır ve ana üyedir diyebiliriz.</p>
	<p><span>Add metodu default olarak, yeni elemanları en sona ekler. Bununla 
	beraber Before/After parametresi ile belirli bir indeksten önce veya sonra da 
	eklenebilir. </span></p>
	<pre class="brush:vb">
Col.Add 10 'İçerdeki elemanlar:10
Col.Add 20 'Şimdi 10, 20
Col.Add 30, Before:= 1 'Şimdi 30, 10, 20</pre>
<span>
	<p>
	Excel'in built-in Collectionlarında da yine bu Add metodunu kullandığımızı 
	biliyorsunuz, aşağıdaki gibi:</p>
	<pre class="brush:vb">Workbooks.Add
Worksheets.Add 'bunda After ve Before da kullanılabilir</pre>
	</span>
	<p>
	Ekleme işlemi yapılırken opsiyonel bir parametre olan <strong>Key</strong> de belirtilebilir. 
	Key'in amacı, ilgili elemana indeks yerine ismi ile de ulaşmayı sağlamaktır.</p>
<span>
	<pre class="brush:vb">
Col.Add 10, "on"</pre>
	<p>
	<span>Bu parametre <strong>string bir ifadedir ve benzersiz olmalıdır</strong>. Yani Item 
	olarak 10'u birden fazla ekleyebilrisiniz(niye yapasınız ki :)), ama bunun key'i 
	olan "on" sadece bir kez 
	geçmeli.</span></p>
	<pre class="brush:vb"> Sub keyvaluepair()
Dim notlar As New Collection

'opsiyonel bir parametre olan Key de Add içine yazılabilir
notlar.Add 67, "Volkan" '67 value, volkan key
notlar.Add 67, "Meltem" 'item 2.kez geçiyor, hata vermez
'notlar.Add 12, "Meltem" 'key 2.kez geçiyor, hata verir

End Sub</pre>
	<p>Pratikte key kullanımı çok olmayacaktır diye tahmin ediyorum, daha çok 
	Excelin kendi Collection'larında kullanılır. Workbook veya Worksheetlerde 
	bunu yapıyoruz; Worksheets(1) yerine Worksheets("krediler") yazmak gibi. Bu 
	ikili bilgi ekleme şeklinde programlama literatüründe <strong>Key-Value</strong> 
	pair, yani Anahtar-Değer ikilisi denir. Gerçi Collectionlarda bunların 
	eklenme şekli biraz terstir; Key-Value değil, Value-Key'dir.</p>
	<p>Bununla birlikte, bölge kodu-bölge adı gibi bir ikili bilgiyi(Key-Value) aynı anda 
	eklemek istiyorsanız Collection yerine 
	<a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">Dictionary</a> kullanmanızı öneririm. 
	Açıkçası bana yerel collectionlarda Key kullanımı anlamlı geliyor ama Collection sınıfında 
	bu özelliği neden koymuşlar anlayamadım, Dictionary bu ihtiyacı oldukça 
	karşılarken!</p>
	</span>
	<h4>
	Remove 
	ile elemanları silmek</h4>
	<p>
	Eklenen elemanları çıkarmak için <strong>Remove</strong> 
	metodunu kullanırız. Elemanın indeks numarasını(sade veya parantez içinde) veya Key'i 
	belirtmek yeterlidir. </p>
	<pre class="brush:vb">
Sub silme()
Dim col As New Collection

col.Add 10, "on"
col.Add 20, "yirmi"
col.Add 30, "otuz"
col.Add 40, "kırk"

col.Remove 4 'indeksle
col.Remove (3) 'indeksle
col.Remove "yirmi" 'key ile
'col.Remove (10) 'hata. value ile silme yoktur

Debug.Print col.Count

End Sub	&nbsp;
</pre>

	<p>Collectiondaki tüm elemanları silmek yani collection'ı boşaltmak için For 
	Next içinde Remove kullanırız.</p>
	<pre class="brush:vb">For i = coll.Count To 1 Step -1
    coll.Remove i
Next i</pre>
	<p>
	
    Collection'ı boşaltmaya benzese de tamamen farklı şey olan iki yöntem daha 
	vardır.</p>
	<pre class="brush:vb">
'Collection'ı yeniden tanımlamak.
Set col = New Collection 

'veya onu yoketmek.
Set col = Nothing	
</pre>

	<h4>Count</h4>
	<p>Belirli bir anda Collection'ımız içinde kaç eleman var bunu kontrol etmek 
	istersek <strong>Count </strong>özelliğini kullanırız.</p>
	<pre class="brush:vb">Debug.Print col.Count</pre>
<span>
	<pre class="brush:vb">Sub collornek()
Dim col As New Collection
col.Add 10, "on"
col.Add 20, "yirmi"

Debug.Print col.Item(2)
Debug.Print col(2)
Debug.Print col("yirmi")
Debug.Print col.Count

col.Remove "yirmi"
Debug.Print col.Count

Set col = Nothing
Debug.Print col.Count

End Sub</pre>
	</span>
	<h3>	Elemanlara erişim</h3>
	<p>
	Indeks(item no) ile veya doğrudan item numarası ile elemanlara erişebiliriz. 
	Item özelliği 
	Collectionlar için default özellik olduğu için bunu belirtmesek de olur. 
	Bu arada dizilerden farklı 
	olarak Collectionlarda indeks 0'dan değil 1'den 
	başlar. Daha önce belirttiğimiz gibi elemanlara
	ismiyle ulaşmak da mümkündür, tabiki eklerken <strong>Key</strong> ile 
	eklemişsek.</p>
	<pre class="brush:vb">Sub erişim()
Dim col As New Collection
col.Add 10, "on"
col.Add 20, "yirmi"

Debug.Print col.Item(2) '20
Debug.Print col(2) '20 
Debug.Print col("on") '10
End Sub</pre>
	<h4>İlk ve Son üyeye erişmek için</h4>
	<p>Collectionlar sıralı yapılar oldukları için ilk ve son eleman özel öneme 
	sahiptirler, en azından bazı durumlarda böyledir. İlk elemana ulaşmak basit, 
	indeksi 1 veririz. Son elemana ulaşmak için eleman sayısını saydırır, ve bu 
	sayıyı indeks olarak veririz.</p>
	<pre class="brush:vb">col(1) 'ilk eleman
col(col.count) 'son eleman</pre>
	<p>Tüm elemanlar üzerinden geçmek için For döngüsü kullanırız. Klasik veya 
	For Each.</p>
	<pre class="brush:vb">'Klasik For
For i = 1 To coll.Count
   Debug.Print coll(i)
Next i

'For Each
For Each m In meyveler
    Debug.Print m
Next m</pre>
	<p>Farkettiyseniz elemanları sadece okuduk, yani onları çağırdık ve 
	değerlerini ekrana yazdırdık. Onlara bi değer atamadık, değerlerini 
	değiştirmekdik. <span>Bunun bir sebebi var: <strong>Collectionlar 
	read-onlydir</strong>. Yani eklediğimiz elemanların değerini değiştiremeyiz, 
	onları sadece okuruz. Bu önemli bir dezavantaj gibi görünmekle birlikte 
	kullanım amacına uygun kullandığımıda çok da aradığımız bir özellik değildir. 
	Ama siz ille de değiştirilebilir bir dizi yapısı kurmak istiyorsanız, ya 
	normal dizi ya da Dictionary kullanmalısınız.</span></p>
	<p><span class="dikkat">NOT</span>:Collectionlarla hiçbir zaman Key'i elde edemeyiz. 
	Bu ifade, Collection'ların Readonly olmasından farklı birşeydir. 
	Key'i elemana ulaşmak için kullanırız ama 
	ona değer atayamayız,&nbsp; değerini de elde edemyiz. Yani 1.indekste yer alan Item=10, 
	Key="On" olan bir elemana Col(1) diyerek 10 değerine ulaşabilirken "On" 
	değerine hiçbir şekilde ulaşamayız. "On"u sadece bu elemana ulaşırken 
	indeksin alternatifi olarak kullanırız.</p>
	<h3>"Elemanlar arasında X var mı?" kontrolü</h3>
<p>Collection sınıfının çok az üyeye sahip olduğunu gördük. Ve bunların 
arasında da başka dillerde olan <strong>Contains</strong>, <strong>Exists</strong> 
gibi "Var mı?" kontrolü yapacak bir metodu yok. Bunun yerine bu işi görecek bir 
fonksiyon yazılmaktadır.
</p>

<p>Bu örnekte Item'ın Collection'da olup olmadığna bakıyoruz.</p>
<pre class="brush:vb">
Function ColdaVarmı(col As Collection, kontrol As Variant) As Boolean
    On Error Resume Next
    ColdaVarmı= False
    Dim x As Variant
    For Each x In col
        If x = kontrol Then
            ColdaVarmı= True
            Exit Function
        End If
    Next
End Function

'kullanımı da şöyledir
Sub VarmıCol()
Dim sayılar As New Collection
sayılar.Add 10
sayılar.Add 20
sayılar.Add 30
 
Debug.Print ColdaVarmı(sayılar, 20) 'True döner
Debug.Print ColdaVarmı(sayılar, 40) 'False döner
End Sub
</pre>

<p>Item yerine Value değerinin olup olmadığına bakmak içinse şöyle bir fonksiyon yazabiliriz. Bu fonksiyon	<a href="http://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba">
	buradan</a> alıntıdır. Bu daha pratiktir ve tüm collection'ı dolaşmak zorunda kalmadığı için daha hızlıdır. Ama bunu kullanabilmek için Value özelliğini kullanmanız gerekiyor. Yaptığı iş şudur. Contains'e önce doğrudan True atıyor, sonra obj değişkenine col(key) ile atama yapmaya 
	çalışıyor, sanki varmış gibi. Eğer varsa atar ve fonksiyondan çıkar, 
	Contains'e de en son True atandığı için True döner; obj=col(key) satırı hata 
	alırsa, yani bu key bu Collection içinde yoksa, hata ele alma bloğunda 
	Contains'e False atanır.</p>
	<pre class="brush:vb">
Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:

    Contains = False
End Function

'kullanım şekli yukardakine benzer, sadece Value parametresi string olmak zorunda olduğu için stringe çevrilerek yazılır.

Sub VarmıCol()
'kullanımı da şöyledir
Dim sayılar As New Collection
sayılar.Add 10, CStr(10)
sayılar.Add 20, CStr(20)
sayılar.Add 30, CStr(30)

Debug.Print ColdaVarmı(sayılar, CStr(20)) 'True döner
Debug.Print ColdaVarmı(sayılar, CStr(40)) 'False döner

End Sub</pre>
	<h3>
	Collectionları sıralama</h3>
	<p>
	Collection sıralama için malesef hazırda gelen bir sıralama metodu 
	bulunmamaktadır. Bu nedenle aşağıdaki gibi bir sub prosedür yazarız. Bunu 
	istersek Function olarak yazıp dönen değeri de atayabilirdik.</p>
	<pre class="brush:vb">
Sub ColSırala(mycol As Collection)

Dim i As Long, j As Long
Dim geçici As Variant

For i = 1 To mycol.Count - 1
    For j = i + 1 To mycol.Count
        If LCase(mycol(i)) > LCase(mycol(j)) Then 'büyük küçük harf ayrmı olmasın diye
            'küçük elemanı depolayalım
            geçici = mycol(j)
            'küçük elemanı çıkaralım
            mycol.Remove j
           'küçük elemanı büyük elemandan önceye yerleştirelim
             mycol.Add geçici, geçici, i
        End If
    Next j
Next i

End Sub</pre>
	<p>
	Aşağıdaki örnekle de bir collectionımızı sıralayalım.</p>
	<pre class="brush:vb">
Sub ColSıralamaÖrneği()
Dim col As New Collection

col.Add "volkan"
col.Add "meltem"
col.Add "doruk"
col.Add "doga"

Call ColSırala(col)
'Sıralama sonrasında yazıralım
For i = 1 To col.Count
    Debug.Print col(i)
Next i

End Sub</pre>
	<p>
	Tabi sayılardan oluşan bir collection'ı sıralamak istediğinizde yukarıdaki 
	fonksiyonu biraz değiştirmek gerekir. Aşağıdaki değişen bloğu 
	görebilirsiniz. İlk olarak LCase dersek sayıları metin gibi algılar ve 150yi 
	20'nin öncesine koyabilir, zira ilki "1" ile ikincisi "2" ile başlıyordur. 
	Bir diğer değişiklik de Key kısmında "k" şeklinde bir prefix(önek) ekleriz, 
	zira bu parametre String olmalıdır.</p>
	<pre class="brush:vb">
If mycol(i) &gt; mycol(j) Then 'Lcase kullanılmaz
    geçici = mycol(j)
    mycol.Remove j
    mycol.Add geçici, "k" &amp; geçici, i 'k prefixi eklenir
End If
</pre>

<p>Bu arada siz de Diziler konusunda gördüğümüz diğer sıralama algoritmalarını Collectionlara uyarlamaya çalışabilirsiniz.</p>
	<h3>Dizi ve Collection karşılaştırması</h3>

	<ul>
		<li>Dizilerde Index 0'dan da 1'den de başlayabilir, Collectionlarda her zaman 
	1'den başlar.</li>
	
		<li>Dizilerin aksine Collectionlarda yeni eleman ekleme ve çıkarma<strong>
		</strong>oldukça basittir, herhangi bir boyut, index v.s belirtmeden eleman 
	eklenebilmektedir.</li>

		<li>Diziler genellikle belirli bir boyuta sahiptirler, bu ya statik 
	dizlerdeki gibi baştan belirlenir veya dinamik dizilerdeki gibi sonradan 
	belirlenir, her halükarda <span style="text-decoration: underline">genelde</span> dizinin
		boyutu sabittir. Tabiki dinamik dizilerde 
		<span class="keywordler">ReDim Preserve</span> deyimi kullanılarak 
		dizinin içeriği korunacak şekilde boyut artırılabilir ancak bu birkaç kez kullanılırsa verimsiz bir 
	yöntem olur. İşte böyle durumlarda collection kullanmak daha mantıklıdır, zira 
	collectionlarda boyut diye birşey yoktur. </li>
	
		<li>Collectionlar, eleman sayısı artınca hantallaşır ve performans kayıpları yaşanabilir. 
	(Binlerce elemandan bahsediyorum tabi)</li>
		<li>
		Collectionlar read-only'dir. Yani eğer elemanları değiştirmeyi 
		düşünüyorsanız veya en azından böyle bir ihtimal varsa Collection yerine dizi kullanmalısınız.
</li>
	</ul>
	<p><strong>Genelleme yapacak 
	olursak, boyut baştan bilinmiyor ve çok sık değişecekse Collection kullanmak 
	gerekirken, boyut baştan (veya ilerde bi noktada) biliniyor ve sabit ise 
	dizi kullanmak daha doğru bir yol olacaktır .</strong></p>
	
	
	<h3>Dictionary ile karşılaştırma</h3>
	<p>Dizilerle kıyastan ziyade, Collectionları aslında Dictionary'ler ile 
	kıyaslamak daha mantıklıdır. Dictionary'lerin çok esnek yapıları olmasına 
	rağmen neden Collection kullanalım ki? diye soranlar için bir karşılaştırma 
	listesi Dictionary bölümünde yer almaktadır. <a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">Buradan</a> 
	bakabilirsiniz.</p>
	<h3>Bir örnek</h3>
	<p>İşyerinde, sık güncelleme yaptığım bazı Access dosyalarım bulunmakta. 
	Bilen bilir, Access dosyalarında tablolara yeni alan ekledikçe boyut büyür, 
	üstelik eskisini silip yenisini yükleseniz bile. İçerdeki data büyümediği 
	halde boyut büyümesini engellemek için Compact işleminin yapılması gerekir. 
	Bunlar da büyük dosyalar için vakit alan işlemlerdir. Benim de Application 
	konusunda ele aldığım gibi Schedule edilmiş bir sürü rutinim var. Onlardan 
	biri de bu accessleri compact edip boyutlarını küçülten bir makro. Schedule 
	ediyorum, çünkü neden bu 
	iş gece olabilecekken ben PC başındayken olsun ki!</p>
	<pre class="brush:vb">
Sub accessleri_compact()
'tüm dosyaların kapalı olması lazım, çünkü exlusively açılıyor: Bu işi acceslerin içine timer koyarak hallettim
On Error Resume Next 
Dim app As Object
Dim DBler As New Collection 'dosyalara tek tek aynı işlemi yapmamak için bir collectiona atayacağız
Dim çalıştımı As Boolean
Dim IMsgFilter As Long 'ole mesajını yoketmek için. Buna takılmayın şimdi.

CoRegisterMessageFilter 0&, IMsgFilter 'ole mesajını yokediyoruz

adet = 0

Set app = CreateObject("Access.Application")

'dosya isimlerini değiştirirerek veriyorum
DBler.Add ("C:\..........accdb")
DBler.Add ("C:\..........accdb")
DBler.Add ("C:\..........accdb")
DBler.Add ("C:\..........accdb")

'şimdi de collection içinde geziniyor ve her eleman için aynı işlemi yapıyoruz
For Each d In DBler
    cmp = Left(d, Len(d) - 6) & "_cmp.accdb"
    okmi = app.CompactRepair(d, cmp, False) 'boolean döndürdüğü için böyle yapıyoruz
    
    If okmi = True Then 'başarılı şekilde compact olduysa
        If FileLen(d) = FileLen(cmp) Then 'eğer compact sonucunda dosya daha da küçülmediyse, boşuna işlem yapmaya gerek yok,
                                          'sadece yeni üretilen dosyayı silelim, böylece dosyamızın son erişim tarihini de değiştirmemiş oluruz
            Kill cmp
        Else
            Kill d 'orjinal dosyayı siliyoruz
            Name cmp As d 'Kompakt edilen dosyayı orjinal ismi ile rename ediyoruz
            adet = adet + 1
        End If
    End If
    
Next d

Set app = Nothing

If adet > 0 Then
    rapor = "access compact " & adet & " out of " & DBler.Count
    alici = "12345" 'benim sicilim ve aynı zamanda mail adresim
    Call Mailat2(rapor, alici) ' bu kendime bilgi maili gönderen bi başka prosedür
End If

CoRegisterMessageFilter IMsgFilter, IMsgFilter 'ole mesjaını restore ediyoruz
Exit Sub

hata:
    CoRegisterMessageFilter IMsgFilter, IMsgFilter 'ole mesjaını restore ediyoruz
    rapor = "accesler compact"
    alici = "12345"
    Application.Run "Personal.xlsb!mailnogo", rapor, alici
End Sub	</pre>
	</div>

<h2 class="baslik">İleri konular</h2>	
<div class="konu">

	<h3>Collectionları prosedürlere parametre olarak göndermek</h3>
	<p>Tıpkı dizilerde olduğu gibi Collection'ları da parametre olarak başka bir 
	posedüre gönderebiliriz. Dizilerdeki örneği buraya da uyarlayabiliriz.</p>
<span><pre class="brush:vb">
Sub colgonder()
Dim col As New Collection

col.Add "volkan"
col.Add "meltem"

Call Mesajver(col)
End Sub


Sub Mesajver(coll As Collection)
    MsgBox "Bu collectionda " &amp; coll.Count &amp; " adet elaman var"
End Sub	</pre>
		</span>
	<p>Bu yöntemin güzel bir örneği de hemen bi alttaki kısımda yer 
	almaktadır.</p>
	<h3> 
	Collectionları 
	Dizilere dönüştürme</h3>
	<p> 
	
	Bazı durumlarda elde ettiğimz Collection'ı, Dizi özelliklerinden faydalanmak 
	veya parametre olarak Dizi alan bir fonksiyonda kullanmak için diziye 
	çevirmemiz gerekir. Aşağıdaki fonksiyon bu işi yapmaktadır.</p>
<pre class="brush:vb">
Function CollectionToArray(col As Collection) As Variant()
    Dim arr() As Variant, i As Long, t As Variant

    'collectiondaki elemans sayısından 1 çıakrıp dizi boyutunu belirliyoruz
    ReDim arr(col.Count - 1) As Variant 
    'tüm elemanlar tek te diziye atanır
    For Each t In col
        arr(i) = t
        i = i + 1
    Next t
    CollectionToArray = arr
End Function

'Aşağıda da kullanım örneği bulunuyor
Sub TestCollectionToArray()
    Dim sayıcol As Collection, sayıdizi() as Variant
    Set sayıcol = New Collection
    sayıcol.Add 10
    sayıcol.Add 20
    sayıcol.Add 30
    sayıdizi= CollectionToArray(sayıcol)
    Debug.Print UBound(sayıdizi) '2(indeks 0dan başladığı için)
End Sub 
</pre>
	<h3> 
	Collection Collectionı(İçiçe Collection)</h3>
	<p> 
	İçiçe dizilerde böyle bir kullanım şeklinin amacını ve yöntemini görmüştük. 
	Henüz bakmadıysanız oraya bakmanızı tavsiye ederim. Benim böyle bir kullanım 
	şekline şimdiye kadar ihtiiyacım olmadı, ama yine de sizilerin olabilir diye 
	aşağıya
	<a href="http://bytecomb.com/collections-of-collections-in-vba/#jagged-collections">
	bytecomb</a> sitesinden aldığım bir örneği koyuyorum. Sitede belirtildiği 
	gibi, eğer içteki kümeye sık sık eleman ekleyecekseniz içiçe dizi yerine 
	içiçe collection kullanmak daha mantıklıdır. Onun dışında pek bi kullanım 
	farkı yok gibi görünüyor.</p>
	<pre class="brush:vb">Dim cAnimals As New Collection 
 
' Let's add stats on the Cheetah
Dim cCheetah As New Collection
 
' Easy to add inner collections to the outer collection.  Also, cCheetah refers
' to the same collection object as cAnimals(1).  
cAnimals.Add cCheetah          
 
' Easy to add items to inner collection.
' Working directly with the cCheetah collection:
For Each vMeasurment In GetMeasurements("Cheetah")
    cCheetah.Add vMeasurement
Next
 
' Working on the same collection by indexing into the outer object
For i = 1 To cAnimals.Count
    For j = 1 To cAnimals(i).Count
        cAnimals(i)(j) = cAnimals(i)(j) * dblNormalizingFactor
    Next
Next</pre>
	<h3 id="colsonuc">Fonksiyonlarda dönen değer olarak</h3>
	<p> 
	Bazen bir Colleciton'ı birkaç farklı prosedür içinde kullanmak gerekebilir. o 
	yüzden bunu bir fonksiyon olarak yazıp, dönen değer olarak da ilgili 
	Collection'ın gelmesini sağlayabiliriz.</p>

<p>Aşağıda benim işte kullandığım bir fonksiyon ve bunun kullanım örneği bulunmakta. Her ay Database şifrelerimizin süresi dolmaktadır 
ve bu yüzden şifrelerin aylık olarak yenilenmesi gerekmektedir. Benim SQL'lerin çoğu da Excel'e gömülü ve schedule edilmiş durumdalar. Şimdi yaklaşık 50 civarı rapor olduğunu ve bazısında birden çok connection olduğunu düşünüecek olursak toplam connection sayısı 100 civarında 
olduğunu söyleyebilirim. Bunların bazısı 5-10 sn'de çalışırken bazısı 10-15 dk sürebiliyor. Bunların her birinde manuel değişiklik yapmak çooooooook uzun zaman alacaktır, 
sadece connection sayısı çok olduğu için değil aynı zamanda her manuel 
değişiklik sonrasında sorguların çalışacak olması ve benim bunların bitmesini 
beklemem gerektiği için. </p>
	<p>İşte bu kod ile, öncelikle dosyaları collection'a atıyorum. Bu collection'ı hem şifre değiştirmede hem de sonrasında değiştirdiğim şifrelerin hepsinin değişip değişmediğini görmek için iki kez kullanıyorum. Bu yüzden bir fonksiyona atamak daha mantıklı oldu. Dosyaları açmadan önce otomatik refresh olmasınlar diye EnableEvents=False yapıyorum, varsa protectionları geçici kaldırıyorum, bu detayları aşağıdaki koda koymadım, sadece kafanızda soru işareti olabilir diye 
	belirtmek istedim. Kodun diğer detaylarını <a href="/Konular/VBAMakro/DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">Connection'ları</a> işlediğim sayfada veriyor olacağım.</p>


	<pre class="brush:vb">'fonksiyonu çağırma ve kullanma
Sub collfunc()
	Dim files As Collection
	
	Set files = dosyacoll()
	For Each file In files
		Debug.Print file
	Next file
End Sub

'Fonksiyonun kendisi
Function dosyacoll() As Collection
Const gunlukyol As String = "…………."
Const haftalıkyol As String = "…………."
Const bbsp As String = "…………."
Const hgtakip As String = "…………."
Dim files As Collection
Dim DBtür As Byte

DBtür = Application.InputBox("DB türünü girin. Oracle 1.grup için 1, 2.grup için 2, DB2 için 3", Type:=1)


'collecitionı oluşturalım
Set files = New Collection
With files
    If DBtür = 1 Then
            .Add (hgtakip + "Miy access data.xlsb")
            'Diğer 20 küsur rapor
            '.....    
    ElseIf DBtür = 2 Then
            '10 küsur rapor
            '.....    
    ElseIf DBtür = 3 Then
            '10 küsur rapor
            '.....    
    Else
        MsgBox "yanlış DB türü girdiniz"
        Application.EnableEvents = True
        Exit Function
    End If
End With

Set dosyacoll = files

End Function
</pre>
</div>
</asp:Content>
