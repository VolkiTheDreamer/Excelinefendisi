<%@ Page Title='Objeler Dünyası' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik">
	<div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='İleri seviye konular'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label>
</td></tr></table></div>

<h1>Nesneler Dünyası</h1>
<p>VBA konularının başında nesne kavramına biraz girmiş ve Excel'in
<a href="Giris_ExcelNesneModeli.aspx">nesne modelinden</a> bahsetmiştik. O 
bölümü okumayanların önce orayı okumasını tavsiye ederim. </p>
	<p>Bu 
bölümde ise nesnelere biraz daha yakından bakıcaz. Burada Türkçe'nin avantajını 
kullanarak bu kavrama Türkçe adıyla hitap edicem, zira bir de değişken tipi olarak
	<strong>Object </strong>var elimizde. 
	İkisi birbirine karışmasın diye genel nesne kavramını <strong>Nesne</strong> 
	olarak, tip olanı ise <strong>Object</strong> olarak belirticem.</p>
	
<h2 class='baslik'>Giriş</h2>
<div class='konu'>
<h3>Nedir?</h3>
	<p>Nesnelerin ne olduğuna bakmadan önce nesnelerin ne olmadıklarına bakalım. 
	Basit değişkenlerle nesnelerin birbirinden çok temel bir farkı vardır. Basit 
	değişkenlerin tek bir amacı vardır: Bir değer depolamak.</p>
	<pre class="brush:vb">Dim i As Integer
Dim ad As String

i=10
ad="Volkan"</pre>
	<p>Nesneler ise, bir değer depolamaktan daha fazlasını yaparlar. Nesneler, 
	hem çoklu veri tutarlar hem de bir eylem icra ederler.</p>
	<pre class="brush:vb">Dim ws As Worksheet
Set ws = Activesheet
ws.Name="Kredi" 'veri
Debug.Print ws.Index 'veri
ws.Add 'eylem</pre>
	<h3>Nesnelerin bileşenleri</h3>
	<p>Artık bildiğiniz üzere, Excel'de herşey bir nesnedir, hatta nesneler 
	topluluğu olan collectionlar(sonu "s" ile biten) da nesnedir(Ör:Workbook da 
	nesne, Workbook<strong>s</strong> collection'ı da). </p>
	<p>Bu nesnelerin bazıları sadece veri tutarlar, bu verilere 
	özellik(property) denir, bazı nesneler belirli eylemleri(metod) de icra 
	ederler. Bazıları ise ayrıca kendileriyle ilgili bir eylem olduğunda bir 
	olay(event) meydana getirirler.</p>
	<p>Nesne kavramı olmasaydı herşeyi değişkenlerle yönetmemiz gerekirdi ki bu 
	inanılmaz karmaşık bir dünyaya neden olurdu. Düşünsenize, 30 propertysi olan 
	bir obje içi 30 ayrı değişken tanımlamanız gerekirdi, ki bu sadece değişken 
	tanımlamayla ilgili endişemiz, diğer dezavantajlarını sanymıyorum bile.</p>
	<p><strong>NOT</strong>:VBA, tam anlamıyla bir Nesne Yönelimli 
	Programlama(OOP) dili değildir, ama bu kavramı destekler. Excelent 
	eklentisini yazdığım dil olan VB.NET ise tam bir OOP dilidir.</p>
	<h4>Property'ler</h4>
	<p>Bazı propertyler Readonly'dir(salt okunur), yani bunlara değer 
	atayamazsınız, bazıları ise hem okunabilir hem yazılabilirdir.</p>
	<pre class="brush:vb">
MsgBox Activecell.Address 'bu readonlydir
ActiveCell.Value = 1 'bu hem okunur hem yazılabilirdir.
MsgBox ActiveCell.Value</pre>
	<p>Property'lere değer atamak, eğer dönüş değerleri basit data 
	tiplerindeyse, aynı bunlar gibi atanır. i=1 ile ActiveCell.Value = 1 örneğindeki gibi</p>
	<p>Ancak dönüş değeri nesne olan propertylere nesne atamalarındaki
	<span class="keywordler">Set</span> ifadesi ile atama yaparız.</p>
	<pre class="brush:vb"> Set ilkkolon = ActiveCell.CurrentRegion.Columns(1)</pre>
	<h4>Metodlar</h4>
	<p>Metodlar, nesnelerin eylem icra eden üyeleridir. Sub olarak da Function olarak da tanımlanmış olabilirler. Mesela 
	Workbook nesnesinin Add metodu bir Function'dır, zira geriye birWorkbook 
	nesnesi döndürür. Ancak Save metodu bir Sub'dır, zira geriye birşey 
	döndürmez, sadece bir eylem icra eder.</p>
	<h4>Eventler</h4>
	<p>Belli nesnelerin de belirli eylemler olduğunda meydana gelen olayları 
	vardır. Bu konuyu olaylar <a href="Olaylar_Konular.aspx">bölümünde</a> 
	genişçe ele almıştık.</p>
	<h3>Nesne türleri</h3>
	<p>Ben nesneleri 4 gruba ayırıyorum(bu sınıflama tamamen bana aittir, resmi 
	bir gruplama değildir)</p>
	<p><strong>İlk</strong> grupta Excelin nesne modelindeki nesneler var. Range, Cell, Worksheet 
	gibi. Bunlar Excel librarysinde yer alırlar. Tümüne
	<a href="https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/object-model-excel-vba-reference">
	buradan</a> ulaşabilrisiniz</p>
	<p><img src="../../images/vbaobject1.jpg"></p>
	<p><strong>İkinci </strong>grupta, default gelen librarylerdeki nesneler 
	bulunuyor. Yani ilave herhangi bir library'yi reference olarak eklemeden 
	yaratacağımız nesneler. Collection<strong> </strong>gibi; bu nesne 
	VBA librarysinde bulunur.</p>
	<p><img src="../../images/vbaobject2.jpg"></p>
	<p><strong>Üçüncü</strong> grupta, bir library ekleyerek yaratılan nesneler var. Dictionary, 
	FileSystemObject gibi. Bunlar da Scripting Runtime library'sinde bulunurlar.</p>
	<p><img src="../../images/vbaobject3.jpg"></p>
	<p>Excel'in nesne modeli dışında kendi nesnelerimizi de yaratabiliriz, 
	bunlar da <strong>dördüncü</strong> grup oluyor. Tabi 
	bunun için önce bu nesnenin taslağını oluşturan Class yaratmamız gerekiyor. Classlara bu bölümde değinmeyeceğiz, onlarla ilgili bilgiye
	<a href="Ileriseviyekonular_ClassveClassModuller.aspx">buradan</a> 
	ulaşalabilirsiniz.</p>
	<h3>Nesne üyelerine erişmek</h3>
	<h4>Klasik yöntem</h4>
	<p>Bir nesne üyesine erişmek için en bilinen yol, nesne adını yazıp sonra 
	"." koymaktır. Ör: Workook.Name</p>
	<h4>With - End With</h4>
	<p>Bir diğer yöntem de daha önce gördüğümüz <strong>With </strong>kalıbı. Bu 
	kalıbı <a href="Giris_ExcelNesneModeli.aspx#withend">şurada</a> anlatmıştık, 
	sadece kısa bir örnek verelim.</p>
	<pre class="brush:vb">Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
With fd
  .Title = "Klasör seçin"
  If .Show = True Then
    lblKlasör.Caption = .SelectedItems(1)
  End If
End With</pre>
	<h4>Me</h4>
	<p>
	Bir class modülündeyken(workbook, worksheet modülleri dahil, veya kendinize 
	ait bir calass) kendisine Me ifadese ile erişebiliyorsunuz. Burada unutulmaması 
	gereken şudur; Me, her zaman o an içinde bulunulan classa başvurur. Ör: Thisworkbook 
	modülündeyken: Me.Save</p>
	</div>
	<h2 class="baslik">Nesne tanımlama ve yaratma</h2>
	<div class="konu">
	<h3>Genel tanımlama ve yaratım teknikleri</h3>
	<h4 id="newkeyword">New ifadesi</h4>
	<p>Excel nesne modelinde bulunan nesneleri tanımlarken
	<span class="keywordler">New</span> ifadesini kullanmayız. Zira Excel 
	açıldığında bunlar otomatikman yaratılmış olurlar, o yüzden sadece değişken 
	tanımlamak yeterlidir.</p>
	<p>Bunlara atama yapmak için ise <span class="keywordler">Set </span>
	ifadesini kullanıyoruz.</p>
	<pre class="brush:vb">Dim rng As Range
Dim ws As Worksheet

Set rng = Activecell
Set ws = Activesheet</pre>
	<p>Excel nesne modeli dışındaki nesneleri yaratmak için ise <strong>New</strong> ifadesini 
	kullanmak zorundayız. Bu şekilde nesne yaratmanın da iki yolu vardır.</p>
	<h5>Yöntem 1:Tek satırda</h5>
	<p>İlk yöntemde Dim ve New ifadelerini aynı satırda kullanırız, yani <strong>
	tanımlama</strong> ve <strong>yaratım</strong> aynı anda olur. Sonra 
	nesnenin üyelerini hemen kullanmaya başlayabilirz. Aslında yaratım aynı anda 
	olmamaktadır, bu nesne ilk nerede görülürse işte o sırada yaratım 
	olmaktadır.</p>
	<pre class="brush:vb">Dim coll As New Collection
coll.Add "elma"</pre>
	<h5>Yöntem 2:Ayrı satırlarda(Set'li yöntem)</h5>
	<p>Bu yöntemde ise tanımlama ile yaratım&amp;atama ayrı satırlarda gerçekleşir. 
	Tanımlama Dim ile, yaratım&amp;atama Set ve New ile yapılır.</p>
	<pre class="brush:vb">Dim coll As Collection 'Tanımlama
Set coll = New Collection 'Set ile atama, New ile yaratım</pre>
		<h5>İstisna</h5>
		<p>Eğer ki, elde edeceğimiz nesneyi bir fonksiyon veya metod ile elde edeceksek o zaman 
		New ifadesi kullanılmaz.</p>
		<pre class="brush:vb">Dim oApp As New Outlook.Application 'bunda New gerekli
Dim oMail As Outlook.Mailitem 'bunda New kullanılmaz
Set oMail = oApp.CreateItem(0) 'başka bir nesnenin metoduyla elde ettik</pre>
		<p>Başka bir örnek de veritabanı işlemlerinden olsun. Aşağıdaki iki 
		nesne için de New gerekmedi.</p>
		<pre class="brush:vb"> Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(....)
Set rs = db.OpenRecordset(.....)</pre>
	<h5>Hangi yöntem ne zaman kullanılır?</h5>
	<p>İkinci yöntem performans yönetimi açısından tercih edilir.Aslında günümüz bilgisayarları açısından bakıldığında buradaki performans 
	etkisi artık ihmal edilebilir düzeydedir. O yüzden iki yöntem de 
	kullanılabilir. Ancak belli özel durumlarda ikinci yöntemin kullanılması 
	tavsiye edilir.</p>
	<p>Eğer, tanımladığımız değişkeni <strong>belli bir duruma/şarta</strong> göre yaratmamız sözkonusu 
	ise tek satırda değil, iki satırda yani Set'li yöntemle tanımlarız.</p>
	<p>Mesela mail gönderimi yapılacak bir durum düşünelim. Eğer B2 hücresinde 
	bir değer varsa o zaman mail gönderilsin, yoksa gönderilmesin. O yüzden ilk 
	başta tanımlamayı yapalım, ama henüz nesneyi yaratmamıza gerek yok.</p>
	<pre class="brush:vb">Dim oApp As Outlook.Application

If Not IsEmpty(Range("F2")) Then
   Set oApp = New Outlook.Application 
End If</pre>
	<p>Burada, ilgili nesne bir değişkene atanana kadar onu yaratmaz. Diğer 
	yöntemle farkını görmek için aşağıdaki iki kodu çalıştırıp kendiniz görün, 
	gerçi ben yazılacak değerleri yanlarında belirttim, ama kendinizin de 
	görmesinde fayda var.(Outlook libraraysini eklemeyi unutmayın)</p>
	<pre class="brush:vb">
Sub setliyöntem()
Dim oApp As Outlook.Application
Dim coll As Collection

Debug.Print TypeName(coll) 'Nothing
Debug.Print coll.Count 'hata

Debug.Print TypeName(oApp) 'Nothing
Debug.Print oApp 'hata alınır

If Not IsEmpty(Range("F2")) Then
   Set oApp = New Outlook.Application
End If

End Sub
--------------------
Sub teksatıryöntemi()
Dim oApp As New Outlook.Application
Dim coll As New Collection

Debug.Print TypeName(coll) 'collection
Debug.Print coll.Count '0

Debug.Print TypeName(oApp) 'Application
Debug.Print oApp 'Outlook

End Sub
	</pre>
	<p>Tek satırda kullanım yönteminin avantajı ise, ilgili değişkeni Nothing ile 
	yok etseniz bile tekrar kullanabiliyor olmanızdır. Mesela şu kod problemsiz 
	çalışır.</p>
	<pre class="brush:vb">
Sub coll1()

Dim coll As New Collection

coll.Add "Apple"
Set coll = Nothing
coll.Add "Pear" 'yeni bir collection yaratılır

End Sub</pre>
	<p>
	Ancak aynı kodu Setli yönteme çalışıtırırsak hata alırız.</p>
	<pre class="brush:vb">
Sub coll2()

Dim coll As Collection
Set coll = New Collection

coll.Add "Apple"
Set coll = Nothing

coll.Add "Pear" 'hata

End Sub	</pre>
	<h3 id="binding">Early ve Late Binding</h3>
	<p>Bir diğer nesne yaratım yöntemi ise <span class="keywordler">CreateObject
	</span>ile yaratımdır, ki buna <strong>LateBinding</strong> yöntemi ile 
	yaratım 
	denir.</p>
	<p>Bu yönteme daha çok, VBA'nin default libraryleri olan VBA, Excel ve Office 
	libraryleri dışındaki librarylerde bulunan classlardan nesne yaratmak 
	istediğimizde başvururuz. En sık kullanılan classlar şunlardır:</p>
	<ul>
		<li>Scripting.Runtime librarysi içindeki <strong>FileSystemObject</strong></li>
		<li>Yine Scripting.Runtime librarysi içindeki <strong>Dictionary</strong></li>
		<li>Outlook, Word gibi diğer <strong>Ofis uygulamaları</strong></li>
	</ul>
	<p>Syntax'ı şu şekildedir: <span class="keywordler">CreateObject("library.class")</span></p>
	<p>Değişken tanımlamayı <strong>Object</strong> tipli yapıp sonra da Set ile 
	atama ve yaratmayı yaparız.</p>
	<pre class="brush:vb">Dim obj As Object 'tanımlama
Set obj = CreateObject("Outlook.Application") 'yaratma ve atama
Set obj = CreateObject("Scripting.Dictionary")
Set obj = CreateObject("Scripting.FileSystemObject")</pre>
	<p>NOT: Nasıl olsa LateBinding yapıyorum, o yüzden değişken tanımlamaya gerek 
	yok diye düşünmeyin. Zira object tipli değişkenler hafızda 4 byte yer işgal 
	ederken, değişken tanımlamadığınız durumda bunlar otomatikman Variant 
	algılanacakları için hafızada 16 byte işgal ederler.</p>
	<p><strong>Early Binding</strong>'te ise tanımlamak istediğimiz nesnenin 
	bulunduğu library'yi Tools&gt;References menüsünden eklememiz gerekir.</p>
	<p><img src="../../images/vbaobjectbinding1.jpg"></p>
	<p>Library'yi ekledikten sonra artık klasik yoldan değişken 
	tanımlayabiliriz, nesneyi doğru yazma konusunda Intellisense bize yardımcı 
	olur.</p>
	<p><img src="../../images/vbaobjectbinding2.jpg"></p>
	<p>veya bunu 2 satırlık versiyonla da yapabiliriz.</p>
	<pre class="brush:vb">Dim dict As Dictionary
Set dict = New Dictionary</pre>
	<p>Bu örnekte olduğu gibi aynı classtan başka bir library içinde yoksa 
	library adını yazmamıza gerek yoktur, aksi halde karışıklık olur. Böyle bir 
	durumda library adını da belirtmeliyiz.(Veritabanı işlemlerinde bu durumu 
	görüyoruz)</p>
	<p><img src="../../images/vbaobjectbinding3.jpg"></p>
	<h4>Early 
	Binding vs Late Binding</h4>
	<ul>
		<li>Late Bindingin avantajı, yazdığınız kodu bir başkasına 
		gönderdiğinizde(veya network üzerinden çalıştırdıklarında) onda da 
		kesinlikle sorunsuz çalışacağını biliyor olmanızdır. Zira herhangi bir 
		kütüphane eklenmesi durumu olmadığı için versiyon farkları problem 
		yaratmayacaktır. Early bindingte ise siz Office 2016 ile çalışıyorken, 
		diğer kişi Office 2013 ile çalışıyorsa sizin 
		eklediğiniz Outlook 16.0 reference'ini bulamayacağı için hata 
		alacaktır. O yüzden eğer yaptığınız çalışmayı başka birinin bilgisayarında 
		çalıştırma durumu varsa Late binding kullanmanız faydalı olackatır, aksi 
		halde Early binding kullanın.</li>
		<li>Early bindingle çalışmanın avanatjı ise intellisenseten 
		faydalanmaktır. Bu hem daha hızlı hem de hatasız kod yazmanızı 
		sağlayacaktır. LateBinding'te yazdığınız kod Intellisense ile teyit 
		edilmemiş olacağı için tam çalışmadan önce birkaç kez derleme hatası 
		almanız olasıdır.</li>
		<li>Early bindingle çalışmanın bir diğer faydası da kodun çalışma 
		hızıdır. VBA, ilgili nesnenin ne olduğunu direkt bileceği için arka planda bir 
		dönüştürme işlemi yapmasına gerek kalmayacak ve de kod hızlı çalışcaktır. Zaten 
		Early denmesinin sebebi de budur, <strong>öncende/erkenden</strong> hangi classla 
		çalışacağımıza karar vermiş oluyoruz. Latebindingte ise <strong>
		sonradan/gecikmeli</strong> bir tespit işlemi de olacağı için kod daha yavaş 
		çalışacaktır. Ortalama hız farkı 2 kattır.</li>
		<li>Bazı nesneler Late bindingle ile yaratılamaz. Mesela 
		CreateObject("DAO.Database") gibi bir kullanım sözkonusu değildir. Bu 
		yöntemin kullanımı için ilgili nesnelerle ilgili bazı teknik önayarların 
		yapılmış olması gereklidir.
		<a href="https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/activex-component-can-t-create-object-or-return-reference-to-this-object-error-4">
		Burada</a> ve
		<a href="https://support.microsoft.com/en-us/help/828550/you-receive-run-time-error-429-when-you-automate-office-applications">
		burada</a> teknik detaylar yazılı.</li>
	</ul>
	<p>
	<a href="https://support.microsoft.com/en-us/help/245115/using-early-binding-and-late-binding-in-automation">
	Bu sayfaadan</a> çok daha teknik detaylarına ulaşabilirsiniz.</p>
	<p>Benim nihai önerim şudur: Intellisense ve performans sebepleriyle her 
	zaman Early Bindingle başlayın, başkasının bu dosyayı çalıştırması da 
	sözkonusu olacaksa, Early binding olan yerleri latebindige çevirin.</p>
	<p><strong>NOT</strong>: Bir de <strong>GetObject</strong> diye bir 
	fonksiyon vardır. "Mevcutta açık olan bir uygulama varken <strong>
	CreateObject</strong> 
	ile o nesneyi tekrar yaratmanın anlamı yok, ona GetObject ile 
	ulaşabilrsiniz" amacıyla vardır. Günümüz bilgisayarlarındaki bellek 
	kapasitesi düşünüldüğünde çok gereği olmayan bir fonksiyondur. Yine de 
	görürseniz şaşırmayın diye bahsetmek istedim.</p>
</div>
	<h2 class="baslik">Hafızada neler oluyor</h2>
	<div class="konu">
	<h3>Hafıza adresleri</h3>
	<p>Basit veri tipleri sözkonusu olduğunda bunlar için hafızada ayrı yerler 
	açılır. Mesela aşağıdaki örnekte X ve Y için hafızada iki ayrı alan açılır.</p>
	<pre class="brush:vb">Dim X As Integer, Y As Integer
X = 20
Y = 20</pre>
	<p>Bellek gösterimini ise aşağıdaki gibi yapabiliriz. İki değişken de aynı 
	değere sahip olduğu halde iki farklı alan işgal edilir: 2şer byte'tan toplam 
	4 byte.</p>
	<p><img src="../../images/vbaobject4.jpg"></p>
	<p>
	Nesnelerde ise durum biraz farklıdır. Nesneler sözkonusu oluğunda 
	değişkenlerde nesnenin kendisi değil, nesnenin işaret ettiği adres 
	depolanır. Buna programlamada Pointer denir. </p>
	<p>
	Aşağıdaki örnekte bellekte sadece bir alan işgal edilir, o da Object tipli 
	değişkenlerin değeri 4 byte olduğu için toplam 4 byte'tır, 8 byte değil. 
	Zira iki değişken de hafızadaki aynı yere işaret ediyorlar.</p>
	<pre class="brush:vb">Dim wb1 As Workbook, wb2 As Workbook
Set wb1 = Workbooks("deneme.xlsx")
Set wb2 = wb1</pre>
	<p><img src="../../images/vbaobject5.jpg"></p>
	<p>
	Bu da şu demek oluyor; <strong>nesne değişkeni ile nesnenin kendisi farklı 
	şeylerdir</strong>. Aşğağıdaki örnekte <span class="keywordler">VarPtr</span> değişkenlerin bellekteki adresini 
	verirken, <span class="keywordler">ObjPtr</span> nesnelerin kendisinin adresini verir. Adresten kastımız, 
	uzunca bir sayıdır, bu sayının ne olduğu önemli değildir, önemli olan 
	içeriğidir, commentlere bakınız.</p>
	<pre class="brush:vb">Sub hafıza()
Dim wb1 As Workbook
Dim wb2 As Workbook

Set wb1 = ActiveWorkbook
Set wb2 = wb1

Debug.Print wb1.Sheets.Count '1
Debug.Print wb2.Sheets.Count '1

wb1.Sheets.Add

Debug.Print wb1.Sheets.Count '2
Debug.Print wb2.Sheets.Count '2

'wb1 ve wb2 "değişkenlerinin" adresi
Debug.Print "wb1 nesne değişkeninin adresi:" &amp; VarPtr(wb1) 'aşağıdaki ile farklı
Debug.Print "wb2 nesne değişkeninin adresi: " &amp; VarPtr(wb2)

'wb1 ve wb2'nin işaret ettiği yerin adresi
Debug.Print "wb1'in adresi:" &amp; ObjPtr(wb1) 'aşağıdaki ile aynı
Debug.Print "wb2'nin adresi:" &amp; ObjPtr(wb2)
End Sub</pre>
	<h3>Hafızayı temizleme</h3>
	<h4>
	Otomatik ama gecikmeli temizlik</h4>
	<p>
	Bir değişkene bir nesne atadıktan sonra tekrar başka bir nesne atarsak, 
	artık ilk nesne varolmaz, ve bir süre sonra bellekten silinir.</p>
<p>Mesela aşağıdaki örnekte, 1'den 10'a kadar sayıları tutan collection son 
	satırdan itibaren yok olur ve ona erişmenini hiçbir yolu kalmaz.(Bu işlem 
	hemen değil biraz gecikmeli olur)</p>
	<pre class="brush:vb">Dim coll As Collection
Set coll  = New Collection
For i = 1 to 10
  coll.Add i
Next i 
Set coll = New Collection 'ilk nesne yok olur</pre>
	<p>Burda aslında biz bellekte iki tane collection için yer açtık ama son 
	satıra geldiğimizde artık ilkine 
	hiçbir şey atanmış olmadığı için <strong>Garbage Collector</strong> denen sistem 
	bir süre sonra bunu 
	bellekten atar, yani özetle bir nesneye hiçbir değişken işaret 
	etmiyorsa bu nesne bellekten silinir.</p>
	<h4>Manuel ama anında temizlik</h4>
	<p>Bu yöntemi şimdiye kadar birçok örnekte gördük aslında. Bir değişkene 
	<strong>Nothing</strong> değerini atayınca o değişkenle onun başvurduğu nesne arasındaki 
	ilişkiyi kopartırız.</p>
	<p>Aslında çoğu durumda bu işlem gerekli değildir, zira yukarda gördüğümüz 
	gibi bir nesneye başvuran bir değişken kalmadığında bu nesne bellkten 
	gecikmeli de olsa otomatikman silinir.</p>
	<p>Ancak bazı durumlarda, özellikle döngüsel işlemlerde Nothing ataması gerekebilir. Çünkü Garbage Collector'ın ne zaman devreye gireceği belli 
	değildir, ve biz belleği <strong>hemen</strong> <strong>boşaltmak</strong> istiyorsak işte o zaman Nothing 
	ataması yaparız.</p>
	<p>Mesela toplu mail gönderiminde kullandığımız aşağıdaki koda bakalım. 
	oMail değişkenine Nothing atamak faydalıdır, zira bunu yapmazsak ilk mail 
	nesnesi hala bir süre daha bellekte kalmaya devam edecek, taki GC gelip onu 
	yokedene kadar. Eğer tek seferde yüzlerce mail atacaksanız bu işlemi yapmanızı 
	şiddetle öneririm, aksi halde bellekte yüzlerce mail nesnesi birikebilir ve 
	işlem bellek yetersizliğinden yarıda kesilebilir.</p>
	<p>Bununla beraber son satırdaki oApp değişkenine Nothing ataması çok da 
	kritik değildir. Ben yine de alışkanlıkla bunu yapmayı tercih ediyorum. yorum. </p>
	<pre class="brush:vb">
Sub çoklumail_Button1_Click()
Dim oApp As Outlook.Application
Dim oMail As Outlook.MailItem
Dim alıcılar As Range, a As Range

Set oApp = New Outlook.Application
Set alıcılar = Range(Range("A2"), Range("A2").End(xlDown))

For Each a In alıcılar
    Set oMail = oApp.CreateItem(olMailItem)
    With oMail
        .Subject = "Doğum günü"
        .To = a.Value
        .Body = a.Offset(0, 3).Value &amp; "Doğum gününüz kutlar, ailenizle birlikte mutlu yıllar dilerim"
        .Body = .Body &amp; vbCrLf &amp; "Gönderenin adı soyadı"
        .Send
    End With
    Set oMail = Nothing 'zorunlu değil ama faydalı
Next a

Set oApp = Nothing 'zorunla da değil kritik faydası da yok
End Sub</pre>
</div>

</asp:Content>
