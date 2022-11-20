<%@ Page Title='VeriTabanı İşlemleri - DAO ve ADO' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Diğer Uygulamalarla iletişim'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Veritabanı İşlemleri - DAO ve ADO</h1>
	<p>Bu bölümde gösterilen örneklere ait dosyaları
	<a href="../../Ornek_dosyalar/Makrolar/vbadata.rar">buradan</a> 
	indirebilirsiniz.</p>

<h2 class="baslik">DAO DETAY</h2>
<div class="konu">
	<h3>Giriş</h3>
	<p>Önceki sayfayı okuduysanız DAO’nun MS Accese’e özgü bir API olduğunu anlamış 
	olmalısınız. Daha önce belirttiğimiz gibi MS Access dışında bir 
	veritabanına bağlanmak için yeni kodların ADO ile yazılması tavsiye ediliyor 
	ancak mevcut kodlarınızı ADO’ya çevirmeniz pek bir anlamı yok. Accesse 
	gelince, kullanımı daha pratik olduğu için ben hala DAO'yu kullanıyorum. 
	Seçim size kalmış, siz hangisinde kendinizi rahat hissederseniz onu 
	kullanabilrsiniz.</p>
	<h4>Motor/Engine</h4>
	<p>DAO, dataya erişim için JET DB motorunu kullanırdı, ancak 2007 versiyonu 
	ile birlikte ACE geldi. ACE ile eski mdb dosyalarına da yeni accdb 
	dosyalarına da erişebileceğiz. Ancak araştırmalarınız sırasında JET’e de 
	rastlarsanız şaşırmayın.</p>
	<p>Şimdi gelin bunu kullanabilmek için neler yapılmalı bir bakalım.</p>
	<h4>Referans ekleme</h4>
	<p>Öncelikle DAO kütüphanesini VBA’e eklemek gerekiyor. Personal.xlsb 
	üzerinde çalıştığımız düşünerek bu dosyadayken aşağıdaki işlemleri yapalım.</p>
	<p>Birçok yerde DAO 3.6 ekleyin denir, ama bunların çoğu eski siteler. Artık 
	Access 2003 kullanan kalmadığını düşünerek bunun yerine <strong>Microsoft 
	Office x.x Access Database Engine Object Library</strong> şeklindeki 
	referansı eklemeniz gerekecektir. Bendeki aşağıdaki gibi Access 2016 olduğu 
	için 16 versiyonunu ekledim, sizde bu biraz daha farklı olabilecektir. (Not: 
	"Microsoft Access 16.0 Object Library" olan kütüphane&nbsp;Access nesne modeline erişmemizi sağlar, 
	buna dikkat edin, isim benzerliği karışıklık yaratabilir, biz 
	onu eklemiyoruz, "Microsoft Office 16.0 Access Database Engine Object 
	Library"'sini ekliyoruz)</p>
	<p><img src="../../images/vbavt1.jpg"></p>
	<h3>Nesneler</h3>
	<p>DAO’daki nesneler veritabanına bağlanmak, dataya erişmek ve veritabanının 
	yapısını değiştirmek için kullanılır. En tepede <strong>DBEngine</strong> 
	nesnesi olan hiyerarşik yapının genel görünümü aşağıdaki gibidir.</p>
	<h4>DAO Nesne Modeli</h4>
	<p><img src="../../images/vbavt2.jpg"></p>
	<p>Biz bunlardan DBEngine ve Workspace’i hiç kullanmayacağız. Bu yüzden 
	yaptığımız tüm database işlemleri default workspace üzerinde olmuş olacak. 
	Farklı oturumlarda farklı workspace açma ihtiyacınız olursa
	<a href="https://msdn.microsoft.com/en-us/library/office/ff822782.aspx">
	buradan</a> detay bilgi edinebilirsiniz.</p>
	<p>DAO ile çalışırken genel süreç şöyledir:</p>
	<ul>
		<li>Database ve recordset nesnesi tanımlanır, </li>
		<li>DB ataması yapılır, </li>
		<li>DB üzerinden recordset yaratılır,</li>
		<li>Sonrasında dataya erişilir,</li>
		<li>Tüm işlemler bitince recordset ve DB Nothing atanıp kapatılır</li>
	</ul>
	<h4>Tanımlamalar</h4>
	<p>DAO ve ADO’nun bazı ortak nesneleri var. O yüzden özellikle iki referansı 
	da birden kullanıyorsanız mutlaka referans(library) ismini nesnelerin önünde 
	kullanmanız gerekir, yoksa karışıklık çıkar ve hata alırsınız. Ancak aynı 
	ismi kullanmayan nesneler için referans belirtmeye gerek yok.</p>
	<pre class="brush:vb">Dim db As DAO.Database '(Bunda DAO’ya gerek yok çünkü ADO’da Database nesnesi yok, karışma olmaz)
ama
Dim rs As DAO.Recordset 'ya da Dim rs As ADODB.Recordset</pre>
	<p>"DAO"’yu yazmak intellisensin çıkması adına bi kolaylık sağladığı için ben size bunu sürekli kullanmanızı(kafa karışıklığı olmayan durumlarda bile) tavsiye ederim.</p>
	<h4>Database nesnesi</h4>
	<p>Öncelikle <strong>Dim db As DAO.Database</strong> diye tanımladık.
	<a href="Ileriseviyekonular_ObjelerDunyasi.aspx#newkeyword">
	New</a> ifadesi olmadan tanımlama yapılır. Zira yaratımını bir fonksiyon ile 
	yapacağız.</p>
	<p>Bazı kaynaklarda bunun arkadasından Set db = DBEngine(0)(0) diye bir kod 
	geldiğini görebilirsiniz. Bu “Workspaces(0).Databases(0)” yazmanın kısa yoludur 
	ama yukarıda belirttiğim gibi biz ikisini de kullanmıycaz, zaten hep default 
	Wokspace(yani 0 indeksli) üzerinde çalışıyor olucaz.</p>
	<p>Bundan sonra gelen kod ise şöyle bir şey olacaktır.</p>
	<pre class="brush:vb">Dim db As DAO.Database
Set db = DAO.OpenDatabase(dbİsmi)</pre>
	<p>Buradaki metodun <span class="keywordler">OpenDatabase</span> olması sizi 
	yanıltmasın, gerçekte bir Access penceresi açılmamaktadır. İlgili 
	database’in bir nevi hafızada açıldığını düşünebilirsiniz.</p>
	<p>Bundan sonrasında bu nesneyle ilgili olarak başka bir işimiz olmayacak. 
	Aslında Database nesnesinin <strong>CreateTableDef</strong>, <strong>
	CreateQueryDef</strong> gibi tablo ve sorgu yaratmaya yarayan metodları var 
	ama ben bunların Excel VBA içinden kullanılması gerektiğini düşünmüyorum. 
	Bizim işimiz daha çok Accesten data okumak ve gerekirse tablolarda 
	güncelleme yapmak, kayıt eklemek, silmek olacaktır. İhtiyacımız olan tablo 
	ve sorguları zaten Access üzerinde yaparız diye düşünüyorum. O yüzden bu tür 
	metodlara değinmeyeceğim. Ender de olarak ihtiyacınız olursa bunlarla ilgili 
	makaleleri Google’da kolaylıkla bulabilirsiniz. (Belki Access VBA ile ilgili 
	bir sayfada bu konuda örnekler yapmayı düşünebilirim. Access VBA’i, Excel VBA 
	kadar sık kullanmasak da zaman zaman oldukça faydasını görmekteyim.)</p>
	<p>Veritabanıyla işimiz bitince <span class="keywordler">Close</span> 
	metodunu kullanarak bağlantıyı kapatırız ve son olarak Nothing ataması ile 
	belleği boşaltırız. Özetle;</p>
	<pre class="brush:vb">Dim db As DAO.Database
Set db = DAO.OpenDatabase("………..accdb")
'diğer kodlar
db.Close
Set db=Nothing</pre>
	<h4>TableDef</h4>
	<p>DAO’da tablolarla ilgili iki nesne bulunur; <span class="keywordler">
	TableDef(s)</span> ve <span class="keywordler">Recordset</span>.<br>TableDef(s) 
	tablolar hakkında metadata sunar. Alanlar, indexler, tablonun adı v.s gibi 
	işlemler için kullanılır. Tabloların yaratılması ve silinmesi de bununla 
	yapılır. İçindeki dataya erişim ise RecordSet ile yapılır. Ancak yukarda 
	belirttiğim gibi biz tabloyla ilgili genel işlere burada girmeyeceğiz. Bu 
	konuda bilgi lazım olursa yine bir google search yapabilirsiniz.</p>
	<p>Keza, Sorgu yaratma gibi sorgu işlemlerinin nesnesi olan QueryDef(s) ile de 
	çok bi işimiz olmayacak. Bununla beraber bir sorgunun içini okumak istersek 
	yine RecordSet nesnesini kullanırız.</p>
	<h3>Recordset nesnesi</h3>
	<p>DAO’da en çok bu nesneyle haşır neşir olacağız. Çünkü data okuma ve 
	manipülasyonu bu nesne ile yapılır. O yüzden bunu "Nesneler" başlığı altında 
	incelemek yerine ayrı bir başlık altında incelemenin daha doğru olduğunu 
	düşündüm. </p>
	<p>Recordset nesnesi bize adından anlaşılacağı üzere belirli bir kayıt seti 
	verir. Bu tüm bir tablo olabileceği gibi çeşitli filtreler uygulanmış bir 
	sorgu sonucu da olabilir.</p>
	<h4>Tanımlama ve Yaratma</h4>
	<p>RecordSet'i tanımlama klasik değişken tanımıyla yapılır ancak ataması 
	yapılırken <a href="Ileriseviyekonular_ObjelerDunyasi.aspx#newkeyword">
	New</a> kelimesi kullanılmaz, zira bunu başka bir objenin(genelde DB 
	objesinin) bir metodundan dönen değerle elde 
	edeceğiz. </p>
	<p>4 çeşit yaratma şekli vardır: DB’den, 
	table'dan, query'den ve başka bir recordsetten. Önce genel syntax'a sonra 
	parametrelere bakalım.</p>
	<p><strong>Syntax:Object.OpenRecordset(Name, [Type], [Options], [LockEdit]).</strong></p>
	<p><strong>Name</strong> olarak tablo adı sorgu adı girilebileceği gibi SQL de girilebilir.</p>
	<pre class="brush:vb">Dim rs As DAO.Recordset
Set rs = dbobj.OpenRecordset(Type, Options) 
Set rs = TableDefObject.OpenRecordset(Type, Options) 
Set rs = QueryDef.OpenRecordset
Set rs = RecordsetObject.OpenRecordset(Type, Options) 'varolan rs’de ilave filtre için</pre>
	<p>Type'a az sonra detaylı bakacağız. Son iki paremetreyi ise neredeyse hiç 
	kullanmayacağız. Bunların bir kısmının geriye dönük uyumluluk içeren 
	paremetreler olup bi kısmı ise küçük uygulamlarda çok kullanılmayan 
	özelliklerdir, daha büyük uygulamalar için zaten Access yerine diğer 
	Veritabanı uygulamaları kullanılmadlır. Bir şekilde kullanım ihtiyacı 
	olursa(Ör:aynı anda iki kişinin güncelleme yapması durumundaki davranışı 
	belirlemek isterseniz) google'da araştırabilirsiniz. Ben hiç ihtiyaç 
	duymadığım için araştırıp öğrenme zahmetine de girmedim açıkçası, o yüzden 
	size de anlatamıyorum. Şimdi gelelim Type'a.</p>
	<h5>Type parametresi: DAO Recordset Tipleri</h5>
	<p>5 tür tip vardır.</p>
	<ol>
		<li><strong>Table-type&nbsp;recordset(dbOpenTable)</strong>: Bunlar, düz 
		tablolara dayanır, yani bu tiple sorgular ve linkli tablolar okunmaz. 
		Lokal tablolar için varsayılan tip budur. Sadece düz kayıt okuma veya 
		kayıt ekleme/güncelleme yapacaksanız bunu kullanabilirsiniz ancak diğer işlemlerde bazı 
		metod ve propertyler(AbsolutePosition, FindFirst v.s) çalışmadığı için 
		bu işlemlere ihtiyacınız olduğunda bunu kullanamazsınız. Bu tiple yaratım yapıldığında, kayıt bulmak 
		için <span class="keywordler">Seek</span> metodu kullanılabilir ama
		<span class="keywordler">Find</span> ve türevleri kullanılamaz. Seek 
		kullanımı için indexlerden yararlanırlır, bu yüzden Find metodundan daha 
		hızlıdır. </li>
		<li><strong>Dynaset-Type(dbOpenDynaset)</strong>:&nbsp;Tablolara ek olarak 
		sorgularda ve linkli tablolarda da kullanılır. Linkli tablolar için 
		varsayılan tip budur. Kayıt bulmak için <strong>Find</strong> metodu kullanılırken 
		<strong>Seek</strong> metodu bunda kullanılamaz.</li>
		<li><strong>Snapshot-type(dbOpenSnapshot)</strong>: Bi recordset elde 
		edilmiş ve resmi çekilmiştir, bunun üzerinden kayıt okumak için bu tip 
		kullanılır. Statik bir veri setine sahip olduğumuz için kayıtlarda güncelleme yapılamaz, 
		yani read-only bir yöntemdir. <strong>Find</strong> metodunu 
		destekler.</li>
		<li><strong>Forward-only-type(dbOpenForwardOnly)</strong>: Snapshota çok 
		benzer, sadece ileri doğru okuma yapar.</li>
		<li><strong>Dynamic-type(dbOpenDynamic):</strong> DynaSet’e çok benzer. 
		Farkı şu: O sırada başka kullanıcılar da recordsetiniz için temel 
		aldığınız tabloda bir güncelleme/ekleme/silme işlemi yaptıysa bunlar da 
		sizin recordsetinize anında yansır.</li>
	</ol>
	<p>Görüldüğü üzere Seek metodunu sadece dbOpenTable tipinde açılmış 
	recordsetlerde kullanabilir. Bu bağlamda örnek bir veritabanı erişim koduna bakalım.</p>
	<pre class="brush:vb">Sub DAOOrnek()
Dim db As DAO.Database
Dim rs1 As DAO.Recordset, rs2 As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs1 = db.OpenRecordset("Data", dbOpenTable)
Set rs2 = rs1.OpenRecordset
End Sub</pre>
	<p>Tip belirtilmezse, default tipler baz alınır: "Name" olarak verilen 
	kaynak, bir tablo ise dbOpenTable, linkli tablo ise dbOpenDynaset.</p>
	<h5>Hangi Tip ne zaman kullanılır?</h5>
	<ul>
		<li>Düz okuma, kayıt ekleme, güncelleme yapılacaksa ve/veya Seek ile 
		hızlı arama yapma ihtiyacı varsa(indeks bulunmalı):<strong>dbOpenTable</strong></li>
		<li>Query ve linkili tablolarda yeni kayıt ekleme, güncelleme, silme + 
		ayrıca Find türevlerini kullanma ihtiyacı varsa:<strong>dbOpenDynaset 
		veya dbOpenDynamic(farkları yukarıda var)</strong></li>
		<li>Küçük veri kümelerinde sadece okuma(ileri geri farketmez) 
		yapacaksanız ve hızlı kayıt arama ihtiyacı yoksa(yani seek 
		kullanmayacaksanız): <strong>dbOpenSnapshot</strong></li>
		<li>Küçük veri kümelerinde sadece
		<span style="text-decoration: underline">ileri okuma</span> yapacaksanız 
		ve hızlı kayıt arama ihtiyacı yoksa(yani seek kullanmayacaksanız) <strong>
		dbOpenForwardOnly</strong></li>
	</ul>
	<h4>Field nesnesi(Alanlar)</h4>
	<p>(Field nesnesi recordsetten bağımsız bir nesne olmakta birlikte hep 
	onunla kullanıldığı için ayrı bir kısım açmak yerine Recordset kısmı altında 
	ele almak istedim.) Önceki kısımlarda belirttiğim gibi, Field yaratma, 
	bunlarda index belirleme gibi konulara girmiyoruz. Bu tür işlemleri VBA 
	içinden yapmak yerine doğrudan Access'te yaparız, zira bunlar genelde 
	dinamik olarak değiştirilecek şeyler değildir. Olur da ihtiyaç duyarsanız 
	MSDN veya googleda bunlara ulaşmak oldukça kolay. Bizim işimiz daha çok bu 
	alanlara erişmek olacak. Erişimin de 3 yolu bulunmaktadır.</p>
	<ul>
		<li>Alan adı ile:rs.Fields("İsim")</li>
		<li>Alan adı kısayolu ile:rs![İsim]</li>
		<li>Alan item no ile:rs.Fields(1)</li>
	</ul>
	<p>rs.Fields(0).<strong>Name</strong>: İlk kolonun adını yani kolon 
	başlığını getirir.<br>rs.Fields(0).<strong>Value</strong>: Bu ise ilk 
	kolondaki geçerli kaydın(satırın) içeriğini döndürür. Bu arada Value 
	özelliği default özellik olup yazılmasa da olur, ama biz iyi bir 
	programlamacı olup yazıyoruz.<br>rs.Fields.<strong>Count</strong>: İlgili 
	kayıt setindeki alan sayısını verir.</p>
	<h4>Yeni kayıt, mevcut kaydı düzenleme ve kaydetme</h4>
	<p><span class="keywordler">AddNew</span> metodu yeni boş bir satır ekler. 
	Sonra Field nesnesi ele alınarak ilgili alan atamaları yapılır. Normalde 
	Acceste manuel kayıt ekledikten veya değiştirdikten sonra onu kaydetme(save 
	etme) diye birşey yoktur ancak VBA’de ismi “Save” olmasa bile bi kaydetme 
	işlemi var, onu da <span class="keywordler">Update</span> metodu ile 
	yapıyoruz</p>
	<p>Bir recordsete yeni kayıt eklendiğinde geçerli(aktif) kayıt otomatikman 
	yeni eklenen kayıt olmaz. Bunun için yeni kayda çapa atarak erişmemiz 
	gerekir. Bunla ilgili detayları az aşağıda göreceğiz. </p>
	<pre class="brush:vb">rs.AddNew
'ilgili alan atamaları yapılır
rs.Update 'Kayıt işlemi gerçekleşir
rs.Bookmark = rs.LastModified 'çapayı attık, şimdi yeni kayıt üzerinde çalışabiliriz</pre>
	<p><span class="keywordler">Edit</span> metodu kaydı değiştirir(Accesteki 
	Update sorgusunun muadilidir). Az önce belirttiğim gibi Update metodu 
	yapılan değişkliklerin yansımasını sağlar, Update Query ile 
	karıştırılmaması lazım. Yani isim benzerliği kafanızı karıştırmasın. 
	Özetle; Accessteki Update Sorgu işlemi DAO'nun Edit metodu ile yapılırken, Acceste otomatik 
	gerçekleşen Save işlemi 
	DAO'nun 
	Update metodu ile yapılır.</p>
	<pre class="brush:vb">'yeni kayıt
rs.AddNew 
rs.Fields(0) = "3333"
rs.Fields(1) = "Aksaray"
rs!Durum = "Açık" 'bu !'li yazım "."lı yazıma alternatif yöntemdir
rs.Fields("Bölge kodu") = 7030
rs.Update 'bunu demeden kayıt eklenmez
&nbsp;
'Editleme
rs.Edit
rs.Fields("Bölge kodu") = 5555
rs.Update 'bunu demeden update etmez</pre>
	<h4>Silme</h4>
	<p>Kayıtlarda dolaşırken cursor'ın bulunduğu kaydı silmek için
	<span class="keywordler">Delete</span> metodunu kullanıyoruz. Silme 
	işlemini yapabilmek için recordseti Table veya Dynaset tipinde açmış 
	olmak gerekiyor, aksi hade hata alınır.</p>
	<p>Silme işlemi soucunda sonraki kayıt otomatikman geçerli kayıt olmaz, o 
	yüzden ilgili işlemden sonra MoveNext yapmanız gerekir.&nbsp;</p>
	<p>Aşağıdaki örnekte Durum kodu 0 olan kayıtlar siliniyor.</p>
	<pre class="brush:vb">
'ön tanımlar
If Not (rs.EOF And rs.BOF) Then

    Do While Not rs.EOF
     Durum = rs.Fields(3).Value
        If Durum = 0 Then
            rs.Delete
        End If

        rs.MoveNext
    Loop

End If	</pre>
	<p>Silme işlemini, belli kriterleri sağlayan kayıtlar için yapacaksanız ben 
	SQL metni çalıştırmanızı(SQL biliyorsanız tabi) tavsiye ederim. (Bunun detaylarını
	<a href="#executesql">aşağıda</a> göreceğiz.)</p>
	<p>Mesela yukardaki örnekte kayıtları silmek için şu kodu çalıştırmak bana 
	daha pratik geliyor.</p>
	<pre class="brush:vb">'ön tanımlar
db.Execute "Delete from tabloadı where Durum=0"</pre>
	<h4>Kayıtlarda dolaşma </h4>
	<h5>Move metodu ile belli satırlardaki kayıtlara konumlanma</h5>
	<p><span class="keywordler">Move</span> metodu, belirli bir satır 
	numarasıyla kullanılabileceği gibi Move’un türevleri şeklinde de 
	kullanılabilir.</p>
	<pre class="brush:vb">Rs.Move 10 '10 kayıt aşağı konumlanır, 10.kayda değil(negatif olursa geriye doğru hareket)
Rs.Move 0 'olduğu yerde kalır
Rs.MoveFirst 'ilk kayda konumlanır
Rs.MoveLast 'son kayda konumlanır
Rs.MoveNext 'bir sonraki kayda konumlanır. Özellikle döngülerde satır satır ilerlerken kullanılır.
Rs.MovePrevious 'bir önceki kayda konumlanır</pre>
	<p>Bunları kullanırken BOF ve EOF ile birlikte kullanımı tavsiye edilir.</p>
	<h5>BOF &amp; EOF</h5>
	<p><span class="keywordler">BOF</span>, ilk kayıttan önceki bir pozisyonda 
	olup olmadığınızı, <span class="keywordler">EOF</span> da son kayıttan 
	sonraki bir pozisyonda olup olmadığınızı gösterir. Kullanım amacı da tabloda 
	hareket ederken tablonun sınırları içinde kalıp kalmadığınızı görmektir.</p>
	<p>İkisi de True olursa geçerli kayıt yok demektir. O yüzden bir recordset 
	içinde kayıt olup olmadığını her ikisinin de False olması veya Not True 
	olması şeklinde aşağıdaki gibi test ederiz.</p>
	<pre class="brush:vb">If Not (rs.EOF And rs.BOF) Then
'kodlar
End If

If crst.EOF=False And rst.BOF=False Then
'kodlar
End If</pre>
	<h5>RecordCount</h5>
	<p>BOF ve EOF'un amacı, RecordSet içinde kayıt olup olmadığını anlamaktan 
	ziyade(dolaylı olarak bu amaca da hizmet edebilirler), hareket sonrasında tablonun dışına çıkıp çıkmadığımızdır. 
	Recordsette kayıt olup olmadığını görmek için <span class="keywordler">
	RecordCount</span> özelliğini kullanıyoruz.</p>
	<p><span class="dikkat">Dikkat</span>:Recordsetimizi TableType tipinde 
	açtıysak, RecordCount sorunsuz çalışır. Ancak DynasetType veya diğer 
	tiplerde açtıysak 
	RecordCount&nbsp;property’si o ana kadar erişilen kayıtların sayısını getirir; ve 
	bu tipte açılan recordsetlerde ilk kayda gelindiğinde VBA kodu okunmaya devam 
	eder.&nbsp;Bu yüzden recordseti açtıktan hemen sonra kayıt sayısını elde etmeye 
	çalışırsak <strong>sonuç hep 1 döner</strong>. Bunun için <strong>Dynaset</strong> tipinde(ve 
	diğerlerinde) açıldığında
	<strong>önce</strong> <strong>MoveLast</strong> ile son kayda konumlanmalı ondan sonra RecordCount'ı elde etmeye 
	çalışmalıyız. Bununla birlikte büyük tablolarda bu yöntem çok vakit alan bi 
	iş olabilir, o yüzden dikkatli kullanılmalıdır. </p>
	<p>Bu noktada ilk önerim, eğer başka nedenlerle gerekli değilse tabloyu 
	DynasetType tipinde(veya diğer tiplerde) değil TableType tipinde açın. Diyelim ki DynasetType 
	tipinde açtık, o zaman ikinci önerim de şudur: Eğer amacınız gerçekten kayıt 
	sayısını elde etmekse MoveLast ise son kayda gidin(büyük tablolarda 
	performans sorunu yaşatabilir) ama amacınız “içerde 
	kayıt varmı yok mu” diye bakmaksa sonucun 1 dönmesi yeterlidir, MoveLast’a 
	gerek yoktur, bu yüzden sade bir “RecordCount &gt;0 mı?” kontrolü yaparsınız, o kadar.</p>
	<h5>BookMark, LastModified</h5>
	<p>Daha sonra dönmek üzere geride bıraktığınız bir kayda bookmark aracılığı 
	ile ulaşabilirsiniz. Hem okunur hem yazılır bir özelliktir.<br>Geçerli kaydı 
	bir Bookmark olarak atamak için, önceden kaydedilmiş bir bookmark 
	kullanılabileceği gibi <span class="keywordler">LastModified </span>özelliği 
	ile en son değişitirilmiş/eklenmiş kayıt da atanabilir. Aşağıdaki örnekte
	<span class="keywordler">AbsolutePosition</span> özelliği de kullanılmış 
	olup kaydın o anki satır numarasını verir. AbsolutePosition, dbOpenTable 
	tipinde çalışmaz, dbOpenDynaset olmalı.</p>
	<pre class="brush:vb">Sub dao_bookmark()
Dim db As DAO.Database
Dim rs As DAO.Recordset
&nbsp;
Set db = DAO.OpenDatabase("….\daodeneme.accdb")
Set rs = db.OpenRecordset("şubeler", dbOpenDynaset)
&nbsp;
Debug.Print rs.AbsolutePosition
rs.MoveNext
Debug.Print rs.AbsolutePosition
rs.MoveNext
Debug.Print rs.AbsolutePosition
rs.MoveNext
Debug.Print rs.AbsolutePosition
x = rs.Bookmark 'çapa atıyoruz
Debug.Print rs.AbsolutePosition
rs.Move 500 'çeşitli işlemler sonucunda şuan 500 kayıt aşağı geldik diyelim
Debug.Print rs.AbsolutePosition
rs.Bookmark = x 'tekrar çapamıza dönüyoruz
Debug.Print rs.AbsolutePosition
&nbsp;
'son değişen kayıt
rs.Move 80
rs.Edit
rs.Fields("Bölge kodu") = 3333
rs.Update
Debug.Print rs.AbsolutePosition
rs.Move 200
Debug.Print rs.AbsolutePosition
rs.Bookmark = rs.LastModified
Debug.Print rs.AbsolutePosition
&nbsp;
rs.Close
db.Close
&nbsp;
Set rs = Nothing
Set db = Nothing
End Sub</pre>
	<h5>Döngüsel Örnek</h5>
	<p>Gerçek dünyada kayıtlarda döngüsel olarak dolaşmak pek daha olasıdır. O 
	yüzden şimdi bir de döngüsel örnek yapalım.</p>
	<pre class="brush:vb">
Sub dao_tablolardadolas()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim tdf As DAO.TableDef

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data")

'başlangıçta kaç kayıt var bakalım
Debug.Print rs.RecordCount

'önce geçici kayıt ekleyelim
For i = 1 To 3
    rs.AddNew
    rs.Fields(0) = "deneme bölge" &amp; i
    rs.Fields(1) = "deneme şube" &amp; i
    rs.Fields(2) = "deneme ürün" &amp; i
    rs.Fields(3) = 3 - i
    rs.Update
Next i
Debug.Print rs.RecordCount

If Not (rs.EOF And rs.BOF) Then

    Do While Not rs.EOF
     Durum = rs.Fields(3).Value
        If Durum = 0 Then
            rs.Delete 'Durum=0 olan kayıtlar silinir
        End If

        rs.MoveNext
    Loop

End If
Debug.Print rs.RecordCount
End Sub

</pre>
	<h4>Seek ve Find<span style="text-decoration: underline"><em>ek</em></span> ile kayıt arama</h4>
	<p>DAO, kayıtları bulmanın iki 
	yolunu bize sunuyor. <span class="keywordler">Seek</span> ve 
	<span class="keywordler">Find</span> türevleri. Bunların kullanımı, Recordset 
	oluşturulurken kullanılan tipe göre değişmektedir. Öncelikle şu ayrımı iyi 
	yapmak gerekiyor. Ne zaman Recordset, ne zaman SQL, ne zaman diğer metodlar?</p>
	<ul>
		<li>Eğer yapabiliyorsak recordsetimizi direkt aradığımız kayıt üzerine 
	oluşturmalıyız. Yani bir SQL ile tek satır döndüren bir recordset 
	tanımlayabiliriz.</li>
		<li>Eğer daha geniş kümeli bir recordsetimiz olacak 
	ve bunu çeşitli aşamalarda farklı şekillerde filtrelemeye/araştırmaya tabi 
	tutacaksak<ul>
			<li>Recordsetin recordsetini yapabileceğimiz gibi</li>
			<li>Aşağıaki diğer arama yöntemlerini kullanabiliriz<ul>
				<li>Seek</li>
				<li>Find 
	türevleri</li>
				<li>SQL</li>
			</ul>
			</li>
		</ul>
		</li>
	</ul>
	<h5>Seek Metodu</h5>
	<p>Recordset tipi olarak <strong>sadece TableType</strong> 
	seçildiyse kullanılabilir. Çünkü tablodaki indekslere ihtiyaç duyar ve doğal 
	olarak da kolonlardan en az birinde indeks olması gerekir. O yüzden Seek 
	metodu uygulanmadan önce Indeks property’si belirtilir. İndeksli arama da en 
	hızlı yöntem olduğu için Find’a göre daha hızlı bir yöntemdir. Tabiki 
	Acceste ilgili tabloda ilgili kolonda indeks olduğundan emin olmalısınız. Bu 
	arada PrimaryKey dışındaki kolonlarda indeks adı genelde kolonadı ile aynı 
	olurken, PrimaryKey olan bir kolonda indeks adı "PrimaryKey" olur, o yüzden 
	indeks olarak da bu şekilde belirtmelisiniz.</p>
	<p>Parametre 
	olarak "=","&lt;","&gt;" gibi karşılaştırma işaretleri ve aranan değer girilir.</p>
	<p>Aradığımız değere konumlanma girişiminin başarılı olup olmadığını
	<span class="keywordler">NoMatch</span> özelliği ile 
	test ederiz. Eğer True dönerse aranan kriterlere göre uygun kayıt 
	bulunamamıştır demektir.</p>
	
	<pre class="brush:vb">
Sub dao_seek()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("Data", dbOpenTable)

rs.Index = "Şube Adı" 'index belirtiyoruz
rs.Seek "=", "Şube115"
If rs.NoMatch Then 'konumlanma başaraılı mı diye kontrol ediyoruz
    Debug.Print "Kayıt bulunamadı"
Else
    Debug.Print rs.Fields("Aylık Gerç")
End If

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>
	
	<h5>Find türevleri</h5>
	<p>Move’un aksine Find’ın solo halde bir versiyonu yoktur. 4 
	çeşidi vardır. <span class="keywordler">FindFirst</span>, 
	<span class="keywordler">FindPrevious</span>, <span class="keywordler">FindNext</span>, and 
	<span class="keywordler">FindLast</span>. Find 
	kullanımı için Recordsetimizin Tabletype dışındaki bir tiple tanımlanması 
	gerekir. Genellikle Dynaset yeterlidir.</p>
	<p>Indeks kullanmak zorunda 
	değildir, bu yüzden indekssiz bir kolonda arama yaptığınızda Seek metoduna 
	göre çok daha yavaş çalışır.<br>Find metodlarıyla "?" ve "*" gibi joker 
	karakterleri kullanabiliyoruz.</p>
	<pre class="brush:vb">
Sub dao_find()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenDynaset)

rs.FindFirst "[Şube Adı] LIKE '*deneme*'"

If rs.NoMatch Then
    Debug.Print "Kayıt bulunamadı"
Else
    Debug.Print rs.Fields("Şube Adı")
End If
rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub	</pre>
	<p>Ben bu metodların bi karşılaştırmasını yaptım. Buna göre 5 
	milyon kayıtlık bir tabloda;
	</p>
	<ul>
		<li>indekssiz bir kolonda FindFirst araması 
	yapmak 127 sn,</li>
		<li>indeksli kolonda FindFirst yapmak 1,5 sn,</li>
		<li>(İndeksli kolonda) Seek 
	ile arama yapmak ise 0,04 sn sürüyor. </li>
	</ul>
	<p>Gördüğünüz gibi indeksli olması 
	her halükarda hızı inanılmaz arttrıyor, ancak Seek’in Find’a göre üstünlüğü 
	ise aşikar. Mutlak değer olarak fazla bir fark olmasa da oransal fark çok 
	büyük.</p>
	<h5>SQL<br></h5>
	<p>Aranan değeri bulmada bir diğer yöntem, Recordseti çekerken 
	tek sonuç döndürecek bir SQL çalıştırmaktır. Veya çoklu sonuç dönecekse de 
	MoveFirst diyerek ilk kaydın sonucunu almak olacaktır. Veya zaten çok sonuç 
	arıyorsak da çoklu sonuç dönen bir SQL hazırlanır.</p>
	<h3>Filter ile Recordseti filtreleme</h3>
	<p>Her ne kadar filtreleme ve sıralama işlemlerini SQL metni içnde 
	yapmamızda fayda olsa da bazen recordset üzerinden de bunları yapmamız 
	gerekebillir. Biz burada sadece Filtreleme işlemine bakacağız.</p>
	<p>Öncelikle belirtmek isterim ki <span class="keywordler">Filter</span> işlemini dbOpenTable tipinde 
	açılmış bir recodsette yapamıyoruz. Diğer tiplerde açılmış olması gerekir.</p>
	<p>Bu işlem için tahmin edileceği üzere Filter property'si kullanılır. Bunun 
	içine 
	SQL metninde yazacağımız gibi bir kriter yazarız. Eğer ki kolon adımız 1'den 
	çok kelimeden oluşuyorsa bunları [] içine yazarız.</p>
	<p>Aşağıda ADO kısmında göreceksiniz, orda da Filter işlemi yapılıyor ancak 
	DAO'da ADO'dan farklı olarak uygulanış şekli biraz farklıdır. DAO'da iki 
	farklı recordsetimizin olması gerekir. İlk recordsetin Filter property'sine 
	kriterler girilir ve ikinci recordset bu ilk recordsetten filtrelenmiş 
	şekilde elde edilir. Hemen örneğimize bakalım.</p>
	<pre class="brush:vb">
Sub dao_filter()
Dim db As dao.Database
Dim rs As dao.Recordset
Dim rsFilter As dao.Recordset

Set db = dao.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenDynaset)

rs.Filter = "[Ürün Adı]='Ürün1'"
Set rsFilter = rs.OpenRecordset

rs.MoveLast
Debug.Print rs.RecordCount '1868

rsFilter.MoveLast
Debug.Print rsFilter.RecordCount '466
End Sub</pre>
	<p>NOT:Tarihsel alanları mutlaka Amerikan formatında(ay-gün-yıl) girilmesi 
	gerekiyor. </p>
	<h3 id="executesql">Execute ile 
	Sorgu/SQL çalıştırma</h3>
	<p><span class="keywordler">Execute</span> metodu ile doğrudan basit bir SQL veya varolan bir 
	eylem sorgusu(Append,Delete,Update) çalıştırılabilir. Hem Database nesnesi 
	hem de QueryDef nesnesi için kullanılabilen bir metoddur.</p>
	<pre class="brush:vb">
Sub Dao_Execute()
Dim db As DAO.Database

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")

db.Execute "Query1" 'veya açık bir şekilde SQL metni
' veya db.QueryDefs(0).Execute

db.Close

Set db = Nothing

End Sub</pre>
	<p>Execute’la birlikte 
	kullanılan parametreler var. Biz bunlara burada girmeyceğiz, detay bilgi 
	edinmek istiyorsanız 
	<a href="https://msdn.microsoft.com/en-us/library/office/ff197654.aspx">
	şuraya</a> 
	bakabilirsiniz.</p>
	<p>Ayırca yukarıda belirttiğimiz gibi, Filter işlemlerini mümkün olduğunca 
	SQL içinde çalıştırmak daha hızlı sonuç almamızı sağlar, özellikle büyük 
	veri kümelerinde. Eğer ki bu elde ettiğimiz veri setinde, farklı case'lere 
	göre dinamik filtrelemeler yapmak gerekirse o zaman Filter'ı devreye 
	sokabiliriz.</p>
	<h3>Datayı Excel’e almak(Import işlemi)</h3>
	<p>DAO 
	kullanımında en sık yapacağımız işlem, eriştiğimiz datayı Excel içine almak 
	olacaktır. Bunun için de birkaç yöntem bulunuyor.</p>
	<h4>1.yöntem: <span class="keywordler">CopyFromRecordset </span>metodu</h4>
	<p>Bu yöntem en hızlı yöntemdir. Çoğu durumda bu yeterli 
	olmaktadır.</p>
	<p><strong>Syntax: </strong>Range.CopyFromRecordset(Data, [MaxRows], 
	[MaxColumns])</p>
	<p>Burda önemli olan husus, başlıkların gelmiyor oluşudur. Başlık için 
	döngüsel bir kod yazılır. MaxRows ve MaxColumns ile çekilen kayıt sayısı 
	sınırlandırılabilir.</p>
	<pre class="brush:vb">
Sub import1()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenTable)

If Not (rs.EOF And rs.BOF) Then
    'başlık yazma kısmı
    For i = 0 To rs.Fields.Count - 1
        ActiveCell.Offset(0, i).Value = rs.Fields(i).Name
    Next i
    'şimdi de data yazılır
    ActiveCell.Offset(1, 0).Select
    ActiveCell.CopyFromRecordset rs
End If

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>
	<h4>2.yöntem: Diziye atayıp diziyi yazdırmak</h4>
	<p>Bu yöntem 2. en hızlı yöntemdir. Eğer 
	diziye atadıktan sonra diziyi başka yerde de kullanacaksanız veya dizi 
	elemanları üzerinde işlem yaptıktan sonra Excel'e aktaracaksanız bu yöntemi 
	kullanabilirsiniz.</p>
	<pre class="brush:vb">
Sub import2()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim hucreler() As Variant
Dim alan As Range


Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenTable)
ReDim hucreler(rs.RecordCount - 1, rs.Fields.Count - 1)

'başlık yazma kısmı
For i = 0 To rs.Fields.Count - 1
    ActiveCell.Offset(0, i).Value = rs.Fields(i).Name
Next i

'diziyi dolduralım
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            hucreler(j, i) = rs.Fields(i).Value
        Next i
    
        j = j + 1
        rs.MoveNext
    Loop
End If

'excele yazalım
Set alan = Range("A2").Resize(UBound(hucreler, 1), UBound(hucreler, 2) + 1)
alan.Value = hucreler

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>
	<h4>3.yöntem: GetRows metodu ile dizi elde ederek</h4>
	<p><span class="keywordler">GetRows</span> ile iki boyutlu bir dizi elde ederiz. Boyutlardan ilki kolonu, ikincisi satır 
	numarasını ifade eder. Tek kolonluk bir veri çekseniz bile 2 boyutlu bir diziniz olur.</p>
	<p>Rs.GetRows tüm recordseti döndürürken Rs.GetRows(x) 
	ilk x kaydı diziye aktarır. Mesela ilgili veri kümesinden sadece örnek bir 
	küme almak istiyorsanız 100 satırlık data çekebilirsiniz. GetRows, Move 
	metodu gibi davranır, yani parametre olarak 100 derseniz 100. kayda gelir. 
	</p>
	<pre class="brush:vb">
Sub import3()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim dizi As Variant

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenTable)


If Not (rs.EOF And rs.BOF) Then
    dizi = rs.GetRows(10) 'ilk 10 kayıt
    Debug.Print dizi(0, 0) 'ilk kolon ilk satır
    Debug.Print dizi(rs.Fields.Count - 1, UBound(dizi, 2)) 'son kolon son satır
End If

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>

<p>Yukarıda belirttiğim gibi, tek kolon çeken bir SQL'iniz bile olsa iki boyutlu 
bir dizi elde edersiniz. Diyelim ki böyle bir veri çektiniz ve ihtiyacınız da, bu veri setini
aralarında ";" işareti olacak şekilde birleştirmek. Bunu döngüsel olarak dolaşıp yapabileceğiniz gibi
WorksheetFunction.Index fonksiyonundan da yararlanabilirsiniz.</p>
<p>Örneğin, diyelimki çektiğiniz veri seti bazı müşteri numaları olsun. Bunları aralarında ";" olacak 
şekilde birleştirmek için şöyle bir kod yazabiliriz:</p>
<pre class="brush:vb">
'önceki kodlar
rs.Open strSQL,con,adOpenStatic,adLockOptimistic
müşteriler=rs.GetRows
müşteriStr=Join(WorksheetFunction.Index(müşteriler,0),";")
'sonraki kodlar
</pre>
	<h4>4.yöntem: Range’e döngüsel şekilde yazdırma </h4>
	<p>En yavaş 
	yöntemdir. Dizilerle çalışmayı bilmiyorsanız veya hücreler üzerinde başka 
	işlemler de yapacaksanız bu yöntemi kullanabilirsiniz ancak büyük data 
	kümelerinde tavsiye edilmez.</p>
	<pre class="brush:vb">
Sub import4()
Dim db As DAO.Database
Dim rs As DAO.Recordset


Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenTable)

'başlık yaz
For i = 0 To rs.Fields.Count - 1
    ActiveCell.Offset(0, i).Value = rs.Fields(i).Name
Next i

If Not (rs.EOF And rs.BOF) Then
    ActiveCell.Offset(1, 0).Select
    k = ActiveCell.Row
    
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            Cells(k, i + 1).Value = rs.Fields(i).Value
        Next i
        k = k + 1
        rs.MoveNext
    Loop
End If

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>
	<h3>Exceldeki datayı Accese atma(Export işlemi)</h3>
	<p>Bunda iki yöntem uygulanabilir.</p>
	<h4>1.Yöntem:Döngü içinde Addnew+update</h4>
	<pre class="brush:vb">
Sub export1()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("exporttable", dbOpenTable)

For k = 2 To [a1].End(xlDown).Row
    rs.AddNew
    For i = 1 To [a1].End(xlToRight).Column
        rs.Fields(i - 1) = Cells(k, i).Value
    Next i
    rs.Update
Next k

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing
End Sub	</pre>
	<h4>2)Access sorgusunu 
	çalıştırmak. </h4>
	<p>Bu yöntem DAO'ya ait bir örnek değildir aslında. Burada Access nesne 
	modeline girmiş oluyoruz. Çok basit bir mantığı var. Aşağıdaki kodda 
	yorumlara bakın lütfen.</p>
	<p>Bu arada bu yöntemde mevcut Excel dosyasını Access'e linklemiş olmak gerekir.</p>
	<pre class="brush:vb">
Sub export2()
Set accessApp = GetObject(adres+"vbadb.accdb", "Access.Application") 'İlgili access dosyasını bir değişkene atıyoruz, ama bunu uygulama olarak atıyoruz
With accessApp
   .Application.Visible = False 'Arka planda çalışsın istiyoruz
   .DoCmd.Openquery "AppendQuery1" 'DoCmd metodu ile kayıtlı bir sorguyu çalıştırıyoruz
   .Run "Modül1" 'Run metodu ile Access VBA ile yazımış bir kodu çalıştırıyoruz
End With
End Sub</pre>
	<p>
	Access sorgusunu çalıştırmanın bir&nbsp; yöntemi de aslında yukarıda 
	gördüğümüz Execute metodudur.</p>
	<pre class="brush:vb">
Sub export3()
Dim db As dao.Database
Dim rs As dao.Recordset

Set db = dao.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("data", dbOpenDynaset) 'açılış tipi dynaset

db.Execute "srg_exceldekiniburayaappend"

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

End Sub</pre>
	<h3>
	Örnek Çalışma - İnteraktif Veri Çekme Formatı</h3>
		<p>Şimdi DAO ile interaktif bir şekilde, yani parametreleri dinamik 
		şekilde değiştirerek nasıl veri çekilir, buna ait bir örnek yapmak 
		istiyorum.</p>
		<p>Bu örnekte, bir tablodan bir bölgenin belli bir aydaki çeşitli 
		ürünlerine ait rakamlarını çekiyoruz. Ürünlerden birine çift 
		tıkladığımızda da şube detayı gösteriliyor. ürüne tekrar çift 
		tıklandığında şubeler kaybolup tekrar ilk haline dönüyor.</p>
		<p>Bir seçim sonunda görünen tablo aşağıdaki gibidir.</p>
		<p><img src="../../images/vbadaointeraktif1.jpg"></p>
		<p>B5 hücresindeki Ürün1'e çift tıklanınca tablo aşağıdaki şekle 
		dönüşüyor,</p>
		<p><img src="../../images/vbadaointeraktif2.gif"></p>
		<p>Kod bloğumuz aşağıda duruyor. Buna göre;</p>
		<ul>
			<li>Önce global değişkenlerimizi(biri sabit) yaratıyoruz</li>
			<li>Sonra Bölge veya Ay bilgileri dğeiştiğinde(Bunlar B1 ve B2 
			hücrelerinde bulunuyor) Change event'ini tetikliyoruz.</li>
			<li>Sonra da çift tıklama eventini handle ediyoruz.</li>
		</ul>
		<p>Her iki event prosedürünü de başına breakpoint koyup F8 ile 
		ilerleyerek kodu incelemenizi tavsiye ederim.</p>
	<pre class="brush:vb">
Dim db As dao.Database
Dim rs As dao.Recordset
Const adres As String = "C:\inetpub\wwwroot\aspnettest\excelefendiana\Ornek_dosyalar\Makrolar"
'--------Bölge ve Ay bilgileri değiştiğinde tetiklenecek prosedür
Private Sub Worksheet_Change(ByVal Target As Range)
'yanlışlıkla başka bir hücreye çift tıklarsa onun içine girmiş olur ve burası tetiklenir
If Target.Row = 4 Then Exit Sub 'başlığa çift tıklanırsa
If IsEmpty(Target) Then Exit Sub 'bir de herhangi boş bir hücreye çift tıklanırsa
    
    Application.EnableEvents = False
    Range("a4").CurrentRegion.Offset(1).Clear 'önce temizlik
    
    If Not Intersect(Target, Range("B1:B2")) Is Nothing Then
        [a5].Select
                        
        Set db = dao.OpenDatabase(adres + "\BölgeŞubeRakamları.accdb")
        
        mySql = "select * from bölgerakam where Bölge='" & Range("bölge") & "' and Ay=" & Range("ayno")
        Set rs = db.OpenRecordset(mySql)
       
        ActiveCell.CopyFromRecordset rs
        
    End If
    Application.EnableEvents = True
    
End Sub
'----Ürün bilgisine çift tıklandığında tetiklenip şube detayını gösterecek olan prosedür
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
If IsEmpty(Target.Offset(1, -1)) And Not IsEmpty(Target.Offset(1, 0)) And Not IsEmpty(Target.Offset(0, -1)) Then
    Application.EnableEvents = False
    Do
        ActiveCell.Offset(1, 0).EntireRow.Delete
    Loop Until Not IsEmpty(ActiveCell.Offset(1, -1))
    Application.EnableEvents = True
    Target.Offset(0, 1).Select
    Exit Sub
End If

If Not Intersect(Target, Range([b5], [b5].End(xlDown))) Is Nothing And Not IsEmpty(Target.Offset(0, -1)) Then
    Application.EnableEvents = False
    Set db = dao.OpenDatabase(adres + "\BölgeŞubeRakamları.accdb")
    
    mySql = "select şube,ay,rakam from şuberakam where Bölge='" & Range("bölge") & "' and Ay=" & Range("ayno") & " and ürün = '" & Target.Value2 & "'"
    Set rs = db.OpenRecordset(mySql)
    'Debug.Print rs.Type
    rs.MoveLast 'bir üst satırdaki ' işaretini kaldırıp F8 ile ilerlersek
    'görürüz ki recordsetin tipi dynaset, o yüzden recordcoutn ele etmek için en sona konumlanmalıyız
    şubeadet = rs.RecordCount
        
    For i = 1 To şubeadet
        Target.Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    rs.MoveFirst
    Target.Offset(1, 0).CopyFromRecordset rs
    Target.Offset(0, 1).Select
    
    Application.EnableEvents = True
End If


End Sub
</pre>
</div>
	<h2 class="baslik">ADO DETAY</h2>
	<div class="konu">
	<h3>Giriş</h3>
	<p>Yukarıda belirttiğim 
	gibi Access dışındaki yeni çalışmalarınızda ADO’yu kullanmanızı 
	öneriyorum(hatta isterseniz Access’te bile ADO’yu kullanabilirsiniz, ancak 
	DAO Accese özgü olduğu için daha hızlıdır ve esnektir. Ben iki şeyi bilmekle 
	uğraşmayayım sadece tek şeyi bileyim diyorsanız ADO size yeter, sadece 
	ADO’yu öğrenin)</p>
	<h4>Referans ekleme</h4>
	<p>ADO’yu çalışmalarınızda 
	kullanabilmek için buna ait Library’nin reference olarak eklenmesi 
	gerekir. İşletim sisteminin versiyonuna göre uygun library seçimi yapılır, 
	genelde en yüksek versiyon seçilir. Windows 7 ve sonrası için 6.1 gibi. 
	Ancak yapacağınız çalışmayı başka kişiler de kullanacaksa ve onlar sizin 
	işletim sisteminden daha aşağı seviyelerde bi işletim sistemi 
	kullanıyorlarsa siz de ya daha düşük versiyonu seçmeli
	<span style="color: rgb(0, 0, 0); font-family: &quot;Trebuchet MS&quot;, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(248, 248, 248); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;">
	veya Late Binding yöntemini kullanmalısınız.</span>.</p>
	<p>NOT:DAO’daki TableDef ve QueryDef nesneleriyle yapılan DB seviyesindeki 
	işlemleri ADO ile yapamıyoruz. Bunun yerine ADOX librarysi eklenmelidir. Ama 
	bu tür işlemler zaten konumuzun dışında olduğu için burada buna hiç 
	girilmeyecektir.</p>
	<h4>Yakından bakış</h4>
	<p>Genel mantık DAO’ya benzer. Bundaki süreç ise şöyledir:</p>
	<ul>
		<li>Connection yaratılır</li>
		<li>Recordset yaratılır </li>
		<li>Kayıtlara erişilip işlem yapılır</li>
		<li>Recordset ve Connection kapatılır</li>
	</ul>
	<p>Gördüğünüz 
	gibi burada <strong>Database</strong> nesnesi yok, onun yerine <strong>Connection</strong> nesnesini 
	kullanıyoruz.</p>
	<p>En başta belirttiğim gibi ADO, MS’un en güncel data erişim 
	teknolojisidir ancak ADO bunu tek başına yapmaz, OLE DB Provider denen bir 
	aracı teknoloji ile yapar. Genelde her data kaynağı için ayrı bir OLE DB 
	sağlayıcısı vardır, ancak MS’taki abiler ODBC bağlantı türü(Genel amaçlı 
	provider) için de OLE DB sağlayıcısı yapmışlar, böylece ADO ile her tür veri 
	kaynağına bağlanılabilmektedir.</p>
	<p>Datayı manipüle etmede kullanılan ve 
	arkaplandaki esas yazılıma<strong> DB Engine(VT motoru)</strong> deniyor. Access, DB Engine 
	olarak <strong>Jet</strong> (Joint Engine Technology) kullanır. 2003 öncesi versiyonlarda bu, 
	Jet 4.0 OLE DB provider iken 2007 sonrasında (.accdb database), 
	"Microsoft.ACE.OLEDB.12.0" oldu, <strong>ACE</strong> (Access Connectivity Engine).</p>
	<p>Bu 
	arada sadece Accese değil, daha önce söylediğimiz gibi bir metin dosyasına 
	hatta bir Excel dosyasına bile ADO ile bağlanabiliriz.</p>
	<p>Daha detaylı bilgiye
	<a href="https://msdn.microsoft.com/en-us/library/office/jj249129.aspx">
	buradan</a> ulaşabilirsiniz.</p>
	<h3>Nesneler</h3>
	<h4>ADO Nesne modeli</h4>
	<p>ADO’daki nesne 
	sayısının daha az olduğunu ama metod sayısının çok olduğunu söylemiştik. Temel nesnemiz
	<strong>Connection</strong>’dır. Bunun Open&nbsp;ve&nbsp;Close&nbsp;metodları vardır.</p>
	<p><strong>Recordset</strong>&nbsp;nesnesi ile DAO'da olduğu gibi dataya erişiriz ve gerektiğinde onu işleriz 
	(Update,Delete..)</p>
	<p><strong>Record</strong>&nbsp;nesnesi, Recordsetteki bir satır kaydı gösterir.</p>
	<p><strong>Fields</strong>&nbsp;Collection’ı tablodaki tüm kolonları gösterirken,&nbsp;<strong>Field</strong>&nbsp;nesnesi, bu 
	kolonlardan herhangi birini ifade eder.</p>
	<p>DAO konusunu anlatırken de 
	bahsetmiştik. Bu objeleri nesneleri tanımlarken başlarına library’sini(DAO için DAO, ADO için ADODB) koymakta fayda var, özellilkle ilgili 
	VBA projesi içinde hem DAO hem ADO refere edildiyse. Bazıları için nesne 
	isimleri ortak olmamakla ve bir karışıklığa neden olmamakla birlikte bu 
	alışkanlık iyi bir alışkanlıktır.(Sadece ADO’cu olmaya karar verdiyseniz 
	gerek yok tabi)</p>
	<p>DAO ile ADO nesneleri arasındaki farkları ve benzerlikleri
	<a href="http://www.databasejournal.com/features/mssql/article.php/1490571/From-DAO-to-ADO.htm">
	şu sitede</a> bulabilirsiniz.
	</p>
	<p>Bu arada önemli bir husus da şu: Nesnelerin isimleri aynı olabilir ama 
	metod ve propertyler farklıdır. Yani aynılar diye kullanım şekilleri de aynı 
	olmak zorunda değil.</p>
	<p>Şimdi bu nesnelere yakından bakalım.</p>
	<h4>Connection 
	nesnesi</h4>
	<h5>Tanımlama ve Yaratım</h5>
	<p>DAO’da Database nesnini yaratırken New 
	keywordunu kullanmıyorduk. Çünkü DAO’da 
	nesnelere Set atamasını yaparken bir fonksiyon(OpenDataBase) ile yaratıyorduk. 
	Ancak ADO’da ise yeni(New) Connection nesnesini yaratıp, bu nesnenin kendi 
	Open metodu ile bağlantıyı açıyoruz.</p>
		<pre class="brush:vb">
'Earlybinding
Dim con As New ADODB.connection 'tek satır
'Veya  iki satırda
Dim con As ADODB.connection
Set con = New ADODB.connection

'Late binding
Dim con As Object
Set con = CreateObject("ADODB.Connection")			</pre>
	<p>Connection nesnesini <span class="keywordler">Open </span>metodu ile başlatırız(açarız). DAO’dan farklı 
	olarak sadece Accese değil başka veritabanlarına da ulaşabildiğimizi 
	söylemiştik. Bunun için <strong>Connection String</strong> denen bir ifadeye ihtiyaç duyarız. 
	Syntaxı aşağıdaki gibidir.</p>
	<p><strong>ConnectionObject.Open ConnectionString, 
	UserID, Password, Options.</strong></p>
	<p>Genelde ilk parametre yani ConnectionString 
	yeterlidir. Bunun da Provider ve Datasource değerlerini(ODBC için Driver ve 
	DBQ değerlerini) girmek yeterlidir.</p>
	<p>Şimdi çeşitli conneciton yaratma 
	alternatiflerine bakalım. Bunlarda Access veritabanına OLEDB ile 
	bağlanacağız.</p>
		<pre class="brush:vb">
Sub adoconyarat()
Dim con As ADODB.connection
Dim strDB As String
 
Set con = New ADODB.connection
strDB = adres + "vbadb.accdb"
 
' ConnectionString propertysini belirterek
With con
    .ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB 'oledb için
' odbc için:
'.ConnectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strDB
    .Mode = adModeRead
    .Open
End With
'....
con.Close

End Sub		</pre>
	<p>veya connection string yerine provider ve datasource ayrı ayrı belirtilerek;</p>
		<pre class="brush:vb">
With con
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Properties("Data Source ") = strDB
    .Mode = adModeRead
    .Open
End With	</pre>
	<p>Aşağıda Access dışındaki diğer data kaynaklarına 
	bağlanırken kullanılan ConnectionString’leri görebilirsiniz.<br>&nbsp;<br>Diğer 
	bağlantı türleri için&nbsp;<a href="http://www.connectionstrings.com">http://www.connectionstrings.com</a> sitesine 
	bakabilirsiniz, ancak biz aşağıda zaten birkaç türünü göreceğiz. Şimdi 
	bunlara bakalım.</p>
	<h5>Excel dosyalarına ulaşmak</h5>
	<p>ADO ile 
	başka Excel dosyalara da ulaşabilip veri alabilmekteyiz. Tabi isterseniz bu 
	işlemi ilgili dosyayı yine makro ile 
	açıp sonra copy paste yaptırıp kapatma veya bir sonraki sayfada göreceğimiz gibi 
	refreshlenebilir bir Connection kurma yoluyla da yapabilrsiniz. ADO'nun 
	farkı, ilgili sayfayı olduğu gibi almak yerine satırların/sütunların 
	sayısını almak veya sadece belli bir hücreyi/satırı/kolonu almak için 
	rahatlıkla kullanılabilmesindedir.</p>
		<p>Bağlantıyı aşağıdaki gibi kuruyoruz. Bağlantı için OLEDB sağlayısını 
		kullanırız. Connection string olarak dosya adresi yazıldıktan sonra 
		Extended Properties özelliğine Excelin versiyonu yazılır ve Xml ifadesi 
		eklenir.</p>
		<p>Recordseti açarken de veriyi hangi sayfadan alacağımızı SQL metni 
		olarak yazarız. Kolon başlığına Where ifadesi ile filtre de koyabiliriz 
		Ör: Where Bölge='Akdeniz'.</p>
		<pre class="brush:vb">
Sub ExcelADO()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
dbAdres = "C:\inetpub\wwwroot\aspnettest\excelefendiana\Ornek_dosyalar\pivotdata.xlsx"
constr = "Provider = Microsoft.ACE.OLEDB.12.0; Data source=" + dbAdres + ";Extended Properties='Excel 12.0';"

cn.ConnectionString = constr
cn.Open

'veya aşağıdaki gibi
'With cn
'    .Provider = "Microsoft.ACE.OLEDB.12.0"
'    .Properties("Data Source") = dbAdres
'    .Properties("Extended Properties") = "Excel 12.0" 'burda ayrıca ' içine yazmaya gerek yok
'    .Open
'End With

rs.Open "Select * FROM [Sheet1$]", cn, adOpenStatic, adLockReadOnly, adCmdText
ActiveCell.CopyFromRecordset rs

rs.Close
cn.Close
    
End Sub
</pre>
		<p>
		Kaynak alanımız belirli bir adres ise onu adresiyle birlikte yazarak 
		"SELECT * FROM [Sheet1$A1:C100]" şeklinde veya kaynağımız bir Table ise 
		"SELECT * FROM [Table1]" şeklinde belirtebiliriz.</p>
	<h5>Text dosyasına 
	ulaşmak</h5>
	<p>Text dosyalarına bağlanırken hem OLEDB hem ODBC kullanabiliriz. Text 
	dosyasının kendisini connection string içinde geçirmeyiz, bunun yerine 
	ilgili dosyanın bulunduğu klasörü yazarız. Dosya adını recordset içinde 
	belirtiriz.</p>
		<p><strong>OLEDB<br></strong>Provider=Microsoft.ACE.OLEDB.12.0;Data 
	Source=c:\txtFilesFolder\;<br>Extended 
	Properties="text;HDR=Yes;FMT=Delimited";<br><br><strong>ODBC</strong><br>Driver={Microsoft Text 
	Driver (*.txt;&nbsp;*.csv)};Dbq=c:\txtFilesFolder\;<br>Extensions=asc,csv,tab,txt;</p>
	<p>Önemli Not:ODBC bağlantılarında ilgili veri kümesinin ilk satırı her 
	zaman başlık varsayılır, ve ikinci satırdan itibaren data okunur.</p>
	<p>Şimdi text örneklerine bakalım:</p>
	<pre class="brush:vb">
Sub ado_txt1_odbc()
'odbc örneği
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection

conn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
    "DBQ=" & adres & ";" & "Extended Properties=""text;FMT=Delimited""" 'ODBC'de her zaman ilk satır başlık kabul edildiği için HDR parametresi belirtmedik

Set rs = New ADODB.Recordset
rs.Open "select * from [hatalog.txt]", conn, adOpenStatic, adLockReadOnly, adCmdText

Range("a1").CopyFromRecordset rs

rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing

End Sub</pre>
		<p>
		Şimdi de OLEDB ile</p>
		<pre class="brush:vb">
Sub ado_txt2_oledb()
'oledb örneği
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection
conn.Provider = "Microsoft.ACE.OLEDB.12.0"
conn.ConnectionString = "Data Source=" & adres & ";" & "Extended Properties=""text;HDR=no;FMT=Delimited;"""
conn.Open

Set rs = New ADODB.Recordset
rs.Open "select * from [hatalog.txt]", conn, adOpenStatic, adLockReadOnly, adCmdText

Range("a1").CopyFromRecordset rs

rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing

End Sub</pre>
<p>Çok kullanılmayan FixedWidth tipli dosyalar ve Delimiter'ların farklı türlerinin nasıl kullanılacağı ile
<a href="https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ms974559(v=msdn.10)">buradan</a> 
bilgi edinebilirsiniz.</p>
	<p>Bu arada, şunu belirtmeden de geçmeyelim: Text dosyalarını okumanın, tüm içeriğini elde etmenin 
	ve içeriğini değiştirmenin başka yolları da var, bunun için de 
	<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx">şuraya</a> 
	bakabilirsiniz. Ancak ordaki yöntemler daha çok küçük datalar için anlamlı. 
	Text dosyasını veritabanı amaçlı kullanmak için(yüzbinlerce satırdan 
	bahsediyorum) yine ADO’yu tercih edin. Linkteki yöntemleri ise daha ziyade 
	kısa metin içeren dosyalarda kullanın.</p>
	<h4>Mode propertysi</h4>
	<p>Yukarıdaki 
	örneklerde farkettiyseniz Mode özelliği Connection henüz açılmamışken 
	atandı, çünkü bu özellik sadece kapalı bağlantılarda atanabilir. Bunun 
	alacağı değerler ve açıklamaları aşağıda verilmiştir.</p>
	<ul>
		<li><strong>adModeUnknown</strong>: 
	Default budur. İzinler henüz set edilmemiştir ve tam karar verilemez. ADO, 
	provider’a kendisi karar verecektir. </li>
		<li><strong>adModeShareDenyNone:</strong> Başkalarının da 
	her türlü yetkiyle açmasına izin verir.</li>
		<li><strong>adModeReadWrite</strong>: Read/Write 
	olarak açar. Yani hem okuma hem yazma yapılabilir.</li>
		<li><strong>adModeShareDenyRead</strong>: 
	Sizde açıkken başkaları buradan okuma yapamaz.</li>
		<li><strong>adModeRead</strong>: Sadece okuma 
	yapabilirsiniz, yazma yapamazsınız. </li>
		<li><strong>adModeShareDenyWrite</strong>: Sizde açıkken 
	başkaları buraya yazma yapamaz.</li>
		<li><strong>adModeWrite </strong>: Sadece yazma yapabilirsiniz, 
	okuma yapamazsınız.</li>
		<li><strong>adModeShareExclusive</strong>: Bağlantı sizde açıkken başkaları 
	o an bağlanamaz.</li>
	</ul>
	<p>Mesela sadece okuma amaçlı açılan aşağıdaki kodda 
	AddNew satırında hata alırsınız.</p>
	<pre class="brush:vb">
Sub adomode()
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Set con = New ADODB.Connection

With con
     .ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & adres & "vbadb.accdb"
    .Mode = adModeRead 'Write modu açılırsa sorun yok
    .Open
End With

Set rs = New ADODB.Recordset
rs.Open "data", con, adOpenDynamic, adLockOptimistic
rs.AddNew 'burada hata
rs.Fields(0)="x bölgesi"
rs.Fields(1)="y şubesi"
rs.Fields(2)="z ürünü"
rs.Fields(3)=100
rs.Update

End Sub</pre>
	<h3>Recordset nesnesi</h3>
	<p>DAO’da 
	olduğu gibi ADO’da da veritabanına bir kez eriştiken sonra artık kayıtlarda 
	istediğimiz işlemleri yapabiliyoruz.</p>
	<h4>Tanımlama ve Yaratma</h4>
	<p>DAO’da Recordset nesnesini New keywordu olmadan yaratıyorduk. Set 
	atamasını yaparken de başka bir nesne olan Database nesnesinin OpenRecordSet 
	metodunu kullanıyorduk. ADO’da ise New keywordu kullanılarak(Dim satırında 
	veya Set satırında) yeni bir RecordSet nesnesi yaratılır ve bu nesnenin 
	kendi metodu olan Open Metodu kullanılır(İstisna:Connection veya command 
	nesnelerinin Execute metodu ile recordset elde edeceksek New ifadesini 
	kullanmayız). Aşağıda genel syntaxı bulunmaktadır.</p>
	<p><strong>Recordset.Open Source, ActiveConnection, CursorType, LockType, Options</strong></p>
	<pre class="brush:vb">Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" &amp; strDB
rs.Open "Select * from Subeler", con</pre>
	<p><strong>Source,</strong> bir SQL metni olabileceği gibi, tablo/sorgu adı da olabilir, bir 
	Command nesnesi de. Genel olarak çekilecek kayıt miktarını minimize etmek 
	iyi bir alışkanlıktır. O yüzden mümkünse tam bir tablo yerine bir SQL metni 
	veya amaca hizmet eden bir sorgu(query) seçilmelidir. Hatta sadece test 
	datası görmek istiyorsanız çekilen kayıt sayısını sınırlayabilirsiniz.</p>
	<pre class="brush:vb">
Sub adorecordset()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB

strSQL = "SELECT * from [Tarihsel Data] where Ürün='Ürün1' and Ay=#1/31/2016#"
rs.Open Source:=strSQL, ActiveConnection:=con, CursorType:=adOpenDynamic, LockType:=adLockOptimistic

ActiveCell.CopyFromRecordset rs, 10 'çekilen kayıt sayısını 10 ile sınırladık

End Sub	</pre>
	<p><strong>CursorType</strong>, arama yönü ve görüntüleme tipini ifade eder. 
	Alabileceği değerler şunlardır:</p>
	<ul>
		<li><strong>adOpenForwardOnly</strong>: Default budur. Sadece ileri hareket eder. Aksi 
	gerekmedikçe bunu kullanırsanız daha hızlı erişim sağlarsınız.</li>
		<li><strong>adOpenStatic</strong>: Tüm yönlere izin vardır ve başkaları tarafından yapılan 
	değişiklikler o an size görünmez. Yani bir nevi siz eriştiğiniz anda ilgili 
	kayıt setinin resmi çekilir ve siz hep onu görürsünüz.</li>
		<li><strong>adOpenDynamic</strong>: Bu 
	da tüm yönlere izin verir ama bu sefer başkaları tarafından yapılan 
	değişiklikler size anında görünür.</li>
		<li><strong>adOpenKeyset</strong>: adOpenDynamic’e benzer, 
	ama silinen veya eklenenler size o an görünmez, sadece değişiklikleri 
	görebilirsiniz. </li>
	</ul>
	<p><strong>LockType</strong>, kayıtlar güncellenirken ne tür kilit 
	konacağını ifade eder. Çok kullanıcının eriştiği bir dosyada aynı anda 
	birden çok kullanıcı dataya erişmeye veya değiştirmeye çalışırsa nasıl 
	davranılması gerektiğini belirler.</p>
	<ul>
		<li><strong>adLockReadOnly</strong> :Default budur. 
	Kayıtlar herkeste readonly açılır ve kimse editleyemez.</li>
		<li><strong>adLockOptimistic</strong>: 
	Update sırasında(Update metodu çağrıldığında) kilitler. Başkaları da o 
	sırada görebilir ve editleyebilir.</li>
		<li><strong>adLockPessimistic</strong>: Editlemeye 
	başladığınız anda kilitler. Başkaları o sırada bu kaydı okuyamaz ve 
	editleyemez.</li>
		<li><strong>adLockBatchOptimistic</strong>: adLockOptimistic’in aynısı, sadece 
	toplu güncelleme yapıldığında kullanılır.</li>
	</ul>
	<p>Şimdi, yukarda Mode 
	konusunda verdiğimiz örneğe bakalım. Oradaki Mode’u Read yerine Write 
	yapalım, ki veritabanına yazma izni vermiş olalım. Ancak bu sefer de lock 
	tipini adLockReadOnly yapalım. Böyle bir durumda yine yazma izni verilmemiş 
	olur. Mesela çok kullanıcılı bir dosyada, sadece belli kullanıcıların yazma 
	izni olsun istiyorsanız Mode’u Write yaparsınız, ama kullanıcı grubuna göre 
	kimini ReadOnly lock tipinde kimini Diğer lock tiplerinde açarsınız.</p>
	<pre class="brush:vb">
Sub adomode2()
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Set con = New ADODB.Connection

With con
    .ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & adres & "vbadb.accdb"
    .Mode = adModeWrite
    .Open
End With

Set rs = New ADODB.Recordset
rs.Open "data", con, adOpenDynamic, adLockReadOnly
rs.AddNew 'burada hata
rs.Fields(0)="x1 bölgesi"
rs.Fields(1)="y1 şubesi"
rs.Fields(2)="z1 ürünü"
rs.Fields(3)=100
rs.Update

End Sub</pre>
	<p><strong>Options</strong> 
	parametresini genelde boş bırakıyoruz, böyle olunca ADO, bunun ne olduğuna 
	kendi karar vermeye çalışıyor. Ben şimdiye kadar belli durumlar haricinde 
	bunu kullanmadım. Sadece aşağıda belirttiğim gibi Seek metodu kullanılırken 
	bunun özel bir değer olarak girilmesi lazım, onun dışında tespiti ADO’ya 
	bırakıyorum.</p>
	<h4>Field nesnesi(Alanlar)</h4>
	<p>DAO’da söylediğim gibi, 
	bizim Field’larla işimiz daha çok bunlara erişmek şeklinde olacak. Erişimin de 3 yolu bulunmaktadır.</p>
	<p>Alan adı ile:rs.Fields("İsim")<br>Alan 
	adı kısayolu ile:rs![İsim]<br>Alan item no ile:rs.Fields(1)</p>
	<p><strong>Recordset.Fields(0).Name</strong>: İlk kolonun adını yani kolon başlığını getirir.<br>
	<strong>Recordset.Fields(0).Value</strong>: Bu ise ilk kolondaki geçerli kaydın(satırın) 
	içeriğini döndürür. Bu arada Value özelliği default özellik olup yazılmasa 
	da olur, ama biz iyi bir programlamacı olup yazıyoruz.<br><strong>rs.Fields.Count</strong>: 
	İlgili kayıt setindeki alan(kolon) sayısını verir.</p>
	<h4>Yeni kayıt, mevcut düzenleme ve kaydetme</h4>
	<p>Burada bir çok şey DAO’ya benzer, o yüzden benzer olanları sadece 
	belirteceğim, bunların detaylarına girmektense farklılık gösterenlere 
	değinmeyi tercih edeceğim.</p>
	<p><span class="keywordler">AddNew</span> ve <span class="keywordler">Update</span> metodları aynen DAO’daki gibi 
	geçerlidir.<br>DAO’dan farklı olarak ADO'da, Update işleminden sonra aktif kayıt 
	yeni kayıt olur, DAO’da ise yeni kayda çapa atmak gerekiyordu.</p>
	<p>ADO kayıt 
	seti her zaman edit modunda olduğu için ayrıca bir Edit metoduna ihtiyaç 
	duyulmamaktadır.</p>
	<p>Update metodu var olmakla birlikte MoveNext gibi 
	cursorın konumunu değiştiren işlemlerde otomatik Update kolaylığı gelmiştir. 
	Bununla birlikte değişikliklerden sonra cursor konumunu değiştiren bir 
	metodu uygulanmaycaksa Update yine de yapılmalıdır.</p>
	<h4>Silme</h4>
		<p>DAO konusunda bahsettiğim gibi, silme işlemi için ben SQL metni 
		çalıştırmayı tercih ediyorum. Üstelik ADO'da silme işlemi biraz daha 
		detylı olabiliyor. O yüzden ileri bir vadede bu kısımda güncelleme 
		yapana kadar SQL çalıştırmak ile devam edebilirsiniz.</p>
	<h4>Kayıtlarda dolaşma</h4>
	<p>DAO’daki Move ve türevleri aynen var. Oraya bakabilirsiniz.</p>
	<h4>RecordCount 
	ile kayıt sayısı</h4>
	<p>Recordsetteki kayıt sayısı için DAO’daki 
	gibi <span class="keywordler">RecordCount</span> özelliği kullanılır. Ancak DAO’da olduğu gibi burda da dikkat 
	edilmesi gereken bazı hususlar vardır. Cursor tipi ve provider türüne göre 
	sonuçlar farklılık gösterir.</p>
	<ul>
		<li>Cursor tipi adOpenForwardOnly 
	ise -1 döner.(Connection nesnesnin Execute metodu recordseti bu tipte açar)</li>
		<li>Cursor tipi adOpenStatic or adOpenKeyset ise sorunsuz çalışır</li>
		<li>Cursor tipi adOpenDynamic ise data kaynağına göre sonuç değişir<ul>
			<li>Destekliyorsa sorunsuz</li>
			<li>Desteklemiyorsa -1</li>
		</ul>
		</li>
	</ul>
	<p>Kayıt sayısını elde 
	etmenin bir yolu da <span class="keywordler">GetRows</span> ve Ubound 
	birleşimi olacaktır ve bu yol her zaman garantidir.</p>
	<pre class="brush:vb">
Sub adokayıtsayısı()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB

strSQL = "SELECT * from [Tarihsel Data] where Ürün='Ürün1' and Ay=#1/31/2016#"
rs.Open Source:=strSQL, ActiveConnection:=con, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
dizi = rs.GetRows

Debug.Print rs.RecordCount 'cursortype durumuna göre değişkenlik gösterir
Debug.Print UBound(dizi, 2) + 1 'her zaman garantidir

End Sub	</pre>
	<p>Bu yukardaki kodun DAO karşılığı aşağıdaki gibi olacaktır.</p>
	<pre class="brush:vb">
Sub dao2()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = DAO.OpenDatabase(adres + "vbadb.accdb")
Set rs = db.OpenRecordset("SELECT * from [Tarihsel Data] where Ürün='Ürün1' and Ay=#1/31/2016#", dbOpenDynaset)

rs.MoveLast
x = rs.RecordCount
Debug.Print x

rs.MoveFirst 'Getrows demek için tekrar başa geliyoruz
dizi = rs.GetRows(x)
Debug.Print UBound(dizi, 2) + 1

End Sub	</pre>
	<h4>Seek,Find, Filter ve SQL ile kayıt arama</h4>
	<p>DAO’da olduğu gibi 
	ADO’da da aradığımız bilgiyi bulmanın birkaç yolu bulunmaktadır. 
	Okumadıysanız DAO kısmındaki giriş notlarını okumanızı tavsiye ederim.</p>
	<h5>Seek</h5>
	<p>DAO’da bahsettiğimiz Find ve Seek arasındaki farklar ve bunların 
	özellikleri büyük ölçüde geçerlidir. Mesela en hızlısı yine Seek metodudur. 
	Ancak, Seek metodunun kullanılabilmesi için RecordSetin <strong>Options</strong> parametresi
	<strong>adCmdTableDirect</strong> olarak belirtilmelidir.</p>
	<p>Ayrıca açılan recordsetin index 
	özelliğini ve seek metodunu destekleyip desteklenmediğini 
	<span class="keywordler">Supports</span> metodu ile 
	kontrol etmemiz gerekir.</p>
	<pre class="brush:vb">
Sub adoseek()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB

rs.Open "data", con, adOpenKeyset, adLockOptimistic, adCmdTableDirect 'Seek'in kullanılması için adCmdTableDirect olmalı

If rs.Supports(adIndex) And rs.Supports(adSeek) Then
    rs.Index = "Şube Adı"
    rs.Seek ("Şube1")
End If

'nomatch yok. filter ve recordcount yapılabilir. veya BOF/EOF kontrolü
If (rs.BOF = True) Or (rs.EOF = True) Then
    Debug.Print "Data Not Found"
Else
    Debug.Print rs.Fields("Bölge").Value
End If
End Sub	</pre>
	<h5>Find</h5>
	<p>Kayıt aramadaki diğer alternatiflerimiz arasında Find 
	var. Özellikle küçük bir recordset üzerinde çalışırken kullanılabilir. Büyük 
	recordsetlerde çok hantal olacaktır.</p>
	<p>DAO’da Find yerine Find’ın 
	türevleri bulunmakta idi, ADO’da ise sadece Find bulunmaktadır. Syntax 
	aşağıdaki gibidir.<br><strong>Find (Criteria, SkipRows, SearchDirection, Start)</strong></p>
	<p>MSDN'de Find'ı kullanmadan önce ilk kayda konumlanılması önerilyor ancak 
	ben bunu yapmadığımda da kod çalışıyor, yine de tavsiyeye uyuyorum.</p>
	<p>Eğer arama yapacağımız küme büyük bir veri kümesiyse öncesinde
	<span class="keywordler">Sort</span> metodu ile recordseti sıralamakta fayda 
	var. Bunun için recordseti açamadan önce <strong>CursorLocation</strong> 
	özelliğine <strong>adUseClient</strong>&nbsp; değerini atamak gerekiyor.</p>
	<p>Önemli bir detay; Metinsel alanlara filtre uygulanırken kriter tek tırnak arasına alınır. 
	Aşağıdaki kodda bir örneği var(Şube112 filtresi).</p>
	<pre class="brush:vb">
Sub adofind()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB

'rs.CursorLocation = adUseClient 'sort yapılacaksa burasını açın
rs.Open "data", con, adOpenKeyset, adLockOptimistic
rs.MoveFirst 'Find kullanmadan önce ilk kayda konumlanılmalıdır

'rs.Sort = " [Şube Adı]" 'Find’ı kullanmadan önce sıralama yaparak arama hızı artırılabilir.
rs.Find "[Şube Adı] = 'Şube112'" 

'nomatch yok. filter ve recordcount yapılabilir. veya BOF/EOF kontrolü
If (rs.BOF = True) Or (rs.EOF = True) Then
    Debug.Print "Data Not Found"
Else
    Debug.Print rs.Fields("Bölge").Value
End If

End Sub</pre>
	<h6>Diğer hususlar</h6>
	<ul>
		<li>NULL 
	yazarken “is null” değil sadece “Null” yazarız.</li>
		<li>DAO’da Find 
	türevleri içinde “AND” kullanılabilirken ADO’da kullanılamıyor. Yani sadece 
		tek bir kolon için arama yapılabilirsiniz.</li>
		<li>Joker eleman desteklenir ve Like kelimesi ile kullanılır.<br>rs.Find = 
		"ŞubeAdı like '*AKSARAY*' "<br></li>
	</ul>
	<p>
	<a href="https://www.techrepublic.com/article/why-ados-find-method-is-the-devil/">
	Şu sitede</a> Find metodunun kullanılmaması gerektiği ile ilgili tavsiyeler 
	var, kararı siz verin.<br></p>
	<h5>Filter</h5>
	<p>Bir diğer yöntem Filter’dır. Filter’ın SQL’deki Where kısmına benzeyen 
	bir kullanım şekli vardır. Hani demiştik ya, ADO kullanırken SQL bilmeye 
	gerek yok, işte bu yöntem ile bir nevi SQL ile filtre uygulamış oluyoruz.
	<strong>Syntax’ı şöyledir: rs.Filter = Kriterler</strong></p>
	<p>Özellikle belli kriterlere göre <strong>birden çok kayıt getirmek istediğinizde</strong> 
	bunu kullanırız. Tek kayıt ararken daha çok Seek(destekleniyorsa) veya Find 
	kulanılımalı.</p>
	<p>Önemli bir nokta var, o da <strong>Recordseti açmadan önce</strong> filtreyi uygulamamız 
	gerektiği. Filtreyi kaldırmak istediğimizde ya "" atarız veya 
	adFilterNone uygularız. Bir diğer nokta da, DAO'da olduğundan farklı olarak 
	burda ayrı bir recordset tanımlamamıza gerek yok, aynı recordset üzerinde 
	filtre uygulanmaktadır.</p>
	<pre class="brush:vb">
Sub adofilter()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB

rs.Filter = "[Bölge] = 'Başkent 1' and [Ürün Adı]='Ürün1' and [Aylık Gerç] > 40000"
rs.Open "data", con, adOpenKeyset, adLockOptimistic

'nomatch yok. filter ve recordcount yapılabilir. veya BOF/EOF kontrolü
If (rs.BOF = True) Or (rs.EOF = True) Then
    Debug.Print "Data Not Found"
Else
    ActiveCell.CopyFromRecordset rs
End If

End Sub</pre>
	<h6>Diğer hususlar</h6>
	<ul>
		<li>Filter, DB2’da 
	desteklenmiyor.</li>
		<li>AND/OR keywordunu Find içinde kullanamıyoruz, 
	Filterda ise kullanabiliyoruz.</li>
		<li>Joker eleman desteklenir ve Like 
		ifadesi ile kullanılır.<br>rs.Filter =”ŞubeAdı like ‘*AKSARAY*’ ”<br></li>
	</ul>
	<h5>SQL</h5>
	<p>Aranan değeri bulmada son yöntem ise SQL kodu çalıştırmaktır. Aradığımız 
	değer tek ise tek sonuç döndürecek bir SQL yazmamız gerekir. Eğer yazdığımız 
	SQL’in çoklu sonuç döndürme ihtimali varsa MoveFirst diyerek ilk kaydın 
	sonucunu alırız, tabi SQL’de uygun OrderBy işlemini yapmış olmamız kaydıyla.</p>
	<p>Aradığımız şey birden çok değer ise çoklu sonuç döndüren bir SQL hazırlarız. 
	Aşağıda SQL çalıştırma(Hem Select hem de eylem sorguları) detayları 
	bulunmaktadır.</p>
	<h3>Execute ile SQL çalıştırma</h3>
	<p>DAO’da hem Database hem de QueryDef nesnesi için Execute metodu vardı, ADO’da da hem Connection 
	için hem de Command için var. İkisi de recordset döndürüyor. Syntaxları 
	aşağıdaki gibidir:</p>
	<pre class="brush:vb">Set rs1 = conn.Execute (CommandText, RecordsAffected, Options)
Set rs2 = cmd.Execute( RecordsAffected, Parameters, Options )</pre>
	<p>Tabi çalıştırılacak SQL, Update/Insert gibi değişiklik yapan bir SQL ise bir 
	recordset nesnesine atanmadan doğrudan kullanılabilir.</p>
	<p>Genel olarak DAO’daki mantıkla aynıdır. Oraya 
	bakabilirsiniz. Aşağıda bir örneğimiz de var zaten.</p>
	<pre class="brush:vb">
Sub adosql()
Dim con As New ADODB.Connection
Dim rs As ADODB.Recordset 'New ifadesini kullanmıyoruz

strDB = adres & "vbadb.accdb"
con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
cmdstr = "Select * from data where [Bölge] = 'Başkent 1' and [Ürün Adı]='Ürün1' and [Aylık Gerç] > 40000"
Set rs = con.Execute(cmdstr)

'nomatch yok. filter ve recordcount yapılabilir. veya BOF/EOF kontrolü
If (rs.BOF = True) Or (rs.EOF = True) Then
    Debug.Print "Data Not Found"
Else
    ActiveCell.CopyFromRecordset rs
End If

End Sub</pre>
		<h4>
		Command Nesnesi</h4>
		<p>
		Yukarda gördüğümüz üzere, command nesnesini SQL metinleri çalıştırmak 
		için kullanabiliyoruz. Aşağıda daha detaylı bir örnek var</p>
		<pre class="brush:vb">
Sub adocommand()
    Dim Conn As New ADODB.connection
    Dim Cmd As New ADODB.Command
    Dim Rs As New ADODB.Recordset

    Cmd.CommandText = "SELECT * from data"
    Cmd.CommandType = adCmdText
    
    strDB = adres & "vbadb.accdb"
    Conn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    Cmd.ActiveConnection = Conn
    
    Set Rs = Cmd.Execute
    ActiveCell.CopyFromRecordset Rs
End Sub</pre>
		<p>
		Her ne kadar connection nesnesi ile de SQLçalıştırabilsek de, command 
		nesnesinin bazı avantajları vardır. Bu avantajları sayesinde baen tercih 
		edilebilirler.</p>
		<ul>
			<li>Command nesnesi daha hızlıdır</li>
			<li>Connection'ın aksine parametre kabul edebilir. Ama biz buna 
			burada girmeyeceğiz.</li>
			<li>Conection'da olmayan bazı özellikler Commandda bulunur, bunların 
			da detayına girmeyeceğiz.</li>
		</ul>
		<p>Bütün bu avantajları nedeniyle command nesnesi bazı durumlarda tercih 
		sebebi olabilmektedir. Bu nesnenin detaylı araştırmasını size 
		bırakıyorum.</p>
		<p>
		NOT:SQL metni çalıştırmanın bir diğer yolu da, recordsetin içinde bunu 
		belirtmektir, ki bunu daha yukarıda zaten görmüştük. Bu yöntemde tabiki 
		sadece Select sorguları çalıştırılır.</p>
	<h3>Datayı Excel’e almak(Import işlemi)</h3>
	<p>DAO’daki 
	yöntemlerle aynı olduğu için tekrar yapmak istemiyorum. Oraya 
	bakabilirsiniz.</p>
	<h3>Exceldeki datayı Accese atma</h3>
	<p>Burda kullanılacak yöntemler de DAO’daki gibidir. Oraya bakabilirsiniz.</p>
	<h3 id="passsecurity">Şifre güvenliği</h3>
	<p>Bu maddeyi sadece ADO’ya koydum. Zira 
	DAO ile sadece Access’e bağlanacağımız için ve onda da çoğunlukla şifresiz 
	dosyalara bağlanacağımız için böyle bir sürece gerek bulunmamaktadır. Ancak 
	olur da şifreli bir Access dosyanız varsa, burdaki yöntemleri DAO’da 
	uygulayabilirsiniz.</p>
	<p>Neden bahsediyorum, tabiki connection string içine yazdığınız bağlantı 
	şifresinden. Bunu kodlarınız içine ulu orta yazarsanız bir güvenlik sorunu 
	yaratmış olabilirsiniz. Bu güvenliği artırmak bizim elimizde. Aşağıda benim 
	uyguladığım çeşitli yöntemler bulunuyor.</p>
	<ul>
		<li><strong>VBA project’e protection konması:</strong> İlgili dosyanın VBA koduna birden çok kişinin erişmesi gerekiyorsa(tek 
		geliştirici değilseniz) bu 
	yöntem kullanılmamalıdır. Detaylara 
		<a href="Ileriseviyekonular_Add-InlerveCustomMenuler.aspx#protection">buradan</a> ulaşabilirsiniz.</li>
		<li><strong>XlVeryHidden modundaki bir sayfadan şifrenin okunması:</strong>
		Detaylara 
		<a href="DortTemelNesne_Worksheet.aspx#sayfagizleme">buradan</a> 
		ulaşabilirsiniz.</li>
		<li><strong>Alakasız bir klasöre koyacağınız bir text 
	dosyadan okunması</strong>. (Ör:“C:\yemek tarifleri\sebzeliler\brokoli ziyafeti.txt”) 
	Üstelik bu dosyada gerçek bir yemek tarifi olmasında fayda var. Sadece 
	aralarda bir yerlerde şifrenizi gizleyebilirsiniz.<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx"> I/O işlemleriyle</a> de 
	ilgili karakterleri okutabilrsiniz.</li>
	</ul>
	<p>Bu 3 maddeyi bir arada bile 
	kullanarak ekstra güvenlik sağlayabilirsiniz.</p>
	<p>Bunların dışında bir yöntem 
	daha var ki bu en güvenlisidir: Şifrenin ve hatta User’ın kullanıcıya 
	(her defasında) sordurulması. Bunun dezavantajı ise Schedule edilmiş kodlarda 
	kullanılamaması. Manuel çalıştırılan kodlarda kullanımı uygundur.</p>
		<h3>Örnek Çalışma - İnteraktif Veri Çekme Formatı</h3>
		<p>DAO'da yaptığımız örneğin ADO versiyonu da aşağıdaki gibidir.</p>
		<pre class="brush:vb">
Dim rs As New adodb.Recordset
Dim con As New adodb.Connection
'--------Bölge ve Ay bilgileri değiştiğinde tetiklenecek prosedür
Private Sub Worksheet_Change(ByVal Target As Range)
'yanlışlıkla başka bir hücreye çift tıklarsa onun içine girmiş olur ve burası tetiklenir
If Target.Row = 4 Then Exit Sub 'başlığa çift tıklanırsa
If IsEmpty(Target) Then Exit Sub 'bir de herhangi boş bir hücreye çift tıklanırsa
    On Error GoTo hata
    Application.EnableEvents = False
    Range("a4").CurrentRegion.Offset(1).Clear 'önce temizlik
    
    If Not Intersect(Target, Range("B1:B2")) Is Nothing Then
        [a5].Select
                        
        strDB = adres + "\BölgeŞubeRakamları.accdb"
        con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
        
        mysql = "select * from bölgerakam where Bölge='" & Range("bölge") & "' and Ay=" & Range("ayno")
        rs.Open mysql, con, adOpenForwardOnly, adLockOptimistic
       
        ActiveCell.CopyFromRecordset rs
        rs.Close
        con.Close
        
    End If
hata:
    Application.EnableEvents = True
    
End Sub
'----Ürün bilgisine çift tıklandığında tetiklenip şube detayını gösterecek olan prosedür
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
If Target.Column <> 2 Then 'B kolonu dışında çift tıklarnırsa çıksın
    Cancel = True
    Exit Sub
End If

If IsEmpty(Target.Offset(1, -1)) And Not IsEmpty(Target.Offset(1, 0)) And Not IsEmpty(Target.Offset(0, -1)) Then
    Application.EnableEvents = False
    Do
        ActiveCell.Offset(1, 0).EntireRow.Delete
    Loop Until Not IsEmpty(ActiveCell.Offset(1, -1))
    Application.EnableEvents = True
    Target.Offset(0, 1).Select
    Exit Sub
End If

If Not Intersect(Target, Range([b5], [b5].End(xlDown))) Is Nothing And Not IsEmpty(Target.Offset(0, -1)) Then
    Application.EnableEvents = False
    
    strDB = adres + "\BölgeŞubeRakamları.accdb"
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    
    mysql = "select şube,ay,rakam from şuberakam where Bölge='" & Range("bölge") & "' and Ay=" & Range("ayno") & " and ürün = '" & Target.Value2 & "'"
    rs.Open mysql, con, adOpenStatic, adLockOptimistic
    şubeadet = rs.RecordCount 'adOpenStatic olduğu içn sorunsuz çalışır
        
    For i = 1 To şubeadet
        Target.Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    rs.MoveFirst
    Target.Offset(1, 0).CopyFromRecordset rs
    Target.Offset(0, 1).Select
    rs.Close
    con.Close
    Application.EnableEvents = True
End If


End Sub
	</pre>
</div>
</asp:Content>
