<%@ Page Title='InsertMenusu PivotTable' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Insert Menüsü'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>


<h1>Pivot Table(Özet Tablo)</h1>
<h2 class="baslik">Genel Bakış</h2>
<div class="konu">
	<p>Elimizde liste olarak bulunan bir data kümesini çeşitli seviyelerde özet hale getirmek için 
	PivotTable aracını kullanırız.</p>
	<p>Örneğin, elimizde bir bankanın Bölge-Şube-Ürün (veya bir market zinciri 
	için Bölge-Mağaza-Reyon) detayında rakamları var ve biz bunu Bölge-Ürün 
	bazında görmek istiyoruz diyelim. Datamız şöyle olsun(Tüm liste aşağı doğru 
	uzuyor. İlgili excel dosyasını <a href="../../../Ornek_dosyalar/pivotdata.xlsx">buradan</a> 
	indirebilirsiniz)</p>

	<p>	
	<img alt="" src="/images/insert_pivot0.jpg"></p>

	<p>Şimdi ilgili data kümesi içinde herhangi bir hücredeyken <strong>Insert</strong> menüsünden 
	<strong>PivotTable</strong> butonuna basalım. Burada önemli bir nokta var, o 
	da ilgili kümenin 
	başlıklarında hiç boş hücre olmaması lazım, aksi halde hata alırsınız. (Bu 
	arada bazı arkadaşlarımın tüm data kümesini seçtikten sonra Pivot düğmesine 
	bastığını görüyorum, buna gerek yok, küme içinde herhangi bir hücrede 
	olmanız yeterlidir.) </p>
	<h3><a name="datasource"></a>Data Kaynağı</h3>

	<p>Akabinde aşağıdaki dialog kutusu çıkacaktır. Burada <strong>Select a table or 
	range </strong>kutusuna otomatikman ilgili data kümesinin alanı gelecektir. 
	Excel sayfanızda bir <strong>ListeliTablo(Table)</strong> varsa buraya bunun adı çıkacaktır, ki Özet Tablonuzun 
	kaynak hücre grubunun dinamik olması açısından Table olarak almanız çok daha iyi 
	olur. Böylece kaynak dataya yeni data eklendikçe <strong>Change Data Source
	</strong>yapmak zorunda kalmazsınız. Aşağıdaki örnekte Table7 isminde bir 
	tablodan veri sağlanacağını görüyoruz.(Hemen bir alttaki 
	<strong>external 
	data source </strong>kısmını şimdilik görmezden gelin, bununla ilgili bir örneği 
	ayrıca yapacağız.)</p>
	<p>Özet tablomuzu nereye koyacağımızı bir alttaki seçeneklerde belirtiyoruz. 
	Ayrı bir sayfada görmek istiyorsak <strong>New Worksheet</strong>, mevcut sayfa üzerinde bir 
	yerde istiyorsak <strong>Existing Worksheet </strong>diyip tam olarak başlangıç konumunu 
	belirtiyoruz. Biz aşağıdaki örnek için New Workseheet deyip ilerleyelim.</p>
	<p><img alt="" src="/images/insert_pivot1.jpg"></p>
	<h3>Özet tabloyu düzenleme</h3>
	<p>OK deyip devam ettiğimizde yeni açılan sayfada karşımıza şöyle bir görüntü çıkacaktır.</p>
	<p>
	<img alt="" src="/images/insert_pivot2.jpg" height="60%" width="60%" class="zoomla"></p>
	<p>Sağ taraftaki <strong>PivotTable Fields </strong>panelinde <strong>ROWS</strong> bölümüne Bölge'yi, 
	<strong>COLUMNS</strong> Bölümüne Ürün Adını, ve 
	<strong>VALUES</strong> bölümüne de Aylık Gerç alanını sürükleyelim. Sonuç 
	şöyle olacaktır:<br>&nbsp;
	<img alt="" src="/images/insert_pivot3.jpg"><br>Eğer en sağdaki Grand Total 
	kolonunu gereksiz buluyorsanız (çünkü bazı durumlarda ürünlerin toplamı 
	anlamsız olabilir) bu başlığa sağ tıklanıp <strong>Remove GrandTotal
	</strong>denilerek, 
	otomatik gelen bu kolon silinebilir, ancak en alttaki GrandTotal bu örnek 
	için kalmalıdır, zira bu bize banka toplamını verecektir, ve anlamlı bir 
	bilgidir, tabiki isterseniz yine de sağ tıklayıp aynı kaldırma işlemini 
	yapabilirsiniz. Sonrasında istediğiniz bölgeye istediğiniz bilgiyi 
	sürükleyerek oynayabilirsiniz. (Not:Diptoplamları menülerden de kaldırıp 
	ekleyebiliyorsunuz)<br>
	</div>

<h2 class="baslik">Detaylar</h2>
<div class="konu">


<h3>Pratik hususlar</h3>
	<ul>
		<li>Özet Tablonuzu yeni data geldikçe güncellemek için tablo üzerinde 
		herhangi bir yerdeyken sağ tıklayıp <strong>Refresh</strong> butonuna 
		basmanız yeterlidir. (Ancak, Tablonuzun kaynak datası bir Table değil de 
		bir hücre alanı ise DataSource kısmını da güncellemeniz gerekir. Siz en 
		iyisi datanızı hep <strong>Table</strong> şeklinde muhafaza edin)</li>
		<li>Özet tablonuzun kaynağı bir Table olsa bile, Table manuel veya makro 
		ile refresh olduğunda Özet tablonuz da eş zamanlı olarak refresh
		<span style="text-decoration: underline"><strong>olmaz</strong></span>. 
		Bunun için sizin özet tablonuzu ayrıca refresh etmeniz gerekir.</li>
		<li>Sum of Rakam, Row Label gibi başlıkları değiştirmek çok kolay, sadece 
	bu başlığa gelin ve yeni isim yazın, veya bunu <strong>sağ tık</strong>&gt;<strong>Value Fields Settings'te</strong> de 
	yapabilirsiniz, en üstteki <strong>Custom Name </strong>yazan yere yeni 
		başlığınızı yazın.</li>
		<li>Özet Tabloları hızlı bir 
	şekilde Benzersiz(Uniqe) değerleri almak için de kullanabilrisiniz.(Bir 
		kolonu tümden seçip başka bir yere kopyalayıp RemoveDuplicates yapmak 
		yerine)</li>
		<li>Güzel görünümlü Özet Tablolar için bir öneri: Solda boş bir kolon olsun, gridleri kaldırın, PivotTabloların 
	kendi gridi var zaten. Ayrıca PivotTable Options'ta aşağıdaki işaretli 
		kısmın işaretini kaldırın. Böylece kolonlarınızın genişliği sürekli 
		genişleyip durmaz.
		
		<img alt="" src="/images/insert_pivotcolumnwidth.jpg"></li>
		<li>
		Başlıklar:Bir Özet Tablo içindeyken <strong>Design&gt;Layout</strong>'ta
		<strong>Compact </strong>yerine <strong>Outline</strong> veya <strong>Tabular
		</strong>yaparsanız başlıkların manalı hale geldiğini görürsünüz.(Default 
		format neden bunlardan biri değildir bilmiyorum)</li>
	</ul>

	<h3>Ertelenmiş Güncelleme(Defer Layout Update)</h3>
<p>Özet Tablonuzda değişiklik yapmak istediğinizde her değişklik sonrasında tablonuz otomatik güncellenir, ancak bazen 
tablonuzda birden çok değişiklik yapmak isteyeceksiniz ve eğer bu güncellemeler 
de çok vakit alıyorsa 
<strong>PivotTable Fields</strong> panelinin en altında bulunan <strong>Defer Layout Update</strong> seçeneğini işaretlemeniz 
yeterlidir, böylece değişikleriniz sonrasında hemen güncelleme olmaz, tüm 
değişiklik işleriniz 
bitince bu seçeneğin 
hemen sağındaki <strong>UPDATE </strong>düğmesine tıkladığınızda güncelleme 
gerçekleşir. Çok göze çarpmayan bir özelliktir ama yeri geldiğinde çok faydalıdır.</p>
	<p><img src="/images/insert_pivotdeferupdate.jpg"></p>

<h3>Rakamsal İçerik</h3>
<p>Values bölümüne sürüklediğiniz alana bakarak Excel otomatik olarak rakam 
içeriğini tespit etmeye çalışır. Rakamsal alanlar için varsayılan değer <strong>Toplam
</strong>aldırmakken, Metinsel alanlar için <strong>Adet </strong>saydırmak olacaktır.</p>
	<p>Otomatik gelen bu bilgiyi değiştirmek için PivotTable Fields'tan ilgili 
	alana tıklayıp Value Field Settings kutusunu açarak veya doğrudan ilgili 
	kolondaki bir hücreye sağ tıklayıp <strong>Summarize Values By</strong> 
	diyerek istediğimiz değişikliği yapabiliriz.</p>
	<p>
	<img alt="" src="/images/insert_pivot_field.jpg"></p>

<p>Bu arada <strong>aynı alanı ikinci bir kez daha</strong> sürükleyip, bu sefer bu alanın başka 
bir içeriğini, mesela Sum varken bir de Count'ını aldırabiliriz. Veya yukardaki 
kutunun ikinci sekmesinde bulunan yüzdesel gösterim şekli gibi farklı gösterim 
şekillerini de gösterebiliriz. Bu işlemi yine ilgili kolonda herhangi bir 
hücreye tıkalyıp <strong>Show Values As</strong> seçeneği ile de yapabiliriz. 
Mesela biz % of Column Total seçeneğini seçerek ilgili bölgenin toplam içinde ne 
kadar pay aldığını gösterelim.</p>
	<p>	
	<img alt="" src="/images/insert_pivot_showas1.jpg" height="30%" width="30%" class="zoomla"></p>
	<p>Sonuç aşağıdak gibi olacaktır.</p>
	<p>	<img alt="" src="/images/insert_pivot4.jpg"></p>
	<p>	Keza, bir de tarihsel derinlikteki datamız var ve bu data aylık bazda. Bizden 
	bunun her ay için Yıllık(Kümüle/YTD) versiyonunu da hazırlamamız istendi 
	diyelim.(Tabi bu ürünlerin toplanabilen Yeni Satış veya Gelir Tablosu 
	kalemleri gibi Flow(Dönem içi) tipli ürünler olduğunu varsayıyorum. Ör:Yeni 
	Kredi Kart Adedi, Ücret Komisyon Tahsilatı. Veya bir mağaza için aylık işlem 
	adedi, sipariş adedi gibi.</p>
	<p>	Şimdi bu tarihsel datayı Pivotlayalım.</p>
	<p>	<img alt="" src="/images/insert_pivotshowasrunning.jpg"></p>
	<p>	Show Value as'e tıklayıp sonra da <strong>Running In ... </strong>diyelim ve Ay 
	alanını seçelim. Ben daha okunaklı görünmesi adına <strong>Tabular</strong> 
	görünüm şeklini seçtim ve <strong><a href="#tabularrepeatlabel">Repeat Row Labels</a></strong> dedim. Sonuç 
	aşağıdaki gibi olacak. Gördüğünüz gibi bu yöntem kümülatif toplam almanın 
	muhteşem kolay bir yolunu sunmaktadır.</p>
	<p>	<img alt="" src="/images/insert_pivotshowasruninng2.jpg"></p>
	<h3>Detaya inme(Drilling)</h3>
<p><strong>Values</strong> alanındaki bir hücreye çift tıkladığnızda ilgili kesişim kümesindeki detaylı bilgi yeni bir sayfada açılır. Bu örnekte B4'e(Başkent1 
bölgesinin rakamına) tıkladığınızda karşımıza çıkacak görüntü şu olacaktır.</p>
<p><img alt="" src="/images/insert_pivot_drill.jpg"></p>
	<p>Bu yöntem özellikle kaynak data ile özet tabloyu birbirinden 
	ayırdığımızda kaynak datayı tekrar elde etme yolu olarak pratik bir imkan 
	sunmaktadır.</p>
	<p>Belki gözünüzden kaçmış olabilir, tekrar edelim. <strong>Sadece Values</strong> 
	alanındaki bir hücreye çift tıkladığımızda detaylı dataya(başka bir sayfada) 
	ulaşmış oluyoruz. Label'lara çift tıkladığımızda ise bize, hangi alanın alt 
	kalem olarak tabloda gösterilmesini istediğimiz sorulur. İstenirse seviye 
	seviye aşağı doğru inilir. Mesela biz önce Başkent2'ye çift tıklayalım ve 
	Şubeyi seçelim.</p>
	<p><img src="/images/insertpivotdrilllabel0.jpg"></p>
	<p><img src="/images/insertpivotdrilllabel1.jpg"></p>
	<p>Gördüğünüz gibi + işaretleri otomatik eklendi. Hem de tüm bölgelere, 
	sadece Başkent2'ye değil. Şimdi de Şube10'a çift tıklayalım ve Ürünü 
	seçelim. Sonuç aşağıdaki gibidir:</p>
	<p><img src="/images/insertpivotdrilllabel2.jpg"></p>
	<p>&nbsp;</p>

<h3>GetPivotData</h3>

<p>Özet tablo içinden, belli kesişim noktasına ait bilgiyi çıkarmak için 
GETPIVOTDATA fonksiyonunu kullanırız. Bunu kullanabilmek için Excel Options'ında 
aşağıdaki seçeneğin işaretli olması gerekir.</p>
	<p>
	<img alt="" src="/images/insert_pivotgetpivtooption.jpg" width="80%" height="80%" ></p>
	<p>Bunu ayrıca Pivot Table sekmesindeki Options butonunun yanındaki oka 
	tıklayarak, Generate GetPivotData seçeneğine işaret koyarak da 
	yapabilirsiniz.</p>
	<p>
	<img alt="" src="/images/insert_pivotgetpivot2.jpg"></p>
	<p>Şimdi bu seçenek işaretli değilken aşağıdaki tablomuz üzerinden bu formülü nasıl kullanacağımıza 
	bakalım. A21 ve B21 hücrelerine <a href="DataMenusu_VeriDogrulama.aspx">DataValidation</a> 
	ile bir combobox oluşturdum. C21 hücresine de şu formül yazdım. </p>
	<pre class="formul">=GETPIVOTDATA("Aylık Gerç";$A$3;"Bölge";$A$21;"Ürün Adı";$B$21)</pre>
	<p>
	<img alt="" src="/images/insert_pivotgetpivot3.jpg"></p>
	<p>Resimden görüleceği üzere, formülümüz kesişim rakamını getirdi. Tabi, 
	bunu aynı sayfa üzerinde yapmak anlamsız olabilir, ancak daha büyük Özet 
	Tablolarda, veya özet tablonuz gizlenmiş kolonlarda yer alıyorsa veya başka 
	bir sayfadaysa oldukça kullanışlı bir formüldür. Gerçi başkaları çok faydalı 
	bir formül olduğunu söylese de ben bu formülü pek kullanmam. Kesişimler için 
	kullandığım başka formüller var ve özet tabloları da kaspayacak daha genel 
	formüller. Onları
	<a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx#Kesisim">bu sayfada</a> 
	ele alıyor olacağız. </p>
	<p><strong>NOT</strong>:Yukardaki formülü uzun uzun elle yazmak yerine 
	geçici olarak "<strong>Generate Getpivotdata</strong>" seçeneğini aktive 
	edip, arkasından Values bölümünden herhangi bir hücre seçtiğinizde, Excel 
	size otomatik olarak formülü üretir. Siz sadece bölge ve ürün adını 
	parametrik yapacak değişklik için müdahale edersiniz.</p>


<h3>Generate Report</h3>
<p>Şimdi diyelim ki, yukardaki örnekte, her bölge için ayrı bir sayfada özet 
tablo oluşturmak istiyorsunuz. Tablomuzu şu hale getirelim, yani ROWS alanına 
Şube Adını, ROWS'un hemen üstündeki FILTERS alanına Bölge'yi koyalım.</p>
	<p>
	<img alt="" src="/images/insert_pivotshowreports1.jpg"></p>
	<p>Şimdi bu haldeyken, <strong>PivotTable </strong>alt menüsünde<strong>
	</strong>Options'ın yanındaki küçük oka 
	ve akabinde <strong>Show Report Filter Pages...</strong> butonuna 
	tıklayalım</p>
	<p>
	<img alt="" src="/images/insert_pivotshowreport.jpg"></p>
	<p>Buna tıkladıktan sonra aşağıdaki kutu çıkacak, Bölgeyi seçip OK diyelim..</p>
	<p>
	<img alt="" src="/images/insert_pivotshowreport2.jpg"></p>
	<p>İşlem tamamdır, şimdi aşağıdaki gibi her sayfada ayrı bir bölgenin 
	tablosunun oluştuğunu görebilirsiniz.</p>
	<p>
	<img alt="" src="/images/insert_pivotshowreports3.jpg" class="zoomla" width="80%"></p>







	<h3>Dışardaki bir datayı kaynak olarak kullanmak</h3>
	<p>Bazen kaynak datamız excelde olmayabilir, bunu Access gibi bir 
	veritabanından almamız gerekebilir. Şimdi de böyle bir örnek gösterelim. 
	Kaynak olarak verdiğim excel dosyasına bir access veritabanı içine import 
	edelim(Access bildiğinizi varsayıyorum, eğer bilmiyorsanız şimdilik bu kısmı 
	atlayabilirsiniz)</p>
	<p>Insert PivotTable dediğimizde karşımıza çıkan kutuda, işaretli yere 
	bastığımızda database ve sonrasında tablo seçimini yaparız. bu tablodaki 
	alanlar Özet Tablonun Fieldları olarak sağ panelde yerini alır. Sonra rutin 
	işlemleri uygulayabiirsniz.</p>
	<p><img alt="" src="/images/insert_pivotexternaldata.jpg"></p>
	<p>Burda altı çizilecek bir nokta şu olabilir. Accesteki tablonuzun boyutu çok 
	büyükse kaynak datayı excel içinde tutmanıza gerek olmayabilir, bunun için 
	PivotTable Options'ta Data sekmesine "Save sourca data with file" seçeneğindeki 
	işareti kaldırmanız yeterlidir.</p>







	<h3 id="PC">Pivot Cache ve birden fazla görünüm</h3>
	<p>Bir özet tablo yarattığımız zaman, Excel arka planda bizim göremediğimiz 
	bir kopya data üretir ve bu data üzerinden özet tabloları manipüle eder. 
	Buna <strong>Pivot Cache</strong> denir. Pivot Cache sayesinde,
	<a href="#Data">yukarda</a> gördüğümüz gibi kaynak datayı dosyadan 
	ayırdığımızda bile Özet tablomuz çalışmaya devam eder. Böylece dosya boyutumuz 
	da önemli ölçüde küçülmüş olur. Peki kaynak datayı görmek istersek ne 
	yapıcaz? Grand Total'e çift tıklayarak drill-up yaparız ve işte size kaynak 
	data! Tabi <strong>PivotTable Options&gt;Data&gt;Enable Show details</strong> seçeneğinin işaretli olması 
	lazım.</p>
	<p>NOT:2007 öncesinde PivotCahe ile ilgili önemli bir sorun vardı. Siz aynı 
	özet tablodan farklı görünümler elde etmek için kopya aldığınızda her bir 
	özet tablonunu <span style="text-decoration: underline"><strong>ayrı</strong></span> 
	bir PivotCahesi oluyordu, bu da dosya boyutunu artırmaktaydı. 2007 verisyonuyla 
	birlikte Excel, paylaşılan PivotCahce yöntemini devreye alarak bu sorunu 
	çözmüştür. Ancak yine de bazen Cachelerin ayrı olmasını isteriz, çünkü aynı 
	cacheden beslenen tablolardan birinde refresh yapıldığında hepsi birden refresh 
	olur veya birinde bir <a href="#Grup">gruplama</a> yaptığınızda tüm diğer 
	özet tablolarda da aynı gruplamanın olduğunu görürsünüz. İşte, böyle 
	birşeyin olmasını istemiyorsanız, ki bazen istemeyeceksiniz, cachelerin farklılaştırılması gerekir. Bunun 
	için şu adımları uygulayın:</p>
	<ol>
		<li>İkinci özet tabloyu kesin(Cut)</li>
		<li>Bunu sıfır bir dosyaya yapıştırın</li>
		<li>Sıfır dosyadaki bu özet tabloyu refreshleyin</li>
		<li>Copy ile hafızaya kopyalayın</li>
		<li>Orjinal dosyaya yapıştırın
</li>
		<li>Geçici dosyayı kapatın</li>
	</ol>
		<p>Yeni özettablo şimdi kendi cache'sini kullanacak ve diğerleri refresh 
		olduğunda bu refresh olmayacak, ayrıca yine diğerlerindeki gruplamadan 
		da etkilenmeyecektir. Bunu yapmanın bir yolu da uygun bir VBA kodu 
		çalıştırmaktır. Ancak bunu ve pivotcachelerle ilgili başka neler 
		yapılabilir bilgisini konuyla ilgili
		<a href="../VBAMakro/Ileriseviyekonular_PivotTableChartveSlicernesneleri.aspx#pivotcache">VBA sayfalarına</a> 
		gelince göreceğiz.</p>
	<p>İşlem olduktan sonra PivotCacheleri görmenin tek yolu ise aşağıdaki gibi 
	bir VBA kodu çalıştırmaktır. </p>
<pre class="brush:vb">
Sub cacheleri_gor()
Dim pt As PivotTable
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    Debug.Print pt.Parent.Name, pt.Name, pt.CacheIndex
  Next pt
Next ws

End Sub</pre>

</div>
<h2 class="baslik">Options Butonu ve Field Settings</h2>
<div class="konu">
	<p>Options butonunda 6 sekme bulunur. Bunlara tek tek bakalım. <strong>
	Burada önemli bir nokta var, o da şu: Buradaki yapılacak ayarlamaların her Özet Table için ayrı ayrı olmasıdır. Yani genel 
	bir Pivot Table ayarlaması yapılmamaktadır.</strong></p>
	<h3><strong>Options&gt;Layout &amp; Format sekmesi</strong></h3>
	<p>Burda boş değerlerin(Yani ana datada verisi olmayan) ve 
	hatalı(#N/A, # DIV/0 gibi) değerlerin nasıl gösterileceğine ait seçenekler 
	var. Boş dataların boş görünmesi bazen sıkıntılara sebep olabilmektedir, o 
	yüzden bunların 0 görünmesini isteyebilirsiniz. Hatalı değerler de yine bazı 
	durumlarda sıkıntı yaratır hatta bir de diptoplamın da hatalı görünmesine 
	neden olmaktadır. Bunu da 0 olarak ayarlayabilirsiniz.</p>
	<p><img alt="" src="/images/insert_pivotemptyorerror.jpg"></p>
	<p>Örneğin Başkent1 bölgesinin Ürün3 ve Ürün4ü boş geliyor olsa ve biz bu 
	ayarlamayı yapsaydık tablomuz şöyle görünürdü:</p>
	<p><img alt="" src="/images/insert_pivotemptyorerror2.jpg"></p>
	<p>Tabi yine diptoplamın düzelmediğini görüyoruz ancak, tabloyu Pivottan 
	çıkarıp normal Range haline getirdikten sonra basit bir toplam alma 
	formülüyle istediğimiz sonuca ulaşırız. İdeal çözüm tabiki kaynak 
	datada hatalı sonuçların olmamasını sağlamaktır.</p>
	<h3><strong>Options&gt;</strong>Totals &amp; Filters sekmesi</h3>
	<p>Burada sütun ve satırlarda diptoplam gösterilecek mi gösterilmeyecek mi 
	bunun ayarlaması yapılır. Bunu başka yerlerde de yapabildiğimiz için burada 
	detaya girmiyorum.</p>
	<p>Ayrıca, bir alan üzerinde çoklu filtreleme yapma imkanı da verir. Default 
	olarak bu ayar seçimsizdir. Şimdi diyelimk ki İstanbul bölgerinden 110 mio 
	üzerinde hacmi olan bölgeleri filtrelemek isityoruz. Önce İstanbulları 
	filtreleyelim. <strong>Label Filters&gt;Begins With=İst</strong> 
	yazalım</p>
	<p><img alt="" src="/images/insert_pivotgruplabel2.jpg"></p>
	<p>&nbsp;</p>
	<p><img alt="" src="/images/insert_pivotgruplabel.jpg"></p>
	<p>Şimdi bi de Value Filter alanına greater than 110.000.000 yapalım</p>
	<p><img alt="" src="/images/insert_pivotfiltre100üstü.jpg"></p>
	<p>Gördüğünüz gibi ikinci filtre ilkini ezdi. İkisini aynı anda yapmak için 
	aşağıdaki sarı renkli <strong>Allow multiple filters per field</strong> 
	kutusunu&nbsp;işaretleylim.</p>
	<p><img alt="" src="/images/insert_pivotoptionsallowmultifilter.jpg"></p>
	<p>Şimdi hem Label Filters'ı hem Value Filtersi tekrar uygulayalım, sonuç 
	aşağıdaki gibidir. Bu arada Row Labels'taki filtre işaretine tıkladığınızda 
	her iki Filtrede de Tick işareti olduğunu görebilirsiniz.</p>
	<p><img alt="" src="/images/insert_pivotfiltrelabelanddata.jpg"></p>
	<h3><strong>Options&gt;</strong>Display sekmesi</h3>
	<p>Burda çok kritik bir seçenek yok. Belki eski bildiğimiz Pivot Table formatına dönmemize 
	izin veren seçenek ilginizi çekebilir. Aşağıdaki renkli kutuya tıklayarak bu isteğinize 
	kavuşabilirsiniz.</p>
	<p><a name="klasikPT"></a>Diğer seçenekler de zaten yeterince açıklayıcı olduğu için detaya 
	girmiyorum.</p>
	<p><img alt="" src="/images/insert_pivotdisplayclasik.jpg"></p>
	<h3><strong>Options&gt;</strong><a name="Data">Data sekmesi</a></h3>
	<p>Data menüsündeki önemli özellikler şunlardır.</p>
		<p><img alt="" src="/images/insert_pivototiongrup.jpg"></p>
	<ul>
		<li>Sarıyla işaretli seçenek(Save source data) default seçili gelir. Bunu seçmezseniz, 
		pivot tabloya kaynaklık eden <a href="#PC">PivotCacheyi</a> farklı bir 
		dosyada kaydeder, bu da dosyanın boyutunu küçültür. Çok büyük kaynaklı 
		dosyalarda bu işlemi yapabilirsiniz. Aşağıdaki aynı dosyanın kaynak 
		datasını dosyayla birlikte kaydedilip kaydedilmeme durumundaki boyut 
		farkı görünmektedir.<br>&nbsp;<img alt="" src="/images/insertpivot_dosyaboyut.jpg"><br>
		Bu seçeneğin işaretlenmesiyle dosya boyutu artar ama dosya açılışı daha 
		hızlı olur, seçenek işaretlenmezse boyut küçülür ama açılış hızı süresi 
		uzar, zira o sırada pivotcache yeniden yaratılmaktadır.</li>
		
		<li>Yeşilli seçenek de(Refresh data ...), dosya açılır açılmaz pivot tablonun otomatik 
		güncellenmesini sağlar. Bu, özellikle kaynak datanın gece belli bir 
		saatte schedule edilmiş olması durumunda kullanıcıların dosyayı açtığında 
		güncel datayı görmesi adına faydalı bir seçenektir. Ancak otomatik 
		refreshin rahatsız edici olabileceği durumlarda bu seçenek kapatılıp, 
		kullanıcılara uygun bir bilgilendirme de yapılabilir. <br><br>Eğer, 
		kaynağı dosya içinde kaydetmemeyi tercih ettiyseniz, bu yeşilli seçeneği 
		seçmenizde de fayda var. Aksi halde, bu tablo üzerinde çeşitli filtre 
		v.s işlemi yapmaya çalıştığınızda(veya mail göndermek amacıyla boyut 
		küçülttüyseniz, dosyayı alan kişi işlem yapmaya çalışırsa) şöyle bir 
		uyarı mesajıyla karşılaşır:
            “The PivotTable report was saved without the underlying 
		data. Use the Refresh Data command to update the report.”. Bu bir hata 
		değildir ama tecrübeyle sabittir ki insanlar bunu hata sanıp hemen siz 
		arıyorlar, "dosya bozulmuş" diye. O yüzden dosya açılır açılmaz refresh 
		olsun ki, bu tür şikayetlerle uğraşmayın. Olur da unutursanız, yapılması 
		gereken ilgili pivot tabloyu manuel refreshlemektir.</li>

		<li>Bu yukardaki seçeneklerde renksiz gösterilen seçenek işaretli iken 
		DrillDown yapılabilmektedir, eğer Özet tablo üzerinde bir hücre çift 
		tıklandığında drill yapılmak istenmiyorsa bu işaret kaldırılır.</li>
		<li>Pivot tabloların kaynağı olan alanlarda bazı datalar silinse bile 
		normal ayarlara göre bu data gerek özet tablonun kendisinde gerek bu 
		özettablo üzerine eklenmiş Slicerlarda görünmeye devam eder. Bunların 
		görünmemesi için Data sekmesindeki <strong>Retain items deleted from the 
		data source</strong> seçeneğini None yapmanız gerekir.<br>
		<img src="/images/insertpivotretainnone.jpg" height="290" width="466"></li>
	</ul>

<h3>Sağ tık&gt;(Value)Field Settings</h3>
	<p>Özet tablonun Satır/Sütun alanlarında mı yoksa ortadaki data alanında 
	olup olmadığınıza bağlı olarak sağ tıkladığınızda iki farklı Field Settings 
	kutusu çıkar.</p>
	<p>Önce orta alanda sağ tıkladığımızda çıkan Value Field settingse bakalım. 
	Burada iki sekme bulunmaktadır.</p>
	<p><img alt="" src="/images/insert_pivotfieldvaluesetting.jpg"></p>
	<p>Burdaki seçeneklere aynı zamanda bir hücreye sağ tıklayarak da 
	ulaşılabilmektedir.</p>
	<p><img alt="" src="/images/insert_pivotshowvalueas.jpg"></p>
	<p>İlk sekmedeki seçenekleri zaten tüm örnekler boyunca oldukça kullandık. 
	Burada Toplam, Adet, Ortalma, Min, Max gibi standart&nbsp; hesaplamalar var. Aslında 
	çok önemli bir eksik vardı, <strong>Distinct Count</strong>, o da Excel 2013 ile eklendi, 
	ancak listede doğrudan göremezsiniz. Bunu <span class="keywordler">Data 
	Model</span>'e ekleyerek görebiliyoruz.</p>
	<p>Data modele ekleme konusu PowerPivotla alakalı olduğu ve daha geniş bir 
	yer vermek gerektiği için bunu <a href="YeniEklenenAraclar_PowerPivot.aspx">bu sayfada</a> ayrıca ele 
	alıyoruz.</p>
	<p>Ancak burada Normal Pivotla ilgili olarak Distinc Count konusuna 
	deineceğiz. Özet Tablo hazırlarken bazen bir kolondaki tekil(uniqe) adetleri 
	saydırmak istersiniz. SQL diliyle söyleyecek olursak "Distinct count" almak 
	istersiniz. Excel 2013le birlikte artık bunu Data Modele ekleyerek 
	yapabiliyoruz. Önceki versiyonlarda böyle bir özellik malesef yok. Aşağıdaki 
	göresellerde işlemin nasıl yapılacağı görülmektedir.</p>
	<p><img alt="" src="/images/insert_pivotadddatamodel.jpg"></p>
	<p>&nbsp;</p>
	<p><img alt="" src="/images/insert_pivotadddatamodel2.jpg"></p>
	<p>Gördüğünüz gibi normal Count ilk bölge için 132 sonucnu veriyor çünkü, 4 
	ayrı ürün için çoklama yapıyor. Halbuki bu bölgede 33 şube var, işte bunu da 
	Distinct Count veriyor.</p>
	<p><img alt="" src="/images/insert_pivotadddatamodel3.jpg"></p>
	<p>Value Field Settings'in ikinci sekmesinde ise bir değeri Toplam/Adet gibi standart hesaplama 
	şekilleriyle değil de, birşeyin yüzdesiz olarak gösterme imkanı buluyoruz.</p>
	<p>Aşağıdaki örnekte % of Row Total yaptım:</p>
	<p><img alt="" src="/images/insert_pivotshowpercent.jpg"></p>
	<p>Tabi burda ürünlerin tutarları orantsız olduğu için daha çok son iki 
	ürünün ağırlığı yüksek çıktı. Böyle bir analizden ziyade bir ürün hangi 
	bölgede ne kadar paya sahip, bunu görmek isteyebiliriz. Bunun için de % of 
	Column Total demek gerekiyor.</p>
	<p><img alt="" src="/images/insert_pivotshowperc2.jpg"></p>
	<p>Şimdi örnek dosyamızın ikinci datası olan liste üzerinden bir pivot 
	yapalım. Aylık bazda ürünlerin hacmi ne olmuş, onu görelim. Tablomzu şöyledir.</p>
	<p><img src="/images/insert_pivottarhiseldata.jpg"></p>
	<p>Peki bu tabloda mesela Ürün1 toplamda ne zaman mesela 80 milyonu geçmiş, 
	onu görmek istiyorum, diğer ürünler için de belli eşikleri ne zaman geçmiş 
	görmek istyorum. O zaman daha önce gördüğümüz <strong>Running Total in</strong> .. seçeneğini seçeriz. Base Field 
	olarak da Ay seçeriz ve tabomuz bu hale gelir. Gördüğünüz gibi Ürün1 
	ağustos ayında 80 bandını geçmiş.</p>
	<p><img src="/images/insert_pivotshowas.jpg"></p>
	<p>Rank, Difference gibi diğer seçenekler üzerinden de siz alıştırma 
	yapabilirsiniz. </p>


	<h3>Diğer alanlar</h3>
	<p>Printing ve All text alanlarıyla çok işim olmadı, sizin de çok olacağını 
	düşünmüyorum ancak kurcalamak 
	isterseniz de kurcalayabilirsiniz.<br></p>
</div>

<h2 class="baslik">Menüler</h2>
<div class="konu">
<p>Bir özet tablodaki herhangi bir hücreyi seçtiğinizde Ribbon'da Analyze ve Design isimli iki menü ortaya çıkar. Bunların içindeki bir çok alt menüyü zaten yukarda kısım kısım gördük. Nedendir bilmem, Microsofttaki abiler bazı araçları sadece bir yere değil birden fazla yere koyuyor. Mesela Özet tabloların altında veya sağında diptoplam görünsün mü görünmesin mi kararını hem 
Options butonundan hem de Design menüsünden verebiliyorsunuz.</p>

<p>Bazı özellikler ise kurcalayarak çok kolay keşfedebileceğiniz detayda. O yüzden bunlara da girmiyorum. Şimdiye kadar bahsetmediğimiz bir iki özellik var, onlara bakalım.</p>


<h3 id="Grup">ANALYZE&gt;Gruplama</h3>
	<p>Şimdi diyelim ki Özet tablonuzun satır sayısı çok fazla ve burada 
	gruplanabilecek bazı kayıtlar var. İlk etapta sadece bu grubu görmenin 
	yeterli olduğunu, isterseniz detaya daha sonra inebileceğinizi 
	düşünüyorsunuz. Bunun için verileri gruplama toolunu kullanacağız. Bu 
	örnekte, bölgeleri İstanbul ve İstanbul dışı olarak gruplayalım.</p>
	<p>Gruplayacağım bölgeleri seçiyoruz, sağ tıklayarak Group diyoruz(bunu 
	Analyze menüsünde de yapabilirdik)</p>
	<p><img alt="" src="/images/insert_pivotgrup3.jpg"></p>
	<p>Şimdi aynısını bi de İstanbulları seçerek yapıypruz.</p>
	<p>Gördüğünüz üzere Bölge2 adında yeni bir alan eklendi ve gruplara otomatik 
	olarak isim verildi. Yeni eklenen alanın aynı kolonda mı yoksa yeni açılan 
	bir kolonda mı geleceği, Özet Tablonun formatına bağlı olarak değişir. 
	Compact Form'daysa aynı kolonda gelir, Outline veya Tabular formda ise 
	farklı kolonda. Aşağıdaki örnekte Outlime formda olduğu için farklı kolonda 
	geldi.</p>
	<p><img alt="" src="/images/insert_pivotsgrup4.jpg"></p>
	<p>Biz bu isimleri ilgili hücrelere gelerek ilave bir şey yapmadan direkt 
	değiştirebiliyoruz, aşağıdaki gibi. Group1'i İstanbul Dışı, Group2'yi 
	İstanbul olarak değiştirdim.</p>
	<p><img alt="" src="/images/insert_pivotgrupname.jpg"></p>
	<p>Tarihleri ve sayılar sözkonsu olduğunda gruplamanın özel bir şekli de 
	oluyor. Başlangıcı ve bitişi belli olan gruplar, tek tek seçim yapmadan 
	kolaylıkla oluşturulabiliyor.</p>
	<p><img alt="" src="/images/insert_pivotx.jpg"></p>
	<p>Ben yarıyıllık bir görüntü elde etmek istediğim için aşağıdaki gibi seçim 
	yaptım.(1 Ocaktan başlamadığı için 182 gün demedik, 152 dedik)</p>
	<p><img src="/images/insert_pivotgrupdate.jpg"></p>
	<p><img src="/images/insert_pivotgrupgun2.jpg"></p>
	<p>Yine bölge isimlendirmesinde olduğu gibi burda da ilgili dönemleri 
	manuel değiştirebiliyorum.</p>
	<p><img src="/images/insert_pivotgrupyy.jpg"></p>
	<p>Ancak burda farkettiyseniz, Bölgelerde olduğu gibi Collapse/Expand(+/-) 
	butonları gözükmedi. Eğer bunların gözükmesini istiyorsanız yine manuel 
	seçerek ilerleyebilrisiniz, yani ilk 6 ayı seçip 1.YY diyip, temmuz sonrasına 
	ise 2.yy şeklinde gruplayabilirsiniz. Böyle yapınca +/- butonları çıkar.</p>
	

	<h3>ANALYZE&gt;Fields, Items &amp; Sets(Calculated Fields&amp;Items)</h3>
	<h4>Calculated Field</h4>
	<p>Bazen elinizdeki kaynak datada eksik bir kolon olduğunu görürsünüz. Bunun için bir çözüm şekli, 
	kaynak tabloya bu kolonu eklemek olabilir. Bir diğer çözüm ise Özet Tablo işlemi uyguladıktan 
	sonra manuel formül yazmaktır, ama bunun neden ideal çözüm olmadığını da az 
	sonra göstereceğim. İşte, Calculated Fields bu iki yönteme bir alternatiftir.</p>
	<p>Şimdi diyelim ki aşağıdaki gibi bir tablomuz var, bunu Kanal bazında 
	özetleyeceğiz. Ama oransal bir bilgi olan faiz oranını da özet tabloda 
	görmek istiyoruz. Oransal kalemleri özet tablolara doğrudan almak 
	doğru değildir, zira bunların toplanmaycağı aşikardır, bu örnekte 
	ortalama aldırmak da doğru değildir, çünkü her kredinin tutarı farklıdır. Bu 
	yüzden bunların 
	ağırlıklı ortalamasını almak gerekir. Bunun için de pay ve paydayı ayrı ayrı 
	toplayıp, bu toplamlar üzerinden işlem yapmak gerekir. Bütün bu işlemi 
	datayı çektiğmiz SQL'de de yapabilirsiniz ancak bu datanın daha uzun sürede 
	gelmesine neden olur.</p>
	<p><img alt="" src="/images/insertpivotsetitem1.jpg"></p>
	<p>Şimdi ilk olarak, hatalı sonuca bakalım. Yani doğrudan, faiz oranının ortalamasını alalım.</p>
	<p>
	<img alt="" src="/images/insertpivotsetitem2.jpg"></p>
	<p>Şimdi de öncelikle manuel formül(SUMIF) yazarak sonuca ulaşalım, ama bu çok 
	sağlıklı bir yöntem değildir, zira satır ekleme/eksilme durumlarında, mesela 
	belli bi anda kanal sayısı 4 olabilir, başka zaman 2 olabilir, böylece manuel formül 
	kolonunuda eksik veya fazla satır olabilir. Ayrıca diptoplam için girilen 
	formül SUMIF değil, SUM olacak, yine burdaki satır sayısı değiştirdikçe 
	buna da müdahale etmek gerekecektir. Bu yüzden rakamı doğru vermekle 
	birlikte bu yöntem ideal çözüm 
	değildir.</p>
	<p>
	<img alt="" src="/images/insertpivotsetitem3.jpg"></p>
	<p>
	Şimdi ideal çözüm olarak Calculation Field yaratmaya bakalım.</p>
	<p>
	<img alt="" src="/images/insertpivotsetitem4.jpg"></p>
	<p>Yeni alanımız sağdaki panelin en altında yerleşir.</p>
	<p>
	<img alt="" src="/images/insertpivotsetitem5.jpg"></p>
	<p>Sonuç da aşağıdaki gibi olur.</p>
	<p>
	<img alt="" src="/images/insertpivotsetitem6.jpg"></p>
	<p><strong class="dikkat">Dikkat</strong>:Calculated Fieldlarda formüle yazdığınız her şey ayrı ayrı işleme 
	girer. Ör:Calculated Field'daki formülünüz A*B şeklinde 
	ise, sol taraftaki row label bazında A'ları toplar, sonra da 
	B'leri toplar. Bu toplamları çarpar, yani ilgili row label kalemini 
	oluşturan tüm satırlar için A*B yapıp da bunları toplamaz. O yüzden biz 
	Tutar*Faiz yapmadık, Faizgeliri/Tutar yaptık. İhtiyacınız böyle birşeyse Table içine A*B şeklinde bir kolon hazırlayın, sonra 
	da bunu Özet tablo içine 
	normal bir alan olarak alın.</p>
	<p><strong>NOT</strong>:Hesaplanmış alanlar, normal alanlardan faklı olarak sadece Value 
	olarak kullanılır, yani Row veya Column'a gelemezler. Bunlar için Calculated 
	Item kullanıyoruz.</p>
	<p>Bir diğer husus da Calculated fieldlarda sadece Toplama işlemi yaplır. Field settingsten ortalama 
	veya adet seçseniz 
	bile işe yaramaz.</p>
	<h4>Caldulated Item</h4>
	<p>Bazen de Rows veya Columns'ta yer alan alanlardan yeni bir alan türetmek 
	istersiniz. Mesela Başkent1 ve Başkent2 diye iki bölge var diyelim, 
	bunlardan Başkent diye bir Ana bölge türetmek isteyebilirsiniz. Bunu da 
	Calculated Item ile yapıyoruz. Bu biraz da yukarda bahsettiğimiz Gruplamaya 
	benziyor, ancak gruplamadan bir farkı, burda +/- düğmeleriyle açıp daraltacağımız 
	bir formatın oluşmaması, özet tablomuzun adeta yeni bir kayıt eklenmiş gibi 
	görünmesidir. Daha büyük farkı ise burada iki veya daha fazla şeyi 
	toplamaktan daha karışık formüller de yazabiliyor olmamız.</p>
	<p>Mesela aşağıdaki işlemi Gruplama ile de yapabiliriz. Görüntü açısından 
	fark dışında pek de bir fark yok gibi görünüyor.</p>
	<p><img src="/images/insert_pivotcalcitem.jpg"></p>
	<p>Ancak bazen Gruplama yetersiz kalır, işte böyle durumlarda mecburen 
	Calculated Item yaratmak gerekir.</p>
	<p>Bu arada yukarda, elemanlara doğrudan isimleriyle ulaştık, bunlara index numarası ile de ulaşılabilir. Bu index numarası da 
	mutlak veya göreceli olabilir. Mutlak index, ilgili alanın hangi sütun veya 
	satırının seçileceğini belirtirken, göreceli index ise yeni hesapladığımız 
	Calculated itema olan uzaklığı belirtir. Yeri her defasında sabit olan 
	alanlar için mutlak index kullanılması gerekirken yeri sürekli değişebilen 
	durumlarda göreceli index kullanmak gerekir. Bunlarla ne demek istediğimi az sonraki örnekle daha iyi 
        anlayacaksınz. Göreceli başvurularda + ve - işaretlerini kullanırız.</p>
	<p>Biz şimdi iki yöntemi de kullanacağımız bir örnek yapalım.&nbsp;Mesela tarihsel datamızda Ay alanı için bir Calculated Item 
	yaratalım. İhtiyacımız olan şey de hep son ay ile ilk ay rakamlarının oranını 
	yani yıllık büyümeyi gösteren bir alan. Bunu yeni bir Ay elemanı gibi düşündüğümüz için 
	Calculated Item yapıyoruz. Datamızın ilk hali aşağıdaki gibi ve yeni 
	alanımız en sona eklensin istiyoruz.</p>
	<p>
	<img src="/images/insert_pivotcalcitem2.jpg" ></p>
	<p>Bu örnekte ilk ay datasının yeri sabittir ve hep birinci kolondur, o yüzden 
	buna Ay[1] ile ulaşacağım, yeni alanımızı en sona koyacağız, yani son ay 
	kolonu da hep bundan bir önceki kolon olacak, o yüzden buna da Ay[-1] göreceli 
	başvurusu ile ulaşacağım. Bu durumda yeni yaratacağımız Item'ın formülü 
	aşağıdaki gibi olacaktır. </p>
	<p><img src="/images/insert_pivotcalcitem3.jpg"></p>
	<p>Şimdi diyeceksiniz ki, neden Ay[-2] yazdık, az önce Ay[-1] demiştik. 
	Çünkü bir de hep son ay bir önceki aya göre ne kadar büyümüşüz buna bakmak 
	istiyorum, bunun için de hep son 2 ayın farkına bakacak formülü yazarım, ki 
	bu formül sadece göreceli başvuruları içeriyor olacak.</p>
	<p>
	<img src="/images/insert_pivotcalcitem4.jpg"></p>
	<p>Şimdi bu da yeni bir kolon olarak geleceği ve ilk yarattığımız Calcuated 
	Item'ın bundan sonra görünmesini istediğim için, bundaki göreceli indexi -1 yerine -2 yaptım. </p>
	<p>Sonuç aşağıdaki gibi olacaktır.(Başkent2'nin 
	Ocak rakamı eksik olduğu için DIV/0 hatası çıkmış ve bu hata diptoplama da 
	sirayet etmiş, buna şimdilik takılmayın, gerekirse IFERROR ile hataları 
	sıfır getirtebilirsiniz)</p>
	<p><img src="/images/insert_pivotcalcitem5.jpg"></p>
	<p>Bir de Solve Order diye bir şey var, oluşturduğunuz Calculated Itemların 
	sırasını belilersiniz. Bu özellikle, önce A hesaplansın, B'nin hesabında A 
	kullanılacak tarzı durumlar varsa önem arzetmektedir. Ben burda bu detay 
	girmiyorum, böyle bir ihtiyaç olursa bu ekrandan o işi kolayca 
	halledebilirsiniz.</p>
	<h3>DESIGN&gt;Layout</h3>
	<p>Diyelim ki "Bölge Kodu - Ürün - Ay - Aylık Gerç" formatındaki (örnek 
	dosyada tarihsel data sayfasında) tablonuzu bölge detayı olmadan yani bölge 
	toplamında görmek istiyorsunuz. Çünkü ya nihai tablo üzerinde vlookup işlemi 
	kullanacaksınız, veya başkalarına liste olarak göndereceksiniz, bu kişiler 
	filtreleme v.s yapmak istediklerinde yapamayacak. Artık bunu Excelde yapmak 
	çok kolay. Hemen bakalım:</p>
	<p>Aşağıdaki gibi Pivot alma işlemini yapın,</p>
	<p><img alt="" src="/images/insert_pivototodoldur1.jpg"></p>
	<p><a name="tabularrepeatlabel"></a>Sonra, <strong>Design&gt;Report Layout&gt;Show in Tabular Form</strong> 
	seçeneğini seçin, rapor aşağıdaki gibi görünecektir.(Bir diğer seçenek de 
	Outline Form gösterim şeklide veya Classic Pivot Table gösterimi olabilir)</p>
	<p><img alt="" src="/images/insert_pivototodoldur2.jpg"></p>
	<p>Son olarak<strong> </strong>yine <strong>Design&gt;Report Layout&gt;Repeat all 
	Item labels </strong>seçeneğini seçin. Bu işlemi Field settingsi 
	açıp(tabloda herhangi bir yere sağ tıklayarak veya Analyze menüsünden)&nbsp; 
	Layout&amp;Print sekmesinden de yapabilirsiniz.</p>
	<p>Ve işte yeni tablomuz!</p>
	<p><img alt="" src="/images/insert_pivototodoldur3.jpg"></p>
	<p>Gördüğünüz gibi vlookup yapmak için ideal bir liste haline geldi. Sanki 
	ana tablodan bölge kolonu çıkarılmış&nbsp; ve rakamlar otomatik olarak 
	ürün-ay seviyesinde yeniden hesaplanmış gibi. Bu şahane özellik Excel 2010 
	ile aramıza katıldı. Bundan önce ben bu işi Excelent'ta bulunan "pivotta oto doldur" makrosu ile hallediyordum, 
	böylece uzun yıllar kullandığım ve birçok kişinin de kullandığı bu makrom da 
	çöp olmuş oldu :) Ancal Hala 2010 öncesi Excel kullananlar
	Excelent ile istenilen sonuca ulaşabilirler.</p>


	<h3>Slicer ve Timeline</h3>
	<p>Excelin 2010 versiyonu ile Özet Tablolara Slicerla kolay filtre uygulama 
	imkanı gelmiştir. Slicer özelliği Excel 2013 ile birlikte Listeli Tablolara da uygulanabilir 
	hale geldiği için bu konuyu başka bir sayfaya aldım.
	<a href="DataMenusu_SiralamaveFiltreleme.aspx#Slicer">Buradan</a> bakabilrsiniz.</p>
	<p>TimeLine ise, yine Özet tablolara uygulanabilen bir zaman filtreleme 
	yöntemidir. Geniş bir zaman aralığına ait bir datayı özet tablo haline 
	getirdiniz diyelim, ancak belli dönemlerde sadece belirli bir yıla, aya, çeyreğe veya günlere ait veriyi görmek ve hatta bunu bir grafik 
	eşliğinde incelemek isteyebilirsiniz. Aşağıda farklı bir data setine ait bir 
	özet tabloya hem Slicer hem TimeLine filtresi uygulandığnı görebilrsiniz. 
	Slicer'dan sadece TL seçilmiş, Timeline'dan da Kasım ayı seçilmiş. Tabi 
	Kasım'ı seçebilmem için Zaman frekansının Ay olarak seçilmiş olması 
	gerekmektedir, bunu da görselde görebilirsiniz.</p>
	<p><img alt="" src="/images/vbapivottimeline.jpg"></p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
</div>




</asp:Content>
