<%@ Page Title='Makro Nedir' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Makro Nedir</h1>
<h2 class="baslik">Tanım</h2>
<div class="konu">
	<h3 >Makro mu VBA mi?</h3>
	<p>Eve bu ana bölümde özetle "makro yazmayı öğreneceksiniz" diyebiliriz ancak <strong>Makro</strong> ile eş anlamlı kullanılan bir ifade de VBA'dir. VBA, Visual Basic(VB) Programlama dilinin 
	Office programları için uyarlanmış bir versiyonudur. 
	Excel makrolarını yazacağımız dil de <strong>VBA(Visual Basic Applications) for 
	Excel</strong> olarak bilinir, yani Excel için Visual Basic uygulamaları.(Access, 
	Outlook ve PowerPoint için de VBA mevcuttur ancak en popüleri Excel için 
	olanıdır)</p>
	<p >
	Makro terimi yanlızca Office programlarına özgü bir ifade değildir aslında, 
	birçok programda Makro özelliği bulunabilir. Özünde makro, adımlardan oluşan 
	bir program bütündür. Excel için söyleyecek olursak, bir hücreyi seçmek, 
	içine bir değer girmek, bi sayfayı silmek gibi. </p>
	<p >
	Makroları yazarken VBA kullanacağımız için de genel olarak Makro ve VBA 
	ifadeleri birbirinin yerine kullanılabilir ama şunu söyleyebilirm ki VBA 
	ifadesini kullanmak sizi biraz daha profesyonel gösterir :)</p>

	<h3 id="Neden"> Ne işe yarar?</h3>
	<p>	Basitçe Excelin zaten güçlü olan işlevselliğini daha da ileri götürmek gibi 
	bir temel görevi olduğunu söyleyebileceğimiz Makroların, belli başlı 
	kullanım amaçları sonraki paragrafta verilmiştir, tabiki bunlar 
	çoğaltılabilir.</p>

	<p>
	Bunlara geçmeden önce söylemeden geçemeyeceğim bir husus var. Makroların 
	gücünü tam olarak bilmeyenlerin/görmeyenlerin genelde Exceli ve VBA'i/Makroları küçümseme 
	gibi bir huyu vardır. Onlara göre bırakın VBA'yı, VB bile doğru dürüst bir 
	programlama dili değildir. Evet, VB'den daha gelişmiş dillerin olduğu doğru 
	ancak, elimizdeki tool Excel ise Excelin işlevselliğini de ileriye götüren 
	en güçlü ve kolay ulaşılabilir tool da VBA'dir.(Başka dillerle de, VB.Net, 
	C#, C++ ile de Interop dll'leri aracılığı ile Excel otomasyonu sağlanabilir, 
	ki bunlardan VB.Net ile neler yapılabileceğini VSTO ana menüsünde ayrıca ele 
	alıyorum). Her ne olursa olsun, VBA ile yapılabileceklerin de bir sınırı var, 
	ki bence bu sınırlar oldukça geniştir, ancak kimsenin VBA ile bir masaüstü 
	program yazma veya Facebook gibi bir site yapma iddiası da yok. Bir diğer 
	uçta ise, öğrenmeye açık olmadığı için makrolarla neler yapılabileceğini 
	hayal bile edemeyen kişiler var, ve bunlar da makroların gücünü 
	önemsemezler. Ana sayfada ne demiştim, "Excel düşündüğünüzden değil düşünebileceğinizden 
	bile daha karmaşık bir yapıya sahiptir". Neyse, bu iki uç grubu biz de görmezden gelelim, olumluya odaklanalım ve 
	makrolarla hem kendimizi hem de bu iki grup dışındaki çoğunluğu oluşturan 
	insanları nasıl büyüleyeceğimize bi bakalım:</p>

	<ul>
		<li>Rutin yapılan işler başta olmak üzere birçok işin otomatiğe 
		alınması(Böylece hatasız ve aşırı hızlı işlem yaparsınız)</li>
		<li>Excelin kendisiyle yapılamayan veya yapılsa bile aylar yıllar sürecek döngüsel işlemler gibi işlemlerin 
		yapılması(Ör:1000 satırın herbirine ayrı ayrı goal seek yapmak gibi)</li>
		<li>Bazı durumlarda yetersiz kalan Excel fonksiyonlarının yerine kendi 
		fonksiyonlarınızı yazabilirsiniz</li>
		<li><strong>Makrolar işyükünüzü azaltır, siz yokken bile çalışabilir. Kurumunuza 
		verimlilik sağlar, maliyetleri düşürür</strong>.</li>
		<li>Bazı durumlarda insanların size Uzaylı gibi bakmasını sağlar :)</li>
		<li>Beyninizin sürekli çalışmasını, dolayısıyla yeni sinir 
		bağlantılarının oluşmasını sağlar. Bu da daha geç bunayacağınız ve 
		üstelik daha az kilo alacağınız anlamına gelir, evet kilo almak dedim, 
		çünkü beyin, vücudun en çok enerji tüketen organıdır ve ne kadar çok 
		çalışırsa o kadar çok enerji tüketir:)</li>
	</ul>
	<p>
	Daha özele inecek olursak şunları yapabilirsiniz</p>

	<ul>
		<li>Satış ve performans takip raporları oluşturulabilir</li>
		<li>Çok kanallı/şubeli kurumlarda her alıcıya kendisiyle ilgili 
		dosyaların gönderimi sağlanabilir</li>
		<li>Çeşitli frekanslarda raporların belli gün ve saatlerde kendiliğinden 
		çalışıp refresh olması sağlanabilir</li>
		<li>Uyarı mekanizlamaları kurulabilir</li>
		<li>Çeşitli yerlerden toplanan dosyalar birleştirilebilir</li>
		<li>Bütçeler oluşturulabilir</li>
		<li>Dashboardlar oluşturulabilir</li>
		<li>İş ararken sizi diğerlerinin önünde tutar :)</li>
		<li>Bunamanızı yavaşlatacağı için yaşlandığınızda torunlarınızın adını 
		hatırlayabilirsiniz :)</li>
	</ul>
	<p>
	Tabiki, Excel'in her yeni versiyonu ile yazdığımız bazı makrolar gereksiz 
	hale gelebilmektedir ve sizin için de bu geçerli olacaktır. Mesela, yıllar 
	önce bir hücrenin formül içerip içermediğini kontrol eden bir function 
	yazmıştım, ancak artık Excelin 2013 versiyonunda bu formül dahil edilmiş. O 
	yüzden bu kod artık gereksizdir. Keza yine, Pivot Table yaptıktan sonra 
	aradaki boşlukları otomatik dolduran bir makrom vardı, bu da yine 2010 
	versiyonu ile birlikte gereksiz hale geldi. Ama karamsar olmayın, Microsoft 
	çalışanları ne yaparsa yapsın kuruma özgü, spesifik ihtiyaçları karşılayan araçlar 
	geliştiremezler, o yüzden VBA bilen birisi olarak yine her zaman el üstünde 
	kalmaya devam eder, iş ararken avantajınızı korursunuz.</p>

	<h3>Ne bilmek gerekiyor?</h3 >
	<p>Visual Basic bilenler için makro öğrenmek çok daha basit olmakla birlikte ilk defa 
	makro öğrenecek kişilerin gidip de öncelikle VB öğrenmesine gerek 
	bulunmamaktadır. Bu site zaten size direkt olarak makro yazımını öğretmeyi 
	hedeflemektedir. Her ne kadar site uzman kullanıcılar için tasarlanmış olsa 
	da makrolar zaten ayrı bir uzmanlık gerektirdiği için en başından 
	anlatılacaktır.</p>
	<p>
	Bu sitedeki öğrenme sürecinize paralel olarak, Excelin Makro Kaydetme 
	özelliği ile basit denemeler yapabilirsiniz. Böylece yaptığınız her 
	hareketin sonunda nasıl bir kod ortaya çıktığını takip edebilirsiniz. Zaten 
	ne kadar profesyonel olursanız olun yeri geldiğinde Makro Kaydet aracını 
	kullanmanız gerekecek. Sonra ihtiyaca göre oluşan kodda istediğiniz 
	düzenlemeleri yapabilirsiniz.</p>
	</div>


	<h2 class="baslik">Nasıl başlayacağız<a name="personal"></a></h2>
	<div class="konu">
	<h3>Organizasyon</h3>

	<p>Yazdığınız makrolara sık sık ulaşmak isteyecek, Add-in 
	aktifleştirme/pasifleştirme gibi işlemleri yapacak ve tabiki makro kaydetmek 
	isteyeceksiniz. Bunların ne anlama geldiğini bilmiyor olabilirsiniz ama emin olun iyi şeyler. İşte bunları yapmak 
	için öncelikle Ribbona sağ tıklayarak Developer sekmesini etkinleştirelim, ayrıca VBE'yi(VisualBasic 
	düzenleyicisini) QuickAccess Toolbarına almanızda fayda var.(Alt+F11 tuşuyla 
	da ulaşılabilir)</p>
		<p>Makrolarınızı derli toplu tutmanın birkaç yolu bulunmaktadır. Ben 
	sizlere, bunlardan Personal.xlsb ve Add-ins yaratma yöntemlerinden 
	bahsedeceğim. Aslında ikisinin de kendine özgü amaçları vardır. Kendimin, 
	hangisini ne amaçla kullandığımı söylersem size de ışık tutacaktır diye 
	düşünüyorum.</p>
	<p>Çalıştığım kurum içinde diğer kişilerle de paylaşacağım makrolar varsa 
	bunları Add-in olarak hazırlarım, hazırladığım makroların bir menü olarak 
	Ribbonda gözükmesini istediğim için de bu Add-in içine de menüleri oluşturan 
	bir 'başlangıç' makrosu yazarım, en sonunuda Add-ini diğer kişilerle 
	paylaşırım. Add-in yöntemiyle diğer kişilere ne yapmaları gerektiğni 
	anlatmak daha kolay olur. Onlara sadece Add-in'i nasıl kurmaları gerektiğini 
	anlatan kısa bir mail atarım, ondan sonra makroları kullanmaları çok kolay 
	olur. </p>

	<p>Personal.xlsb dosyası ise daha çok kendinize özel makroları içerir, 
	diğerlerinde bu makroların olmasına gerek yoktur. Benim Personal.xlsb 
	dosyamda scheduling kodları, kısayol kodları, bana özel fonksiyonlar gibi 
	özelleşmiş kodlar bulunur, bunların bir kısmını ilerleyen sayfalarda burada 
	paylaşacağım. Kısayol tuşu atadığınız makroların bu dosyada bulunması çok 
	önemli, çünkü Add-in içine koyarsanız ve başka kişiler başka amaçla bu 
	kısayollara tıklarsa yanlış sonuçlarla karşılaşabilirler ve malesef 
	makroların Undo'su yoktur. Örneğin ben Copy 
	PasteSpecial Value için Ctrl+M kombinasyonunu kullanırım. Bir 
	hücredeki sayılara binlik ayraç uygulamak için Ctrl+L kısayolunu. Halbuki bu kısayolların 
	bazılarında önceden MS tarafından tanımlanmış başka görevleri var olabilir 
	ve kişiler bu kısayol tuşlarını zaten o amaçlar için kullanıyor olabilir.</p>

	<p><strong>NOT:</strong>Excel, 2007 sürümünden sonra standart format içine makro kaydetmeye izin 
	vermiyor, burada Excel dosya formatlarına çok değinmeyeceğim ancak
	<a href="../Excel/Giris_DosyaUzantilari.aspx">şurda</a> 
	detaylı bilgi var, linkteki bilgiden de görüleceği üzere <strong>"Personal" dosyanızı 
	.xlsb uzantısıyla kaydetmelisiniz.</strong></p>

	<h3>Personal.xlsb'nin konumu</h3>
	<p>Personal.xlsb dosyası herhangi bir dosya değildir. Bu dosyaya kaydedilen 
	makroları bütün dosyalarınızda kullanabilirsiniz. Çünkü Excel oturumu 
	boyunca hep açık kalmaktadır. Herhangi bir Excel 
	dosyasındaki makrolar ise sadece o dosya açık olduğu sürece çalışacaktır.</p>
	<p>Peki bu dosya nasıl oluşturulur ve nereye kaydedilir? Şu adımları takip edin:<em>(Adımları 
	sonuna kadar okuyup öyle uygulayın)</em></p>
	<ul>
		<li>Boş bir Excel dosyası açın</li>
		<li>View menüsünden dosyayı gizleyin</li>
		<li>Excelden çıkmaya çalışın</li>
		<li>Excel sizi uyaracak ve az önce gizlediğiniz dosyayı kaydedip 
		kaydetmek istemediğini soracaktır</li>
		<li>Evet deyin ve XLSTART klasörüne kaydedin. Office sürümüne göre bu 
		klasörün yeri değişebilmektedir, bunu tespit etmenin kolay bir yolu var<ul>
			<li>Alt+F11 ile VB editörünü açın</li>
			<li>Immediate Window açık değilse Ctrl+G ile bunu açın ve oraya 
			<strong>Application.StartupPath</strong> yazıp Enter'a basın.</li>
		</ul>
		</li>
	</ul>
	<p>
	<img height="60%" src="../../images/vba_giris_startup2.jpg" width="60%" class="zoomla"></p>
	<p>
	Dosyanızı oluşturduktan sonra, her projeniz için ayrı bir modül 
	oluşturmanızı(ilerleyen konularda anlatılacak) ve bu projeyle ilgili tüm 
	prosedürlerinizi bu modül içinde bulundurmanızı tavsiye ederiz.</p>
	<p>
	<strong>NOT: </strong>Bazı kurumlarda, BT politakaları gereği bazı klasörlere erişim izni 
	olmamaktadır, XLSTART klasörü de bunlardan biri olabilmektedir. Bu 
	nedenle Personal.xlsb dosyanız için alternatif bir klasör belirlemeniz 
	gerekebilir, bunu da <span class="keywordler">Excel 
	Options&gt;Advanced&gt;General&gt;At Startup....</span> kutusuna yazarak belirtebilirsniz</p>
	
	
	<p>
	<img height="60%" src="../../images/vba_giris_startup.jpg" width="60%" class="zoomla" ></p>
		<p>
		<strong>NOT</strong>: Personal.xlsb dosyasını doğrudan Record Macro 
		yaparak da olşuturabilirsiniz. Size makronuzun nereye kaydedileceği 
		sorulur, Personal Macro Workbook(Türkçe Excel'de "Kişisel Makro Çalışma 
		Kitabı") seçeneğini seçerseniz bu dosya 
		otomatikman oluşur, ancak yukarda belirttiğim gibi BT politikanız 
		XLSTART klasörüne erişim izni vermiyorsa sorun oluşabilir.</p>

	<h3 id="guvenlik">Güvenlik Ayarları</h3>
	<p>Makrolar, kötü amaçlar için kullanılabilir, tabiri caizse makrolar içine virüs yazılabilir. Mesela isterseniz(ki aslında 
	istememelisiniz) bir add-in yazıp, belirli bir tarih geldiğinde kullanıcıların bilgisayarındaki önemli 
	dosyaların silinmesini sağlayabilirsiniz. 
		<p>Şunu düşünebilirsiniz; Ya makroyu bi düğmeye, menüye atamıyor muyuz, 
		veya VBE üzerinde o makroya gelip de F5 yaparak çalıştırmıyor muyuz, biz 
		bi şeye basmadan nasıl zararlı olabilir ki? Cevap basit:İleride 
		göreceğiz ki Workbook eventlerinden biri de Workbook_Open() eventidir(bir de Auto_open() var) ve dosya 
		açılır açılmaz aktive olur. İşte o sırada olanlar olur :)</p>
	<p>
	İşte bu nedenle makroların kullanılabilmesi için çeşitli güvenlik ayarları 
	ve seçenekleri mevcuttur.</p>
	<p><strong>File&gt;Options&gt;Trust Center&gt;Trust Center Settings</strong> düğmesine&nbsp; 
	tıklayarak bu ayarların olduğu yere geliriz.</p>
	<p>Burada ilk göz atacağımız yer Makro Settings menüsüdür.</p>
<img height="60%" src="../../images/vba_giris_macrosettings.jpg" width="60%" class="zoomla" >
	<p>İlk seçenekle hiçbir makroya güvenmediğinizi belirtmiş olursunuz, makro 
	içeren tüm dosyaların makrosu pasifleştirilir.</p>
		<p>İkinci seçenek varsayılan 
	seçenektir ve makro içeren bir dosya karşısında sizi uyarır. Siz de enable 
	veya disable diyerek ilerlersiniz.</p>
		<p>Üçüncü seçenek benim tercih ettiğim 
	seçenektir. Eğer bir makro dijital olarak imzalandıysa ve <abbr title="Dijital olarak imzalanmış bir makro için sertifika almış kişi veya kurumlar.">Güvenilir Yayıncılar</abbr> bölümüne eklendiyse uyarı çıkmadan makroyu 
	etkin kılar, aksi halde size uyarı çıkarır, yani ikinci seçeneğin biraz 
	gelişmiş şeklidir. Yanlız hiç imza yoksa direkt pasifleştirir.</p>
		<p>Son 
	seçenek ise pek güvenli değildir, zira bütün makroları aktif kılar, bu da 
	sizi hackerlara açık hale getirir, o yüzden bilmediğiniz kaynaktan gelen ve 
		güvenmediğiniz dosyalarla çalışırken makro ayarlarınızı kesinlikle buna 
		getirin. </p>
		<p>Haa tabi bu arada <a href="../Excel/Giris_DosyaUzantilari.aspx">dosya 
		uzantıları </a>bölümünde gördüğümüz gibi Excel'in 2007 versiyonundan 
		sonra standart dosya tipi içine makro kaydedilememektedir, o yüzden 
		standart Excel dosyalar güvenlidir diyebiliriz.</p>
		<p>Biz şimdilik öğrenim sürecinde olduğumuz için son seçeneği aktif 
		yapalım, ancak unutmayın, öğrenim ve test süreciniz bitince 3.seçeneği 
		işaretleyelim.</p>

	<h3>İzinler</h3>
	<p>Evet, makro ayalarını yaptınız ama bi süre sonra Excelin çıkardığı 
	uyarılar canınızı sıkmaya başlar. İkide bir çıkan bu uyarılardan güvenlik 
	sınırları içinde kurtulmanın yolları da var elbette. Şimdi bunları 
	inceleyelim.</p>

	<h4>Dijital İmza</h4>
	<p>Dijital imzayı ticari bir kurumdan alabileceğiniz gibi kendinizin 
	imzalayacağı sertifikalar(self-signed certicates(SSC)) da olabilmektedir.</p>
	<p>Kendi kullanımınız için veya ekip arkadaşlarınız makrolarınızı kullanacaksa 
	SSC kullanabilirsiniz. Bu, makronuzu kullanacak kişilere "bana güven; ben, 
	kim olduğumu söylediğim kişiyim" demek oluyor.</p>
	<p>Ticari kurumlardan alınacak imzalar ise daha güvenlidir ve makronuzu 
	kullanacak kişiye "Bana güven, X firması benim kim olduğumu biliyor ve benim 
	o kişi olduğumu teyit ediyor" demek oluyor.</p>
	<p>İmzalar, sadece makroyu yazan kişinin kim olduğunu belirtmekle kalmaz, bu 
	kişinin elinden çıktıktan sonra hiçbir şekilde değişime uğramadan geldiğini 
	de gösterir, bir nevi mühür görevi görür diyebiliriz.</p>
	<p>Bu konu hakkında internette bol bilgi bulunuyor, bunun detayın burada 
	daha fazla girmeden bir SSC nasıl oluşturulur ona bakalım.<br></p>
	<p><span><span class="dikkat">Dikkat</span></span>:Dijital sertifika, makroyu kimin yazdığını 
	gösteriyor olmakla birlikte, içindeki kodun güvenli olduğunu garanti etmez. 
	Dosyaya güvenmek, tamamen kullanıcıya kalmıştır, bu da onun makroyu yazan 
	kişiye güvenip güvenmemesiyle ilgilidir.</p>
	<p>Peki, bu SSC nasıl alınır. SelfCert.Exe dosyası ile. Windows ve Office 
	sürümlerindeki farklılıklar nedeniyle bu dosyanın konumu değiştiği için 
	dosyayı doğrudan Windows Explorer içinde aramanızı öneririm.</p>
	Dosyayı çalıştırdığınızda şu görüntü çıkacak, isminizi yazın ve OK 
	diyin.<p>
	<img alt="" src="../../images/vba_giris_ssc.jpg"></p>
	<p>Bu işlemi sadece birkez yapmış olacaksınız. Ondan sonra her dosyanız için 
	aşağıdaki işlemi yapmanız gerekecektir.</p>
	<p>VBE'ye geçin, sol panelden imza atayacağınız dosyayı seçin,<span>&nbsp;Tools&gt;Digital 
	Signature&nbsp;</span>yolunu takip edip Choose düğmesi aracılığıyla dosyanızı 
	imzalayın.&nbsp;</p>
				<p><img src="../../images/vba_giris_ssc3.jpg"></p>

	<h4>Güvenilen Yayıncı listesine birini eklemek</h4>
	<p>Arkadaşınızın imzaladığı makrolu bir dosyayı açtınız ve güvenlik uyarısı çıktı, 
	şimdi güvendiğiniz kullanıcıya ait imzalı bir makroyu 
	listeye ekleyeceksiniz.</p>
	<p>Trust Settings içinde '<strong>Trust all documents from this publisher</strong>' düğmesine tıklayarak bu 
	kullanıcıyı güvenilir yayıncılar listesine eklemiş olursunuz.</p>

	<h4>Güvenilir Lokasyon</h4>
	<p>Diyelim ki bir makro içeren dosyayı açacaksınız, imzası yok ama 
	biliyorsunuz ki güvenli. İkide bir uyarı çıkmasın istiyorsunuz, ama makro 
	ayarlarını da daha düşük bir seviyeye çekmek istemiyorsunuz. O zaman onu 
	güvenli alan olarak adlandırılan yere alırsınız ki Excel güvenlik merkezi 
	birdaha sizi rahatsız etmesin.</p>
		<p>
		<img height="60%" src="../../images/vba_giris_trustedloc.jpg" width="60%" class="zoomla"></p>
	<p><span class="dikkat">Dikkat</span>:Belgelerim(Documents/My Documents) klasörünü 
	'güvenilir yer' olarak işaretlemek yerine bu klasör altında başka bir klasör 
	açın, onu işaretleyin. Aksi taktirde tüm Belgelerim klasörünüzü hackerlara 
	açık hale getirmiş olabilirsiniz.</p>
	<p>Ek bilgi:Otomatik olarak güvenilir yer olarak gelen bazı klasörler de 
	bulunmaktadır.</p>
	</div>
	
<h2 class="baslik">Önemli Uyarı!!!</h2>	
<div class="konu">

	<p>Makrolar gerçekten çok faydalı araçlardır, ancak tıpkı faydalı ama yanlış 
	kullanıldığında felakete neden olabilen diğer herşeye benzerler. Tüpgazlar 
	faydalıdır ama yanlış kullanım sonucu patlar. Bıçaklar faydalıdır ama yanlış 
	kullanım yaralanmalara neden olur. O yüzden kullanımları dikkat ister.</p>
	<p>Makrolarla ilgili de dikkat edilmesi gereken iki konu var.</p>
	<ul>
		<li>Makroyu çalıştırdığınızda <strong>Undo yapamazsınız</strong>. O yüzden makroyu 
		çalıştırmadan önce üzerinde çalıştığınız dosyanın son halinin bir 
		yedeğini almanızda fayda var. İsterseniz her makronuzun önüne, 
		MsgBox ile bu soruyu sordurabilirsiniz.</li>
	</ul>
<pre class="brush:vb" style="margin-left: 40px">
cevap = MsgBox("Dosyanın yedeğini aldın mı?", vbYesNo)
If cevap = 6 Then 'bu yes demek oluyor 
    GoTo ilerle
Else
    MsgBox "O zaman yedeği al sonra tekrar çalıştır. "
    Exit Sub
End If

ilerle:

'Diğer Kod bloğu
</pre>
	<ul>
		<li>İkinci dikkat edilecek konu da, makronuzu her yönüyle test etmeden 
		çalıştırırsanız istisnai bir durumda farklı bir sonuçla 
		karşılaşabilirsiniz. Bu yüzden olası her durum için çalıştırmanızda 
		fayda var. Makronuzu özellikle başkaları kullanacaksa, onlara nasıl 
		kullanacaklarını hiç anlatmadan direkt çalıştırmalarını söyleyin, olacak 
		her hata için de gerekli kontrolleri, yönlendirme 
		sorularını ve <a href="DebuggingveHataYonetimi_HataYakalama.aspx">hata 
		denetim elemanlarını</a> koyun. Mesela, makronuzun doğru 
		çalışması için ele alınacak sütunun A sütunu olması gerekebilir, bu 
		nedenle kullanıcıya "İşlem yapılacak bilgi A sütununda değil mi?" 
		diye bir soru yöneltebilirsiniz. Bunların hepsini yeri geldiğinde öğreneceğiz.</li>
	</ul>

</div>
	
</asp:Content>
