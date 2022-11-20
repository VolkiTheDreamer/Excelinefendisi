<%@ Page Title='DigerUygulamalarlailetisim AccessProgramlama' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Diğer Uygulamalarla iletişim'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Veritabanı Programlama</h1>
	<h2>Terminoloji</h2>
	<p>Veritabanı/Database(VT/DB) işlemlerine geçmeden önce genel terminolojiye 
	hakim olmak, konuları anlamak açısından oldukça faydalı olmaktadır. O yüzden 
	ilk olarak terminoloji ile başlayacağız. <br>Kanımca, burdaki terimleri 
	içselleştirmeden ilerlemek sizin VT bağlantılarını tam anlamıyla 
	anlamayacağınızı, ezbere ve copy-paste(internetten veya makro recorderdan) 
	kodları çalışmak zorunda kalacağınızı garantiler. Bu da kodunuzun çirkin ve 
	uzun görünmesine neden olacağı gibi, zamanla “ODBC neydi, neden bunda OLE DB 
	olmuş da ODBC olmamış, bunda neden CommandText de şunda CommandType yazıyor” 
	gibi soruları sormaya başladığınızda sorularınızın cevapsız kalmasına neden 
	olacaktır.</p>
	<p>Terminoloji konusu çok kapsamlı bir konu olmakla birlikte ben buraya 
	özellikle VBA ile doğrudan veya dolaylı alakası olan konuları koymaya 
	çalıştım. İnternette de bol miktarda makale var, ben de bir nevi kendi 
	tecrübelerim ile bu makalelerin derlemesini yaptım diyebilirim, yoksa kendim 
	bir VT uzmanı değilim.</p>
	<h3>Genel Terimler</h3>
	<p>İlk etapta bu bölümde yazılanlardan genel bir fikir edinmeye çalışın, 
	herşeyi anlamaya çalışmayın, zaten bu konulara aşina değilseniz biraz 
	karmaşık gelebilir. Sonrasında hemen diğer Veritabanı sayfalarına geçin. Ancak 
	veritabanı bağlantıları konusunda deneyim sahibi oldukça bu bölüme gelip bilgilerinizi 
	tazeleyin.</p>
	<p><strong>COM(Component object model):</strong> Programalama dilleri, hatta 
	platformlar arası uyusmuzlukları ortadan kaldırma amacıyla ortaya çıkan bir 
	standarttır. Microsoft(MS)’un teknolojisi değildir.</p>
	<p><strong>ActiveX</strong>:Nesnelerin yaratıldıkları dilden bağımsız olarak 
	birbiriyle etkileşmesini sağlayan MS’a özgü COM teknolojisidir.</p>
	<p><strong>API(Application Programming Interface)</strong>:Uygulamaların 
	birbiriyle anlaşmasını sağlayan program paketleri.</p>
	<p><strong>OLE</strong>: MS’un bir uygulmasında bir başka MS uygulamasının 
	içeriğini gösterebilme teknolojisi. Ör: Exceldeki bir grafiği Word içine 
	kopyalayabilirsiniz. Ve bu Word dosyasını birisi açıtığında bilgisayarında 
	Excel yüklü olmasa bile sorunsuz açılır. Bu da bir COM türüdür.</p>
	<p><strong>MDAC</strong>: Microsfotun MDAC(Microsoft Data Access Components) 
	adı verilen teknolojileriyle programlama dillerinden VT bağlantıları 
	yapabilmekteyiz. Bunlar 3 önemli API’yi içerir. <br>&nbsp;<br><strong>ADO,OLE 
	DB,ODBC</strong><br>Bu konuda o kadar çok araştırma yaptım ki, bunları 
	spesifik olarak tanımlayan bir ifade bulamadım. Bunlar için farklı yerlerde 
	API, Model, Standart, Teknoloji gibi isimler kullanılmaktadır. Her ne kadar 
	MSDN sitesi bu 3’ünü aynı kategori içinde ele aldıysa ben araştırmalardan 
	edindiğim izlenime göre bunları Data Sağlayıcıları ve Erişim Arayüzleri 
	olarak iki sınıfta ele alacağım, zira bu 3 terime yapacağım eklemeler de 
	olacak.</p>
	<h3>Veri sağlayacılar</h3>
	<p>Bir veritabanı üzerinde data okuma, yazma veya güncelleme gibi işlemlerin 
	yapılabilmesi için bu kaynaklara ulaşılması gerekmektedir. Veritabanı 
	yönetim sistemlerine erişim yapabilmek için iki temel yöntem bulunur: ODBC 
	ve OLE DB.</p>
	<h4>ODBC(Open Database Connectivity)</h4>
	<p>MS’un veritabanlarına programlama dillerinden erişmek için ortaya attığı 
	standarttır ancak sonunda bir endüstri standardı olmuştur. Neredeyse tüm VT 
	şirketleri kendi ürünlerine ait ODBC driverı(sürücüsü) geliştirmektedirler.</p>
	<p>ODBC, uygulamaların veritabanlarına erişimini sağlayan bir arayüz 
	oluşturur. Driverlar aracılığı ile uygulama ve veritabanı arasında köprü 
	oluştururlar. Driverlar SQL komutlarını veritabanına iletir ve sonuç 
	kümesini döndürür.</p>
	<p>ODBC ile sadece ilişkisel veritabanı(verinin satır ve sütunlar şeklinde 
	tutulduğu) erişim sağlanır. Oracle, SQL Server, DB2, MySql, Access gibi. Bu 
	yüzden iletişim için yanlızca SQL metni kullanır.</p>
	<p>Bununla birlikte 
	ODBC’yi doğrudan VBA ile temasa sokamıyoruz, çünkü low level(alt seviye) 
	bileşenleri var. Onun yerine onu sarıp sarmalayıp dokunabilir hale getiren 
	arayüzlerle(DAO gibi) dokunabiliyoruz, ki buna birazdan gelicez.</p>
	<p>ODBC dünyasında akış şöyledir:<strong> Uygulama(Ör:VBA)--&gt; 
	SQL--&gt;DAO--&gt;ODBC--&gt;Veritabanı</strong>&nbsp;</p>
	<h4>OLEDB(Object Linking and Embedding Database)</h4>
	<p>OLE DB, ODBC’den sonra gelmiştir. MS’un ODBC’ye şekil verip onu ActiveX 
	modeline soktuğu standarttır. COM tabanlıdır. Her tür veritabanınaına erişim 
	için tasarlanmıştır. Yani hem relational(ilişkisel) hem de non-relational(ilişkisel 
	olmayan) 
	veritabanlarından veri çekebilemtekdir. Non-relational’a örnek olarak Email 
	sistemlerini, Text dosyalarını, Excel dosyalarını verebiliriz.</p>
	<p>ODBC’de driver ne ise OLE DB’de provider odur. Yani veriye providerlar 
	aracılığı ile ulaşılır. OLE DB ayrıca ODBC’ye bir köprü de atarak, ODBC 
	driverlarının da kullanımını sağlar. Provider listesinde “OLE DB provider 
	for ODBC Driver” olarak görünen şey budur. Amacı, ODBC ile ADO’yu 
	konuşturmaktır. Çünkü az önce söylediğimiz gibi ODBC normalde sadece 
	DAO(sözkonusu uygulama VBA ise) ile konuşuyor. Neden ADO ile konuşmak 
	istediğine az sonra geleceğiz.</p>
	<p>OLEDB, iletişim için SQL dışında da teknikler kullanır(XML gibi). 
	Outlooktaki belli kişilerden gelen mailleri filtreleme işlemi de OLEDB ile 
	yapılıyor.</p>
	<p>ODBC’den farklı olarak, OLE DB’yi doğrudan kullanabilmekle birlikte 
	genelde VBA ile doğrudan temasa sokamıyoruz, çünkü bunun da ODBC gibi low level 
	bileşenleri var. Onun yerine onu sarıp sarmalayıp dokunabilir hale getiren 
	arayüzlerle(ADO) VBA’ye iletişme sokuyoruz, ki buna birazdan gelicez.</p>
	<p>OLEDB dünyasında akış şöyledir:<strong> Uygulama(Ör:VBA)--&gt;SQL/XML/v.s 
	--&gt;ADO--&gt;OLE DB--&gt;Veritabanı(RDBMS veya nonRDBMS)</strong></p>
	<p><strong>NOT</strong>:Bir ara MS’un ODBC’yi desteklemeyi durduracağı, tüm 
	geliştirmeyi OLE DB üzerinde yapacağı ve ODBC’nin OLE DB tarafından replace 
	edileceği söyleniyordu ama son zamanlarda(2012’den beri) rüzgar tersine 
	esmeye başladı, ODBC tekrar gözde oldu. En büyük sebep de performans olarak 
	gösteriliyor. Gelecekte ne olacağını öngörmek zor.</p>
	<h3>Erişim arayüzleri(API’ler)</h3>
	<p>Bu kısımda da yukardakilerden çok da ayrılamayan ama mantık olarak 
	birbiriyle daha ilişkili olan diğer terimlere, API’lere, bakacağız. Bunların 
	hepsi birbiriyle ilişkili, ancak daha çok “ODBC vs OLE DB” ile “ADO vs DAO” 
	karşılaştırmaları yapıldığı için ben de böyle bir gruplama yaptım. Yoksa hem 
	yukarda kısmen gördüğünüz hem de birazdan göreceğiniz gibi ADO ile OLE DB de 
	oldukça örtüşen kavramlar.</p>
	<p>VBA’de veri erişim arayüzü olarak 3 teknoloji bizlere sunulmuş durumda. 
	Veri erişim arayüzü ne demek? Veriye erişim için bir nesne modelinin 
	hazırlanmış olması demek. VBA de nesneye dayalı programlamayı desteklediği 
	için bu yöntemlerle veriye erişiriz. Bu 3 teknoloji şunlardır:</p>
	<ul>
		<li><strong>DAO (Data Access Objects):</strong>Yerel veritabanları 
		için(Access gibi) geliştirildi. Hatta Access'e özgüdür bile diyebiliriz. 
		ODBC ile bağlanır. İlk bu çıkmıştır.</li>
		<li><strong>RDO (Remote Data Objects):</strong>Oracle gibi daha büyük 
		çaplı veri tabanları için geliştirildi, DAO’nun hafıza ve yüksek kaynak 
		tüketimi sorunları nedeniyle ortaya çıktı. Biz bunu hiç kullanmayacağız.</li>
		<li><strong>ADO (ActiveX Data Objects)</strong>:Sonradan geldi ve her 
		tür VT için kullanılmaya başlandı. MS’un OLE DB teknolojisiyle de 
		etkileşim içinde çalışır. ADO’da DAO’ya göre daha az nesne ama daha çok 
		nesne üyesi(özellik, metod, olay) bulunur.</li>
	</ul>
	<p>Neden 3 ayrı yöntem tane var diye soracak olursanız teknolojinin sürekli gelişmesini 
	ve müşteri beklentilerindeki değişkliği söyleyebiliriz. Mesela DAO’yu web server üzerinde 
	kullanamıyoruz ancak eski kullanıcıları da üzmemek için eskilere destek sürüyor. 
	Bunların en sonuncusu ve güçlüsü&nbsp; ADO’dur diyebiliriz. O 
	yüzden yeni kodlarımızda mümkünse ADO kullanmalıyız, ancak eski kodları da anlamak 
	adına ve sadece Access içinde kalacaksak DAO’yu da öğrenmekte fayda var.</p>
	<p>Sonraki sayfalarda bunlara detaylı bakacağız, ancak şimdi kısaca 
	bahsetmek istiyorum. Bu arada önemli bir hatırlatma 
	yapmak isterim. Gerek DAO gerek ADO olsun, bunlarda 
	<span style="text-decoration: underline">veritabanına linkli bir 
	bağlantı </span>kurmuyoruz, yani datayı refreshlenebilir(sağ tıklayıp 
	Refresh düğmesiyle güncellenebilir) bir şekilde almıyoruz. Bunun yerine veritabanına bağlanıp datayı okuyoruz(Tabi 
	istersek aynı zamanda 
	datayı değiştirebiliyoruz/silebiliyoruz da). Refreshlenebilir data 
	bağlantısı(ki sadece bağlantı kurulur, editleme/silme yapılamaz) için başka 
	teknikler var, onları da sonraki sayfalarda göreceğiz.</p>
	<h4>DAO</h4>
	<p>Microsoft, DAO’yu, Acces'in kullandığı Microsoft JET Database’ine erişim 
	sağlamak için geliştirmiştir. DAO, COM-tabanlı olup ODBC bağlantıları için 
	kullanılır.</p>
	<p>Accessle çalışırken DAO kullanmak daha iyidir diyebiliriz. DAO’nun ADO’ya göre daha çok nesnesi 
	vardır ve Access'e özgüdür demiştik. Bu arada “Accessle çalışırken” 
	ifadesinden 
	kastım Access VBA değil, Excel VBA içinden Accese bağlanmayı kastediyorum. Aynı 
	şekilde aynı kodlar tabiki Access VBA içinde de kullanılabilir ancak Access 
	VBA için Access nesne modelini bilmek gerekir.</p>
	<p>DAO, akılda Jet DB motoru bulundurularak tasarlandığı için ADO seçenek 
	bile olmamalı diyenler de var. MS tarafından ADO’nun DAO’nun tüm özelliklerini taşıyacak 
	şekilde geliştirileceği söylenmiş ancak bu henüz olmamıştır. Ben şahsen 
	Access'le sadece DAO kullanıyorum, size de bunu öneriyorum. Diğerlerinde ADO 
	kullanıyorum.</p>
	<p>Bununla beraber Access 2007’den itibaren DAO yenilenmiş ve 
	ACEDAO adını almıştır, Jet(Joint Engine Technology) motoru da ACE(Access 
	Connectivity Engine) motoru tarafından replace edilmiştir.</p>
	<h4>ADO</h4>
	<p>ADO, MS’un sonradan ortaya çıkardığı, veriye erişim için kullanılan COM 
	nesneleri kümesidir. Programlama dili ve OLE DB arasında katman sunar. Bizim 
	kapsamımızda programlama dili VBA oluyor. VBA kodlamacısı olarak bize, 
	veritabanının nasıl ele alındığını bilmeden dataya erişen programlar yazma 
	imkanı verir. SQL bilmeye bile gerek yoktur, bununla berabaer SQL komutları 
	da çalıştırılabilir. 4 Collectionu ve 12 nesnesi vardır. Kullanımı basittir. 
	Bu detaylara bilahare değineceğiz.</p>
	<p>Yukarıda belirttiğimiz gibi OLE DB ayrıca ODBC’ye bir köprü de atarak, 
	ODBC driverlarının da kullanımını sağlar. Amacı, ODBC ile ADO’yu 
	konuşturmaktır. <br></p>
	<h3>Kullanım ve Farklar</h3>
	<ul>
		<li>DAO ODBC ile, ADO hem OLE-DB hem wrapper sayesinde ODBC ile çalışır.</li>
		<li>DAO’da daha çok nesne daha az metod, ADO’da az nesne çok metod 
		bulunur.</li>
		<li>DAO daha hızlıdır.</li>
		<li>DAO Accese özgü, ADO tüm veritabanlarına uygun(Aynı şey ODBC ve 
		OLEDB için geçerli değil. Tamam; VBA ODBC’ye ulaşırken ara katman olarak 
		DAO’yu kullanır ancak bu sadece VBA ve Access ikilisi içindir, başka 
		programlama dillerinden başka DB’lere bağlanılabilir, fakat onlar için 
		ara katmanın ne olduğunu bilmiyorum) </li>
		<li>DAO web uygulamalarında kullanılmaz, ADO kullanılır.</li>
	</ul>
	<p>Bundan sonraki sayfada DAO ve ADO’yu daha detaylı ele alacağız. Yalnız, bu iki 
	konunun tüm detaylarını ele almayacağız, zira oldukça geniş konular. İşimize 
	yarayan konulara ağırlık verilecek, MIS ağırlıklı bir kişi olarak kullanma 
	ihtimali olan bazı detaylara da girilecek ama kullanma ihtmaliniz olmayan 
	veya çok düşük olan detaylar atlanacaktır. Daha detaylı bir öğrenim için 
	lütfen MSDN veya diğer kaynaklara bakabilirsiniz.</p>
	<p><strong>NOT</strong>: "Ben iki ayrı teknolojiyi bilmekle uğraşmayayım 
	sadece tek şeyi bilsem olmaz mı?" diyorsanız, o zaman size önerim direkt 
	ADO’ya bakın.</p>


	</a>


</asp:Content>
