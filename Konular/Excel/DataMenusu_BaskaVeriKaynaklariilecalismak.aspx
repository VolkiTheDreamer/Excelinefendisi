<%@ Page Title='DataMenusu Dış Veri Kaynakları' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Data Menüsü'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Başka Veri Kaynakları ile çalışmak</h1>

	<h2 class="baslik"> Giriş</h2>
		<div class="konu">
	<p> Öncelikle belirtmek isterim ki, Excel'in farklı sürümlerini 
	kullananlarda Data menüsü farklılık göstermektedir. Şöyle ki, Excel 2016'dan 
	itibaren <strong>Power Query </strong>teknolojisi bir add-in olmaktan çıkıp 
	Excel'in asli bileşenlerinden biri olmuştur ve dış kaynaklarla çalışmak için 
	bu teknolojinin kullanılması beklenmektedir.</p>
	<p> 2016'yla gelen <strong>Get &amp; 
	Transform</strong> düğme grubu işte bize Power Query çözümlerini 
	vermektedir. Gerçi 2016'yı da kendi içinde yine iki ayrı gruba ayırmalıyız. 
	Zira Office 365 çatısı altında kullananlar ile 365 olmayan sürümü 
	kullananların Data menüsü de farklıdır. En iyisi bu farklara direkt ekran 
	görüntülerinden bakalım.</p>
	<p> 2016 öncesinde data menüsü aşağıdaki gibiydi, sadece <strong>Get 
	External Data </strong>grubu vardı.</p>
	<p> <img src="../../images/datapq1.jpg"></p>
	<p> 365'siz Excel 2016'da ise menü şöyledir. <strong>Get &amp; Transform
	</strong>grubu yeni geldi ancak mevcuttaki <strong>Get External Data</strong> 
	hala duruyor. </p>
	<p> <img src="../../images/datapq2.jpg"></p>
	<p> 365'li Excel 2016'da ise <strong>Get External Data</strong> grubu artık 
	yok.</p>
	<p> <img src="../../images/datamenuyeni.jpg"></p>
	<p> 365li Excel'de aradığımız herşeye <strong>Get Data</strong> butonunun 
	altındaki menülerden ulaşmamız gerekiyor. Bu yeni menünün özelliği artık 
	herşeyi Power Query tabanlı çalıştırıyor olması. Arkada kullandığı veri 
	sağlayıcı ise klasik <strong>Oledb</strong> değil, <strong>Oledb.Mashup</strong>'tır. Oledb hakkında detay 
	bilgi için
	<a href="../VBAMakro/DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">
	buraya</a> bakınız. </p>
	<p> Ancak "ben Power Query'yi sevmedim, ona bi türlü alışamadım ve alışmak 
	da istemiyorum, eski yöntemleri kullanmak istiyorum" (ki bence böyle demeyin 
	ve bir an önce Power Query'yi öğrenmeye çalışın) diyorsanız ve tabi 365 Excel kullanıyorsanız, eski dostlarınıza <strong>
	File&gt;Options</strong> üzerinden aşağıdaki <strong>legacy</strong> kısmından 
	istediklerinizi seçerek kavuşabilirsiniz. Burada <strong>MS Query</strong>'yi aramayın, o 
	zaten Get Data altında, <strong>From Other Sources </strong>içinde duruyor.</p>
	<p> <img src="../../images/datapq3.jpg"></p>
	<p> Bu seçimi yapınca Get Data altında <strong>Legacy Wizards</strong> 
	gelir.</p>
	<p> <img src="../../images/datapq4.jpg"></p>
	<p> Biz bu bölüme Power Query'ye değil, eski yöntemlere bakacağız. Power Query 
	ve diğer Power BI araçlarına <a href="YeniEklenenAraclar_Konular.aspx">bu 
	sayfadan</a> ulaşabilrsiniz.</p>
	<p> Ayrca dış veriye VBA(Makro) ile ulaşma yöntemlerini öğrenmek için
	<a href="../VBAMakro/DigerUygulamalarlailetisim_Konular.aspx">buraya</a> 
	tıklayın ve Veritabanıyla olan linkleri inceleyin.</p>
	<p> Şimdi sırayla farklı veri kaynaklarından veri nasıl çekilir bir bakalım.&nbsp; 
	Burdaki tüm kaynaklara yine Legacy Wizards 
	altından erişeceğiz.</p>
	</div>
	<h2 class="baslik">SQL Server ve Access'ten veri çekme</h2>
	<div class="konu">
	<h3> SQL Server</h3>
	<p> Ben evimdeki PC'den bu sayfaları hazırladığım için localhostta bulunan 
	bir database yarattım ve oradan veri çekeceğim.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz1.jpg"></p>
	<p> Server adı olarak siz de ilgili server adını girebilirsiniz. Credential 
	olarak ilgili servera nasıl bağlanıyorsanız onu seçin, Windows açılış 
	bilgileriyle girebilecğeiniz gibi ayrı belirlenmiş bir kullanıcı adı ve 
	şifreniz de olabilir. Sonra Next deyin;</p>
	<p> Bağlanmak istediğiniz veritabanını ve tabloyu seçin.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz2.jpg"></p>
	<p> Next deyin;</p>
	<p> <img src="../../images/datafromsql6.jpg"></p>
	<p> Finish'e basın, sonuç aşağıdaki gibi gelir.</p>
	<p> <img src="../../images/datafromsql7.jpg"></p>
	<p> Şimdi bu tablo üzerinden herhangi bir hücre seçiliyken ya Design 
	menüsünden veya Data menüsünden Properties'i bulup seçin, çıkan pencerede 
	aşağıdaki kırımızlı butona basın,</p>
	<p> <img src="../../images/datafromsql8.jpg"></p>
	<p> Açılan pencerenin Definition sekmesine geldiğinizde bağlantı 
	bilgileriniz(Connection String) ve bağlandığınız tablo/sorgu/sql metni neyse 
	onu görebilir ve gerektiğinde burada istediğiniz değişkliği yapabilirsiniz. </p>
	<p> <img src="../../images/vbaimportsqldataconwiz3.jpg"></p>
	<p> Şuan doğrudan bir tabloya bağlandığımız için Command type=Table olarak 
	görünmektedir ama siz bunu SQL olarak değiştirip aşağıdaki gibi bir SQL 
	yazabilirsiniz.</p>
	<p> <img src="../../images/datafromsql9.jpg"></p>
	<p> Klasik(legacy) yöntemle SQL Servera bağlanırken malesef bağlantı anında 
	SQL yazamıyoruz, mecburen bu yukardaki yöntemi kullanıyoruz. Ancak Power 
	Query bağlantılarında ilk bağlantı sırasında da SQL metni yazabilyoruz.</p>
	<p> Bu arada eğer çoklu tablo seçimi yaparsak aralarında ilişki kurmamıza 
	sağlayan bir araç olan Data Model(Excel 2013) ile ister table olarak ister 
	Pivot Table olarak bir çıktı oluşturabilriiz. Ancak bu konuya da yine Power BI 
	araçlarını işlediğimiz yerde göreceğiz.</p>
	<p> Özet olarak iki tablo arasında <strong>join kurma</strong> yöntemlerine bakacak olursak;</p>
	<ul>
		<li>SQL Server üzerinde(Management Studio gibi bir tool ile) joinleyip 
		bir sorgu(view) olarak kaydetmelk</li>
		<li>Yukarda bahsettiğim gibi Data Model ile(Power BI araçlarında 
		göreceğiz) joinlemek</li>
		<li>Command Text kısmına joini sağlayan SQL yazmak</li>
		<li>MS Query üzerinde birleştirmek(aşağıda göreceğiz)</li>
		<li>Legacy yöntem yerine Power Query bağlantısını sağlayan bir bağlantı 
		kurmak(iset SQL yazarak ister grafiksel aryüzde bağlantı kurarak)</li>
	</ul>
	<h3>Access</h3>
	<p>Bu sefer kaynak olarak Accessi seçelim. Sonra hangi access dosyasını 
	istiyorsak çıkan pencerde de onu seçelim. Karşımıza aşağıda liste çıkacaktır.</p>
	<p><img src="../../images/dataimportaccess1.jpg"></p>
	<p>Tablo seçimimizi yaptıktan sonra, son pencere çıkar, </p>
	<p><img src="../../images/dataimportaccess2.jpg"></p>
	<p>Biz Table olarak getirmek istiyoruz. Tablelarla neler yapılabildiğine
	<a href="HomeMenusu_Tablolar.aspx">buradan</a> bakabilirsiniz. Eğer ki 
	bir önceki "Select Table" kutusunda çoklu tablo seçimine izin verirsek 
	bunlar dersek data modele yüklenir ve seçilen 
	tablolar joinlenerek buradan pivot bir tablo üretmemiz beklenir. Yukarda 
	belirttiğim gibi bu Data Model konusu Power toollarını gerektirdiği için şimdilik bu detaya girmiyoruz, 
	ve sonucu Table olarak getiriyoruz.</p>
	<p>Bu Legacy yöntemle SQL Serverda olduğu gibi sadece tablo veya sorgular import edilebilir 
	ve join yöntemi olarak SQL serverda yazılanlar geçerlidir.</p>
	<h3>Design Menüsü</h3>
	<p>Şimdi hem SQL Server hem Accesste elde ettiğimz Table'a ait Design menüsünden 
	Connectionlarla ilgili 
	olarak neler yapılabildiğine bakalım.</p>
	<p><strong>Convert to range:</strong> Bu işlem hem veri kaynağı ile olan bağı koparar hem 
	de Table formatını(görüntü olarak değil işleyiş olarak) bozar.</p>
	<p><strong>Unlink:</strong> Bu işlem ise sadece kaynakla bağı koparır, Table 
	formatı kalır.</p>
	<p id="properties"><strong>Properties</strong>:</p>
	<p><img src="../../images/dataimportaccess3.jpg"></p>
	<p>İlk penceredeki birçok şey açıkça kendisini anlatıyor, o yüzden onları 
	ayrıca burada açıklamama gerek yok sanırım. </p>
	<h5>Olası formül sorunu</h5>
	<p>Bir önemli nokta var ki,
	<span style="color: rgb(0, 0, 0); font-family: &quot;Trebuchet MS&quot;, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(248, 248, 248); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;">
	eğer tablomuz oluştuktan sonra bir kolona manuel olarak bir formül yazdıysak 
	sonraki refreshlerde yeni gelen data için formüllerin aşağı inme sorunu 
	yaşanabilmektedir. Bu olay Properties penceresinde <strong>Preserve column 
	sort/filter/layout </strong>seçeneğinin işaretlenmediği durumda gerçekleşir. 
	Çözümü (Bu çözüm bu sayfayı yazdığım 2018'de kullanılmaktaydı, belki 
	ilerleyen yıllarda Microsoft buna bir çözüm bulmuş olabilir) ise şöyledir:</span></p>
	<p>
	<span style="font-size: 16px; letter-spacing: normal; background-color: #F8F8F8">
	Ben tabiki formülü manuel olarak aşağı çekmekten bashetmeyeceğim, bu sadece 
	geçici biz çözümdür zira her refreshte bu sorun devam edecektir. Gerçek 
	çözüm şöyledir: </span></p>
	<ul>
		<li>
		<span style="font-size: 16px; letter-spacing: normal; background-color: #F8F8F8">
		Preserve column sort/filter/layout seçeneğini işaretleyin, </span></li>
		<li>
		<span style="font-size: 16px; letter-spacing: normal; background-color: #F8F8F8">
		İlgili kolonu silin ve kolonu tekrar oluşturup formülünüzü yazın. Formül 
		otomatik aşağı inecek ve sonraki refreshlerde de düzgün 
		çalışacaktır.(Tüm kolonu silmek yerine ilk satır hariç tüm içeriği 
		silip, sonra formülü aşağıda indirerek de yapabilirsiniz)</span></li>
	</ul>
	<p>
	<span style="font-size: 16px; letter-spacing: normal; background-color: #F8F8F8">
	Aşağıdaki örnekten gidecek olursak, öncelikle Propertiesten ilgili seçeneği 
	kaldırdım.</span></p>
	<p><img src="../../images/dataimportformul1.jpg"></p>
	<p>Sonra Accese gidip tabloya bir satır ekledim. Ve Excelde refresh yaptım. </p>
	<p><img src="../../images/dataimportformul2.jpg"></p>
	<p>Ggördüğünüz gibi bir satırın formülü gelmedi, üstelik bu satır yeni 
	eklediğime ait değil. Normalde 1866.satırda olmasını beklerdik ama 1867de 
	oluştu. Bunun sebebini tam bilmiyorum ama sanırım datayı rasgele bir sırada 
	çektiği için olsa gerek.</p>
	<p>Sonra propertiese gidip seçeneği tekrar işaretledim, ilk hücre hariç tüm 
	içeriği sildim.&nbsp;</p>
	<p><img src="../../images/dataimportformul3.jpg"></p>
	<p>Son olarak Accese gidip bi satır daha ekledim, ve Excelde Refresh yaptım, 
	bu sefer tüm hücrelerde formül geldi.</p>
	<p><img src="../../images/dataimportformul4.jpg"></p>
	<h4>Connection Properties</h4>
	<p>Bu pencerenin diğer önemli kısımları <strong>Connection&gt;Name </strong>yazan yerdeki düğmeye 
	basınca çıkar.</p>
	<p>Usage sayfasında Refresh control kısmında yazanlar önemli. Eğer belirli 
	bir aralıkta dosyanın refresh olmasını istiyorsanız bunu <strong>Refresh 
	every 60 minutes</strong> yazan yerde yapabilirsinz, bu da departman 
	ortasındaki bir televizyona bağlanmış Güncel Rakamsal Dashboard fikri için 
	güzel bir imkan sağlamış olur. Keza dosya açılır açılmaz refresh olmasını 
	istiyorsanız da bir alttaki seçenek işaretlenir.</p>
	<p><img src="../../images/dataimportaccess4.jpg"></p>
	<p>Definition kısmında ise daha önce bahsettiğim gibi Connection string ve 
	Tablo/View(Query)/SQL metinleri bulunur. Bunlar manuel olarak veya makro ile 
	değiştirilebilirler.</p>
	<p><img src="../../images/dataimportaccess5.jpg"></p>
	<p><strong>Edit Query</strong> ve <strong>Parameters</strong> bu bağlantı yöntemlerinde pasif gelir. 
	Aktif geldiği kısımlar ve kullanımları için MS Query kısmına bakın.&nbsp;</p>
	</div>
	<h2 class="baslik">Text/Csv ve XML veri kaynaklarından veri çekme</h2>
		<div class="konu">
	<h3>Text/csv</h3>
	<p>Bir text dosyasını import etmek için Legacy'den Text'i seçelim. Önümüzde 
	iki ana seçenek vardır. Eğer ilgli dosyada kolonlar belirli 
	karaketerlerle(virgül, boşluk v.s) ayrılmışsa <strong>Delimited </strong>
	seçeneğini seçip ilerleriz, ki benim şuana kadar karşıma çıkan dosyaların 
	neredeyse hep bu formattaydı. Diğer format ise kolonların sabit uzunlukta 
	birbirinden ayrıldığı <strong>Fixed </strong>formattır.</p>
	<p><img src="../../images/dataimporttxt1.jpg"></p>
	<p>
	<strong>Start import at row</strong>:Genelde 1 bırakılır</p>
	<p>
	<strong>File origin</strong>:Eğer türkçe karakterler de varsa 1254-turkish windows 
	seçilir.</p>
	<h4>
	Delimited formatı</h4>
	<p>
	2.adımda uygun delimiter seçilir, eğer seçeneklerden biri burada yoksa Other 
	içine uygun delimeter yazılır.</p>
	<p><img src="../../images/dataimporttxt2.jpg"></p>
	<p>
	<b>
	Text qualifier</b>:&nbsp;Import edilecek dosyadaki metinleri çevreleyen bir 
	karakter varsa bu seçilir, yoksa none bırakıldır. Önrneğin metin şu 
	formattaysa "başkent", "şube1", "2016", "532" text qualifer olarak çift 
	tırnak(") seçilir.</p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	<img src="../../images/dataimporttxt3.jpg"></p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Son aşamada ise almak iste<span style="text-decoration: underline">me</span>diğiniz 
	bir kolon varsa bunu işaretleyebilir, ayrıca kolonların veritipini de 
	belirleyebilirsiniz.</p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	<img src="../../images/dataimporttxt4.jpg"></p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Bir nedenden Excel, çıktıyı istediğiniz formata çevirmezse bile import 
	işleminden sonra da istediğiniz formata çevirebilirsiniz.</p>
	<p>Mesela text tipinde gelenleri sayıya çevirmek için Number tipini 
	uygulayabilirsiniz ancak bu data tipini Number yapmakla birlikte sağa dayalı 
	göstermez; bunun için
	<a href="DataMenusu_CesitliVeriislemearaclari.aspx#texttocolumn">Text to 
	Collumns</a> aracını kullanmanız gerekir.</p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Advanced seçeneği içinde de binlik ve ondalık ayraç seçimi ve negatif 
	sayılarla ilgili bir seçim yapılır.</p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	<img src="../../images/dataimporttxt5.jpg"></p>
	<h4 xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Fixed format</h4>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Fixed formatta seçim, aşağıdaki gibi her kolonun bitimine uygun çizgiler 
	koyarak yapılır.</p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	<img src="../../images/dataimportfixed.jpg"></p>
	<p xmlns:antixss="urn:AntiXSSExtensions" xmlns:friendlytitlelookup="urn:FriendlyTitleLookup">
	Diğer herşey Delimited format ile aynıdır.</p>
	<h3>XML</h3>
	<p>Xml dosya formatı platformlar arasında taşınabilen bir dosya formatıdır. 
	Bir çok dosya okuyucu XML'i okuyabilir. Excel de bunlardan biridir. 
	Okuyacağımız bu dosya yerel bir dosya olabileceği gibi internet üzerindeki bir xml dosya da 
	olabilir.</p>
	<p>Bununu için yine Legacy'den XML seçeriz. Kaynak olarak aşağıdaki sitenin 
	site haritasını girdim.</p>
	<p><img src="../../images/dataimportxml1.jpg"></p>
	<p>Bazı dosyalarda aşağıdaki gibi bir uyarı verir, OK diyip geçelim.</p>
	<p><img src="../../images/dataimportxml2.jpg"></p>
	<p>Import işlemini nereye yapacağınızı da seçtikten sonra işlem tamamdır.</p>
	<p><img src="../../images/dataimportxml4.jpg"></p>
	<p>Siz de benim sitem için deneyebilirsiniz:
	<a href="http://www.excelinefendisi.com/Sitemap.xml">
	http://www.excelinefendisi.com/Sitemap.xml</a> </p>
	<p>XML importu için daha detay bilgi için
	<a href="https://support.office.com/en-us/article/import-xml-data-6eca3906-d6c9-4f0d-b911-c736da817fa4">
	bu</a> ve
	<a href="https://support.office.com/en-us/article/overview-of-xml-in-excel-f11faa7e-63ae-4166-b3ac-c9e9752a7d80">
	şu</a> sayfalara bakabilirsiniz.</p>
	<p>Mevcuttaki bir excel tablosunu Xml olarak <strong>export </strong>etmek için 
	ise
	<a href="DeveloperMenusu_XML.aspx">buraya</a> bakınız.</p>
	</div>
	<h2 class="baslik">Web'den veri çekme</h2>
		<div class="konu">
	<p>Bu örnekte kendi web sitemin sitemap.aspx sayfasından veri çekeceğim. 
	Legacyden Web'i seçince aşağıdaki pencere açılır. Adres çubuğuna sayfa 
	adresini yazıp Go tuşuna basınca aşağıya sayfanın içeriği geldi. Bazı komut 
	dizesi hataları çıktı, bunlara ok diyip geçtim. Sonra kırmızı işaretli 
	yatay oka tıklayınca mavi çerçeve berlidi ve Excel bize o kısmı import 
	edeceğini söylemiş oldu. </p>
	<p><img src="../../images/dataimportweb1.jpg"></p>
	<p>Sonra import dedim ve sayfadaki importlanabilir veri Excele gelmiş oldu. 
	Gelmiş oldu ama istemediğim birçok veri de gelmiş oldu, ilk 55 satır benim 
	için çöp, bunları sildim ve istediğim data bana kalmış oldu.</p>
	<p><img src="../../images/dataimportweb2.jpg"></p>
	<p>Bu veri sağ tıklanarak refresh edilebilir durumdadır.</p>
	<p>Bununla beraber webden veri çekme, bu haliyle çok kullanışlı değildir. 
	Zira her sayfadaki veri bu yöntemle çekilmeye uygun olmayacaktır. Mesela siz 
	de <a href="https://kur.doviz.com/">https://kur.doviz.com/</a> sitesinden 
	veri çekmeye çalışın, çok nitelikli bir veri olmayacaktır. </p>
	<p>Web sitelerinden daha uygun bir veri çekmek için gelişmiş 
	progralama dillerini kullanabilir ve sadece birkaç satırlık kod ile şık formatlı veriler 
	çekebilirsiniz.&nbsp; Ancak koda Excel içnde ihtiyacınız varsa ve düzenli 
	olarak refreshlenebilir olmasını istiyorsanız VBA de kullanabilirsiniz, ancak 
	standart VBA kodunun ötesinde HTML ve Internet Explorer kütüphanelerini 
	kullanabiliyor olmalısınız, ve ayrıca biraz HTML ve Javascript bilgisi 
	fena olmayacaktır. </p>
			<p>
			<a href="../VBAMakro/DigerUygulamalarlailetisim_Webdenvericekme.aspx">
			Şu sayfada</a> konuyla ilgili bilgileri bulabilirsiniz.</p>
	</div>
	<h2 class="baslik">MS Query</h2>
		<div class="konu">
	<p>Evet, geldik eski zamanların en güçlü aracına, Power Query'nin öncülü MS 
	Query'ye. Bu araç ile birçok veri kaynağından ODBC bağlantısı kurarak veri 
	çekebiliyoruz.</p>
	<p>Bu araç Legacy içinde bulumuyor, Other sources içinde bulunuyor(farklı 
	Excel versiyonlarında yeri değişebilir, arayıp bulacağınızdan eminim)</p>
	<p><img src="../../images/datamsquery1.jpg"></p>
	<p>Biz veri kaynağı olarak yine Access seçelim ve sonrasında çıkan 
	pencereden ilgili Access dosyamızı seçelim.</p>
	<p><img src="../../images/datamsquery2.jpg"></p>
	<p>Bir Query Wizard çıkar ve hangi tablolardan hangi kolonları seçmemiz 
	gerektiğini bize sorar, ihtiyacımız olanları seçelim.</p>
	<p><img src="../../images/datamsquery3.jpg"></p>
	<p>Arkasından gelen kutudaki Filter ve Sortu şimdilik olduğu gibi 
	geçebiliriz, sonra son kutumuz çıkar.</p>
	<p><img src="../../images/datamsquery4.jpg"></p>
	<p>Return dersek direkt&nbsp;Excele atar, biz View diyelim ve Editörü açalım.</p>
	<p><img src="../../images/datamsquery5.jpg"></p>
	<p>Burada başkta tablolarla görsel join yapabilir, kriter koyabilir&nbsp; ve 
	hatta mevcut oluşan SQL'i manuel bir SQL ile değiştirebiliriz.</p>
	<p><img src="../../images/datamsquery6.jpg"></p>
	<p>Şimdi diyebilirsiniz ki, ben bu joini Access içinde kurup da yapabilirim, 
	sorguyu da Access içinde kaydeder ve direkt o sorguyu import ederim. 
	Haklsınız ancak bazı durumlarda ilgili databasede sorgu oluşturma ve 
	kaydetme hakkınız olmayabilir. İşte bu durumlar için MS Query oldukça 
	faydalıdır. Tabi PowerQuery'nin yanında MS Query'nin esamesi okunmaz ama 
	yine de öğrenelim, zira bir Oracle bağlantısı için 365li Excelin Home 
	versiyonunuda Power Query ile Oracle bağlantısını direkt yapamıyorsunuz, ya 
	professional versiyonunuz olmalı ya da 365siz Exceliniz. Ama MS Query hep 
	orada, o yüzden öğrenmekte fayda var.</p>
	<p>Oluşan SQL'i SQL butonundan görebilirsiniz. İsterseniz hazır SQL'i buraya direkt 
	yapıştırabilirsiniz veya bunu Excele attıktan sonra Properties'ten de 
	yapabilirsiniz.</p>
	<p><img src="../../images/datamsquery7.jpg"></p>
	<p>Son olarak nihai datayı Excele çıkmak için <strong>File&gt;Return Data do Excel
	</strong>deriz.</p>
	<p><strong>Oracle örneği üzerine not</strong>:Çalıştğım bilgisayarda 64 bit Windows 
	kullanıyorum, Office versiyonu ise 32 bit. Bilgisayara ise Windows uyumu 
	nedeniyle Oracle'ın 64 bitini kurdum. Office ile 
	Oracle'ın bit uyuşmazlığı nedeniyle oradan bir örnek yapamadım malesef ama Accesste 
	nasıl yapıyorsanız aynı mantık Oracle veya başka bir veri kayanğı için de 
	geçerli. Zaten bir kez bi Oracle veritabanına bağlantı kurduktan sonra artık 
	yeni bağlantılar için MS Queryde çalışmak yerine direkt Properties'ten SQL 
	metnini değiştirmeniz yeterli.</p>
	<h3>Query Editörü ve Parametreler</h3>
	<p>Çalıştırdığımız sorguyu değiştirmek için doğrudan SQL metnini 
	değiştirebileceğimiz gibi, MS Query editörüne geçip orada da işlem 
	yapabiliriz.</p>
			<p><strong>Properties&gt;Connection Properties</strong>'ten <strong>
			Edit Query</strong> dediğimizde editör ekranımız açılır. Burada 
			mesela aşağıdaki gibi bir filtre koyabiliriz. Filtre koymak için 
			editörün <strong>View </strong>menüsünden <strong>Criteria </strong>
			seçeneğini işaretleyin, aşağıdaki gibi kriter alanı açılacaktır. 
			Oraya [ ] içinde istediğiniz bilgiyi girip Entera basınca 
			sizdenilgili kriteri girmenizi isteyen bir kutu çıkacaktır.</p>
			<p>Böylece her refresh sırasında ürün bilgisini soran bu kutu 
			çıkacaktır.</p>
			<p><img src="../../images/datamsqueryprameter1.jpg"></p>
			<p>Dinamik bir sorgu için bu bir yöntemdir ancak daha şık bir 
			yöntem, olası seçenekleri bir hücreye
			<a href="DataMenusu_VeriDogrulama.aspx">Validation List</a> olarak 
			girip, oradan seçmektir.</p>
			<p>Bunun için Connection Properties'te <strong>Parameters</strong>'a<strong>
			</strong>tıklarız. Tabi bunu yapabilmek için hali hazırda bir kriter 
			uygulanmış olması lazım, yoksa bu düğme pasif gelecektir. Ancak biz 
			kriter uyguladığımız için aşağda gördüğüüz üzere bu düğme aktifir.</p>
			<p><img src="../../images/datamsqueryprameter2.jpg"></p>
			<p>Bu düğmeye tıklayalım. Aşağıdaki pencere gelecektir. İlk başta en 
			üstteki seçenek seçilidir. İkinci seçenekte sabit bir değer girilir 
			ki bence bu çok anlamsız bir seçenek, zira bunu gerek editör 
			ekranını kullanarak veya doğrudan SQL içinde kendimiz de 
			girebiliriz. Üçüncü seçenek ise bizim aradığımız seçenektir. Bununla 
			Excele "Kriteri şu hüreden al" demiş oluyoruz, aynı zamanda 
			altındaki seçeneği de işaretleriz ki her değişiklikten sonra bir de 
			manuel refresh yapmak zorunda kalmayalım, seçim yapılınca refresh de 
			otomatik olsun.</p>
			<p><img src="../../images/datamsqueryprameter3.jpg"></p>
			<p>G1 hücresinden Ürün3'ü seçince sonuç da böyle olur.</p>
			<p><img src="../../images/datamsqueryprameter4.jpg"></p>
	</div>
	<h2 class="baslik">Data Connection Wizard</h2>
		<div class="konu">
	<p> Dataya ulaşmak için yine Legacy'de bulunan Data Connection Wizard'ı da 
	kullanabilir.z Bu sihirbaz bize ODBC veya OLEDB başta olmak üzere çeşitli 
	bağlantılar kurmamızı sağlayabilir.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz6.jpg"></p>
	<p> Biz bunlardan ODBC ve en alttaki Other/Advanced'ı kullanıcaz.</p>
	<h4> ODBC</h4>
	<p> ODBC örneğinde de yine Access'e bağlanalım.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz7.jpg"></p>
	<p> Aynı dosyayı seçelim.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz8.jpg"></p>
	<p> Sonra tablo seçimini yapalım.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz9.jpg"></p>
	<p> İşlem tamamdır. Şimdi Properties'ten connection stringe bakalım. Gördüğünüz gibi bunda connection string ODBCdir(ODBC 
	ifadesini doğrudan 
	görmüyoruz ama DSN yazmasından bunun ODBC olduğunu anlıyoruz) ve sadece ODBC 
	bağlantılarda kullanılanbilen <strong>Edit Query </strong>butonu aktif durumdadır. 
	Buna basınca MS Query editörü açılacaktır.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz10.jpg"></p>
	<h4> Other/Advanced seçimi</h4>
	<p> Bu sefer karşımıza Oledb sağlayıcılar çıkar.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz11.jpg"></p>
	<p> Yine aynı Access dosyasına bağlanalım ve bağlantımızı test edelim.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz12.jpg"></p>
	<p> İlgili tabloları seçelim</p>
	<p> <img src="../../images/vbaimportsqldataconwiz13.jpg"></p>
	<p> Ve oluşan Connection stringe bir bakalım. Gördüğünüz gibi bunda 
	Connection String OLEDB'dir ve sadece ODBC 
	bağlantılarda kullanılanbilen <strong>Edit Query</strong> butonu pasif durumdadır.</p>
	<p> <img src="../../images/vbaimportsqldataconwiz14.jpg"></p>
	<p> &nbsp;</p>
</div>	
</asp:Content>
