<%@ Page Title='FormulasMenusu1 DiziFormulleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='7'></asp:Label></td></tr></table></div>

<h1>Dizi Formülleri ve Sumproduct</h1>
	<p>Buraya kadar konuları sırayla okuyarak geldiyseniz şimdi formüllerde 
	ustalaşma zamanınız geldi demektir. Çünkü burada artık olmayanı&nbsp;görme, 
	zihninizde canlandırma kabiliyeti kazanacaksınız.</p>
	<p>Bu formüllerle birkaç aşamada, genelde yardımcı kolonlar aracılığıyla, yaptığınız işlemleri tek bir formülle yapar 
	hale geleceksiniz. Bu sitede çeşitli yerlerde belirttiğim gibi, o anki 
	ihtiyaca göre bazen şık olmayan ama daha hızlı olan yöntemi seçmek gerekebilir. Yani 
	uzun uzun dizi formülü yazmak yerine geçici bir pivot table veya başka araçlarla daha hızlı 
	ilerleyeceğinizi düşünüyorsanız ve aciliyetiniz de varsa öyle ilerleyin. Ama 
	mesela kalıcı bir 
	dashboard/karne/scorecardv.s tasarlıyorasnız veya çıktı alınacak bir sayfa üzerinde 
	çalışıyorsanız veya bulduğunuz değeri başka bir formül içinde kullanacaksanız işinizi 
	tek seferde bitirmeye çalışmanız gerekecektir, ki bu durumların çoğunda bu 
	sayfada öğrendiklerinizi kullanabilirsiniz.</p>
	<p>Sayfa başlığında olmamakla birlikte <strong>Veritabanı fonskiyonlarını
	</strong>da konu 
	bütünlüğü açısından burada ele alıyor olacağız.</p>
	<p><strong>NOT</strong>:Örneklerin birçoğunda normal hücre 
	alanlarından(range) ziyade kolay anlaşılırlık adına <strong>
	<a href="HomeMenusu_Tablolar.aspx">Table</a> </strong>veya <strong>
	<a href="FormulasMenusuDiger_NameManager.aspx">Name</a></strong>'ler<strong>
	</strong>kulanılmış olup, bu konularda bilgi sahibi değilseniz öncelikler 
	bunları incelemenizi tavsiye ederim.&nbsp;</p>
	<p>Bu sayfadaki örneklerin hepsini bulacağınız örnek dosyayı 
	<a href="../../Ornek_dosyalar/Formuller/diziformulleri.xlsx"> 
	buradan</a> indirebilirsiniz. Ayrıca Microsoftun
	<a href="https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7">
	şu</a> sayfasında da dizilere ait detaylı bilgiler bulabilirsiniz, gerçi ben 
	oldukça detaylı bir içerik hazırladım ama yine de gözatmak isteyen meraklı 
	kullanıcılar bakabilirler.</p>

<h2 class='baslik'>Dizi Formülleri</h2>
<div class='konu'>
	<h3>
	Nedir?</h3>
	<p>
	Birden çok değerden oluşan 
	kümelere dizi, bu dizileri kullanan formüllere de dizi formülü diyoruz. </p>
	<p>
	Dizi formülleri ikiye ayrılmaktadır:</p>
	<ul>
		<li><strong>Tek </strong>hücreye girilen ve <strong>tek değer 
	döndüren </strong>formüller. Bunlar da kendi içinde üçe ayrılalır.
			<ul>
				<li>Doğal dizi formülleri(SUM, AVERAGE gibi dizi kabul eden ama 
				aslında dizi formülü olmayan formüller)</li>
				<li><strong>Control + Shift + Enter(CSE)</strong> tuş 
				kombinasyonuyla girilen dizi formülleri </li>
				<li><strong>SUMPRODUCT</strong> gibi parametre olarak her zaman birden çok dizi 
				alıp bunlarla işlem yapan veya <strong>INDEX, OFFSET </strong>gibi bazen dizi formülü 
				şeklinde davranan yerel fonksiyonlar. Bunlar CSE'ye ihtiyaç 
				duymazlar</li>
			</ul>
		</li>
		
		<li>Çoklu hücre grubuna girilip çok değer döndüren 
			fomrüller. Bunlar da kendi içinde birkaç gruba ayrılır		
				<ul>
					<li>Yerel Excel fonksiyonları(TRANSPOSE, LINEST v.s)</li>
					<li>Yine CSE kombinasyonuyla manuel girilen formüller</li>
				</ul>
		</li>
	</ul>
	<h3>Detaylar</h3>
	<p style="margin-bottom: 19px">Diziden kasıt, tek bir eleman değil de bir 
	grup elemandır. Mesela bir 
	formülün girdisi olarak 1 sayısını ele alalım. Bu tek bir değerdir. Ancak 
	bazen 1 ve 2 nin <strong>aynı anda </strong>formüle girmesini isteriz. İşte böyle bir durumda 
	dizi formülü kullanırız.</p>
	<p>Dizi elemanları formüle manuel olarak girilebileceği gibi, bunları 
	döndürecek başka bir formülle birlikte de girilebilir, veya belirli bir alan 
	seçilebilir. Son olarak formül girişi bitince <strong>CTRL+SHIFT+ENTER(CSE)</strong> 
	tuş kombinasyonuna basılır. Bu tuşlara basıldıktan sonra formül <strong>
	süslü parantezler yani "{ }"</strong> karakterleri arasına sarmalanır.
	<span style="color: red">
	<strong>Bu süslü parantezlerin elle giril<span style="text-decoration: underline">me</span>mesi çok önemli, yoksa hata 
	alırsınız. CSE'ye basınca kendiliğinden gelirler</strong>.</span></p>
	<p>Örnek girişleri aşağıdaki gibi gösterebiliriz</p>
	<ul>
		<li>Sabit giriş:{1;2;5;10} ({ } ile manuel girilenlerde CSE yapmaya 
		gerek yoktur)</li>
		<li>Hücre referansı:A2:A10</li>
		<li>Formül sonucu:TRANSPOSE(A2:A10)</li>
	</ul>
	<h3>1.Kullanım şekli:Çok sonuç döndürme</h3>
	<p>Çoklu sonuç döndürmek için belirli bir alanı seçip oraya tek bir formül 
	gireriz ve sonra CSE tuşuna basarız. Aşağıdaki görüntüde C2:C11 seçilmiş, 
	hepsine A2:A11*B2*B11 formülü girilip CSE tuşlarına basılmıştır. Böyle bir 
	kullanımda tüm hücrelerde aynı formül yer alır. Dizi formülü içieren 
	hücrelerden herhangi biri silinemez veya değiştirilemez, değiştirilmesi 
	teklif dahi edilemez. Silmek için hepsini birden seçip silmelisiniz.</p>
	<p><img src="/images/diziformul1.jpg" height="251" width="342"></p>
	<h4>Avantaj</h4>
	<p>Bu kullanım şeklinin en büyük avantajı, Excelin tek bir formülü hesaba 
	katmaya çalışmasıdır. Yukardaki örneğin alternatifi ne olurdu? C2 hücresine 
	A2*B2 girip onu aşağı doğru kaydırmak. Ama bu durumda 10 çeşit formül olurdu 
	ve Excel bu 10 formülü ayrı ayrı hesaplardı. Az sayıdaki satırlar için önemsiz 
	bir detay olabilir ama çok satırlı dosyalarda büyük performans katkısı sağlar. 
	Özellikle calculation işleminin geçici olarak durudulup tekrar açıldığı 
	durumlarda çok faydalı olabilir. Ayrıca diskte de daha az yer kaplar. 5000 Dikey eksenden oluşan bir kümeyi dizi formülü ile 
	kaydettiğimde 137 KB yer kaplarken, klasik formül 
	girip aşağı kaydırarak kaydedince 159 KB yer kaplamaktadır.</p>
	<p>NOT:Aşırı fazla dizi formülü ise tam tersi etki yaparak performans sorunu 
	yaratır. O yüzden Genel Değerlendirme bölümüde belirtildiği üzere dikkatli 
	kullanılmalıdırlar.</p>
	<h4>Transpoze işlemi</h4>
	<p>Dikey konumdaki A2:A10 arasındaki değerleri yatay şekilde(veya 
	yataydakileri dikey şekilde) bir yere yazdırma işlemi bir dizinin 
	transpozesini almak diye adlandırılır. Yatay veya dikey dönüşüm için 
	<strong>TRANSPOSE</strong> fonksiyonu kullanılır.</p>
	<p>Aşağıdaki J1:J3 arasındaki değerleri L1:N1 arasına girmek için L1:N1 
	seçilip ve aşağıdaki formül girilir.</p>
	<pre class="formul">
={TRANSPOSE(J1:J3)}</pre>
	<p><img src="/images/dizitranspose.jpg" height="101" width="333"></p>
	<h4>Formülü değiştirmek</h4>
	<p>Dizi formülleri değiştirilmek istendiğinde tüm blok seçilir, sonra F2 
	tuşuna basılarak ilk hücrenin içindeki formül değişim moduna getirilir, 
	değişiklik yapılır ve yine CSE kombinasyonuna basılır.</p>
	<h3>2.kullanım şekli:Tek hücre girişi</h3>
	<p>Benim dizi formüllerinde ağırlık vermek istediğim versiyon aslında bu versiyondur. Önce bu kullanım şeklinin neyin alternatifi olduğunu görelim. </p>
	<p>Yukardaki örneği düşünün. Çarpım 
	kolonundaki rakamların toplanmasını isteseydik, bunları toplamamız 
	gerekirdi. Yani önce C kolonunda A ve B'yi çarpıp sonra da bunları 
	toplayarak 2 iş yapardık. Bunun yerine şu formül işimizi görecektir.</p>
	<pre class="formul">{=SUM(A2:A11*B2:B11)}</pre>
	<p>Aslında bu işlem için başka bir fonksiyon var: SUMPRODUCT, bunu aşağıda 
	ayrıca 
	göreceğiz.</p>
	<p>Başka bir örneği inceleyelim. Yine Dizi formüllerinin olmadığı 
	dünyadayız, şimdi yukarıdaki rakamların yanına bir de fark kolonu ekleyelim 
	ve bunlardan en küçüğünü MIN(D2:D12) ile bulalım. Sonuç: -140</p>
	<p>
	<img src="/images/diziformul2.jpg" height="252" width="287"></p>
	<p>Dizi formüllerinin olduğu bir dünyada ise bunu şu dizi formülü ile elde ederiz.</p>
	<pre class="formul">{=MIN(B2:B11-A2:A11)}</pre>
	<h3>F9 tuşu</h3>
	<p>F9 tuşu ile formül editleme modundayken formülün bir kısmını seçip 
	sadece o kısmın sonucunu görebiliyoruz. Bu tuşu daha çok uzun formüllerde 
	belli bir kısmın ara sonucunu görmek için kullanırız. Aynı mantıkla dizi formüllerinde 
	de belli bir kısmın döndürdüğü diziyi görmek 
	adına da kullandığımızda oldukça faydalı olmaktadır. Özellikle dizi 
	formüllerini yeni öğrenirken bunu sık sık kullanıp eldeki dizinin neye 
	benzediğini görmeniz açısından kritik bir ihtiyaçtır.</p>
	<p>Mesela şu formüldeki ilgili kısmı seçip,</p>
	<p>
	<img src="/images/diziformul3.jpg"></p>
	<p>F9'a basınca	sonuç aşağıdaki gibi olmaktadır</p>
	<img src="/images/diziformul4.jpg" height="34" width="376">
	<h3>VE/VEYA operatörleri ile&nbsp; Boolean(TRUE/FALSE) işlemler</h3>
	<h4>Konuşma dili vs Formül dili</h4>
	<p>Konuşma dili ile kasttetiğimiz şey her zaman formül diline aynı şekilde 
	girilmediği için bu farktan bahsetmek istedim. Eğer hali hazırda SQL ile 
	veya Business Objects gibi bir raporlama aracıyla rapor çekiyorsanız kriter 
	alanına değerleri girerken bu farka dikkat etmiş olmalısınız.</p>
	<p>Şimdi, formül dilindeki <strong>VE</strong>, konuşma dilinde iyelik eki olarak düşünülmelidir. 
	Konuşma dlindeki "Akdeniz bölgesi<strong>nin</strong> 2010 satışları" 
	formülde "Bölge=Akdeniz <strong>VE</strong> 
	Yıl=2010" olarak işleme girilir.</p>
	<p>Keza formül dilindeki <strong>VEYA </strong>ise konuşma dilinde garip bir şekilde 
	"ve" olarak kullanlmaktadır. O yüzden burdaki karışıklığa özel dikkat etmeniz 
	gerekir. Mesela konuşma dilince "2010 <strong>ve</strong> 2011" satışları 
	derken formüle "Yıl=2010 <strong>VEYA </strong>Yıl=2011" satışları olarak girilmelidir.</p>
	<p>İkisinin karışımına da örnek verebiliriz. Akdeniz bölgesi<strong>nin</strong> 2010 
	<strong>ve </strong>2011 
	satışları dersek, Bölge=Akdeniz <strong>VE</strong> (Yıl=2011 <strong>
	VEYA</strong> Yıl=2012) şeklinde işleme girecektir.</p>
	<h4>VE/VEYA operatörleri</h4>
	<p>Yukarda anlatılan VE işlemini parantezler arasında "*" işareti ile 
	sağlarken VEYA işlemini "<strong>+</strong>" işareti ile sağlarız. Bu şekilde bu dizide 
	ilgili koşulları sağlayan kaç eleman olduğunu bulmuş oluruz. Eğer formül 
	içinde çarpılıp toplanacak sayısal bir blok yok ise dizi formülleri bu 
	haliyle <strong>COUNTIF(S)</strong> alternatifi olarak karşımıza çıkar. Sayısal blok da işleme 
	girerse o zaman da <strong>SUMIF(S)</strong> alternatifi olur.</p>
	<p>Şimdi aşağıdaki örnek üzerinden gidecek olursak(Örnek tabloda Ürün3 de 
	vardır, ancak resim uzun çıkmasın diye biraz kırptım);</p>
	<p><img src="../../images/diziformultruefalse.jpg"></p>
	<p>Önce <strong>VE </strong>örneğine bakalım: <strong>Bölge1'in şube sayısı:</strong></p>
	<pre class="formul">=SUM((bölgeler="Bölge1")*(ürünler="Ürün1"))
//veya aynı formül sum+if ile aşağıdaki gibi iyazılabilir
=SUM(IF((bölgeler="Bölge1")*(ürünler="Ürün1");1))</pre>
	<p>
	<strong>Bölge1in Ürün1deki toplamı</strong></p>
	<pre class="formul">=SUM((bölgeler="Bölge1")*(ürünler="Ürün1")*satışlar)</pre>
	<p>
	<strong>Bölge1in, Ürün1 ve 2 toplamı ise</strong></p>
	<pre class="formul">=SUM((bölgeler="Bölge1")*((ürünler="Ürün1")+(ürünler="Ürün2"))*satışlar)</pre>
	<p>
	Bu örneklerde sondaki "satışlar" olmazsa COUNTIFS, olursa SUMIFS işlemi 
	yapılmaktadır.</p>
	<h4>
	Nasıl işliyor?</h4>
	<p>
	Aritmetik işlemlerde BOOLEAN değer(yani TRUE/FALSE) döndüren sonuçlar 
	sırayla 1/0 değerlerine dönüştürülür. Bilindiği gibi 0'la çarpım hep 0'dır, 1'le çarpımlar da 
	çarpılan değerlerin toplamını verir. Eğer çarpılanlar yine 0/1den oluşan 
	dizilerse sonucun COUNTIFS, sayısal blok çarpılıyorsa da SUMIFS olduğunu 
	belirtmiştik.</p>
	<p>
	Mesela yukardaki son örnekte "bölgeler=Bölge1" eşitliğini F9 ile çözümlemeye 
	çalıştığımızda aşağıdaki gibi görürüz.</p>
	<p>
	<img src="../../images/diziformultrue2.jpg"></p>
	<p>
	Bunların sayısal karşılığı 1/0 ama bunları şu 
	anda F9 ile göremiyoruz. Bunları 1/0 görmek için biraz aşağıdaki teknikleri uygulamanız lazım gelir. İlgili 
	tekniklerden biri uygulandığında çıkan sonuç aşağıdaki gibidir.</p>
	<p>
	{<span style="color: red">1</span>;1;1;1;1;1;0;0;0;0;1;1;1;1;1;1;0;0;0;0;1;1;1;1;1;1;0;0;0;0}</p>
	<p>
	Ürünlü kısmın 0/1 karşılığı ise şöyledir:<br>
	{<span style="color: red">1</span>;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;0;0;0;0;0;0;0;0;0;0}</p>
	<p>
	Satışların ise zaten direkt kendisi gelir:<br>
	{<span style="color: red">101</span>;155;224;229;108;433;213;337;341;324;387;112;143;247;146;174;412;439 ;204;370;327;212;329;232;341;267;268;355;252;386}</p>
	<p>
	Son olarak da tüm bunlar birbiriyle çarpılır, çarpımda 0 olan kısımlar 
	elenir, geriye sadece 1*1*SatışRakamı olan kısımlar kalır ve bunlar 
	toplanır. Örneğin ilk grup <span style="color: red">1*1*101</span> şeklinde 
	hesaba girer.</p>
	<h4>TRUE/FALSE'ı sayısal değere çevirme yöntemleri</h4>
	<p>TRUE ve FALSE'ın sayısal karşılığının 1 ve 0 olduğunu görmüştük. Bunu 
	yukarda doğrudan koşullu ifadelerin parantezler içinde çarpımı örneğinde gördük. Çarpımın olmadığı 
	ve parametre ayracı olan ";"&nbsp; ile kullanıldığı durumlarda ise aşağıdaki yöntemler uygulanır.</p>
	<p><img src="../../images/diziformul9.jpg"></p>
	<p>İlk yöntemde -- ifadesi aslında iki kere -1le çarpım anlamına 
	gelmektedir, yani ifadeyi aritmetiksel bir işleme tabi tutunca 1/0 haline 
	dönüşmektedir. 3. ve 4. yöntemde de benzer mantık bulunmaktadır. 2. yöntem 
	ise direkt bu amaçla üretilmiş bir fonksiyon olup TRUE/FALSE'ları 1/0'a 
	çevirir.</p>
	<p>Şimdi aşağıdaki tabloya bakalım.</p>
	<p><img src="/images/diziformul10.jpg" height="214" width="156"></p>
	<p>18'nden büyük olan kaç kişi var diye bakacağız.</p>
	<pre class="formul">=SUM(--(G2:G9&gt;18))</pre>
	<p><img src="/images/diziformul11.jpg" height="34" width="392"></p>
	<p>--'yi de kapsayacak şekilde seçip F9 yaparsak 1/0 karşılıkları görünür.</p>
	<p><img src="/images/diziformul12.jpg" height="32" width="160"></p>
	<p>Sonuç 5 olacaktır.</p>
	<h4>IF'li formüller</h4>
	<p>Aşağıdaki tabloya göre Mağaza1'in toplam satış tutarını bulmak <strong>
	SUMIF </strong>fonksiyonu ile oldukça kolaydır. SUMIF'in olmadığı bir 
	dünyada(gerçi hemen herkesteki Excel versiyonu bu formülü destekler diye 
	düşünüyorum) ise bunun alternatifi nasıl yazılır ona bakalım.</p>
	<p><img src="../../images/diziformul40.jpg"></p>
	<pre class="formul">{=SUM(IF(Table2[Mağaza]=F3;Table2[Tutar]))}</pre>
	<p>Gördüğünüz gibi SUMIF'e oldukça benzemekte. Evet, artık herkeste SUMIF 
	vardır diyoruz ama peki ya <strong>MINIF</strong>? Böyle bir fonksiyon 
	2016'ya kadar gelmedi. 2016'daki de MINIF değil çoklu kriteri desteklemesi 
	adına <strong>MINIFS </strong>olarak geldi(Office 365 paketi değilse Excel 
	2016 olsa bile bu formül desteklenmez)</p>
	<p>O halde tek alternatifimiz dizi formülü yazmak olacaktır.</p>
	<pre class="formul">=MIN(IF(Table2[Mağaza]=F3;Table2[Tutar]))</pre>
	<p>IF'li kısım bize 
	{218;299;180;474;413;478;306;298;310;487;117;472;FALSE;FALSE;FALSE;FALSE;FALSE;FALSE} 
	dizisini döndürür, sonra MIN ile de bunlardan en küçüğünü elde ederiz.</p>
	<p>Ekli dosyada MAXIF, SMALLIF MINIFS alternatifleri de bulunmaktadır.</p>
	<h3>Ardışık sayı üretmek</h3>
	<p>Özellikle SMALL ve LARGE başta olmak üzere bazı fonksiyonlarda 1 ile X arası 
	ardışık sayıları kullanmak gerekebilmektedir. Bunlar az ise elle {1;2} gibi 
	girilebilir. Ancak çoksa veya sayısı baştan bilinmeyip başka hücrelerdeki 
	değerlere göre dinamik olarak belirleniyorsa başka bir çözüm bulmamız gerekir.</p>
	<p>Bu amaç için <strong>ROW </strong>fonksiyonunu dizi şeklinde kullanırız. Mesela aşağıdaki 
	formülü K8:K10 arasına girip CSE yaparsak bize sırayla 1,2,3 değerlerini 
	verir.</p>
	<pre class="formul">{=ROW(1:3)}</pre>
	<p>Ancak bu yöntemin bi sakıncısı var, o da A1:A3 arasına bir satır açılırsa 
	bu formül sapıtır. Kendiiniz de deneyip görebilirsiniz. Mesela en tepeye, 1.satırın üstüne 
	yeni kolon açılırsa az önce 1,2,3 gördüğümüz değerle3 2,3,4 olur.&nbsp; </p>
	<p>İşte bu sapma olmasın diye <strong>INDIRECT </strong>fonksiyonunu da formüle dahil ederiz.</p>
	<pre class="formul">=ROW(INDIRECT(1&amp;":"&amp;3))</pre>
	<p>Bunun kullanımıyla ilgili bir örneği aşağıda çeşitli örneklerin bulunduğu 
	kısımda görebilirsiniz.</p>
	<p>NOT:Eğer sonucu dikey eksene yazdıracaksanız veya dikey eksendeki bir 
	bölgeye kriter olarak sokacaksanız <strong>ROW</strong>, yatay eksende 
	kullanacaksanı <strong>COLUMN</strong> kullanırsınız.</p>
	<h3>Rakamların yönü</h3>
	<p>Bu ön bilgiyi şimdi veriyorum, ancak şuan için birşeyi fade etmemesi çok 
	yüksek. Sadece şağıdaki örnekeri görürken, niye burda böyle de şurda şöyle 
	diye soracağınızı tahmi nettiğim için şimdiden vermek isteim. O örneklere 
	geldiğinizde daha iyi bir kavrayış iin bu kısma tekrar gelip bakın.</p>
	<ul>
		<li>Çoklu değerler girilirken manuel girilirlerse { } işaretleri arasına 
		";" ayracıyla ayrılarak girilirler. 
		SUM(SUMIF(Tablo[Bölge];{"Bölge1";"Bölge2"};Tablo[Satış]))</li>
		<li>Eğer giriş manuel değil de bir hücre grubundan alınacaksa<ul>
			<li>İfadeler ";" parametre ayracıyla giriliyorsa hücre grubunun yönü 
			önemli değildir.<br>
			Yatay:=SUM(SUMIF(Tablo[Bölge];G49:H49;Tablo[Satış]))<br>
			Dikey:=SUM(SUMIF(Tablo[Bölge];G51:G52;Tablo[Satış]))</li>
			<li>İfadeler "*" işaretiyle çarpılacaksa hücre grubunun yatay olması 
			gerekir. Dikey hücre grupları TRANSPOSE edilmelidir.<br>
			Yatay:=SUMPRODUCT((Tablo[Bölge]=G49:H49)*(Tablo[Satış]))ış]))<br>Dikey:=SUMPRODUCT((Tablo[Bölge]=TRANSPOSE(G51:G52))*(Tablo[Satış]))</li>
		</ul>
		</li>
	</ul>
	</div>
	<h2 id="dinamikdizi" class="baslik">Dinamik Diziler(Yeni!!!)</h2>
	<div class="konu">
		<p>Eylül 2018'de Micorosft efsane bi işe imza attı. Hesaplama motorunu değiştirdi. Ve Excel kullanıcılarının hayatını kökten değiştirecek bir yapıyı hizmete soktu. <strong>Dinamik Diziler(Dynamic arrays)</strong></p>
		<p>Öncelikle şunu belirteyim, bu özellikten sadece Office 365&#39;i olan kullanıcılar faydalanabiliyor. O yüzden maalesef ben dahi bunu 2 yıl sonra farkedebildim. Zira iş yerimdeki PC&#39;mde Office 365 kurulu değil.</p>
        <p>Peki nedir bu dinamik diziler, ne işe yarar? Aslında bizi iki şeyden kurtarıyorlar: Dizi formülü yazarken hedef alanı uygun sayıda hücre olacak kadar seçmekten ve <strong>CSE</strong> kombinasyonuna basmaktan. Siz formülü yazdığınızda formül otomatikman aşağı &quot;<strong>dökülüyor</strong>&quot;. Evet buna dökülmek(spill) diyorlar. <strong>FILTER, SORT, UNIQUE</strong> gibi efsane fonksiyonlar da cabası. Aşağıda bunların nasıl çalıştığını gif animasyon olarak görebilirsiniz. Ben örnek olarak UNIQUE&#39;i aldım. Siz de diğerlerini deneyebilirsiniz. Şu an çok detaylarına girecek vaktim yok ancak bi ara hepsi için detaylı örnekler koymayı düşünüyorum. O zamana kadar Microsoft&#39;un kendi <a href="https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531">sayfasından</a> ve/veya diğer kaynaklardan faydalanabilirsiniz.</p>
        <p>
            <img alt="dinamik diizi" src="../../images/dinamikdiziunique.gif" /></p>
	</div>

	<h2 id="sumproduct" class="baslik">SUMPRODUCT</h2>
	<div class="konu">
	<p><span class="keywordler">SUMPRODUCT </span>size ilk başta&nbsp;o kadar da ahım şahım bir fonksiyon gibi 
	gelmeyebilir, ama Excel'in dizilerle nasıl çalıştığını anladıktan sonra 
	bu fonksiyonun da hakkını vermeye başlıyorsunuz.</p>
	<p>Öncelikle SUMPRODUCT kendi başına bir dizi formülüdür. Yani bir formül 
	oluşturduktan sonra CSE kombinasyonuna basmanız gerekmez. Aslında tek bir 
	kullanım şekli olmakla birlikte ben tamamen kendi kafama göre 2 gruba ayırdım; Klasik ve klasik olmayan 
	kullanımı diye.</p>
	<h3>Klasik kullanım</h3>
	<p>Diyelim aşağıdaki gibi Kredi tutarı ve faiz oranlarından oluşan bir liste 
	var, ve siz <strong>ağırlıklı ortalama </strong>faizi bulmak istiyorsunuz. Aritmetik ortalama 
	alırsanız hata yaparsınız, zira içerde büyük hacimli ama düşük faizli bir 
	kredi varsa genel toplamı aşağı çekmesi gerekir. Fakat aritmetik ortalama, hacimlerin 
	büyüklüğünü 
	dikkate almadığı için yanıltıcı olacaktır. O yüzden aritmetik ortalama 
	yerine ağırlıklı ortalama alınması gerekmektedir.</p>
	<p><img src="../../images/dizisumproduct1.jpg"></p>
	<p>Ağırlık ortalama için SUMPRODUCT'ın olmadığı dünyada D kolonunda yardımcı 
	bir kolon açılıp Faiz*Tutar çarpımı bulunur, Sonra D kolonunun toplamı B kolonunun 
	toplamına bölünür.&nbsp;&nbsp;&nbsp; </p>
	<pre class="formul">=SUM(D2:D11)/SUM(B2:B11)</pre>
	<p>İşte SUMPRODUCT sizi bu yardımcı kolondan kurtarır.</p>
	<pre class="formul">=SUMPRODUCT(B2:B11;C2:C11)/SUM(B2:B11)</pre>
	<p>Fonskiyonun yaptığı iş bu dizileri çarpıp toplamaktır. Yukarıdaki 
	örnekte her satırdaki Faiz ve Tutarı çarpıyor ve bütün bu çarpımları 
	topluyor.</p>
		<p>Gördüğünüz üzere SUMPRODUCT parametre olarak diziler alıyor ve tüm diğer 
	fonksiyonlarda olduğu gibi parametre ayracı olarak ";" işaretini kullanıyor. 
	Ancak SUMPRODUCT'ı <strong>VE</strong> operatörü olan "*" ile de kullanabiliyoruz. 
		Ve operatörü kullanıyor olmak demek, koşul arıyouz demektir, yani diziyi 
		formüle olduğu gibi sokmak yerine belli koşulda olanları istiyoruz 
		demektir. Gerçi "*" ile kullanıldığında koşul belirtilmese de yine 
		çalışır ama bu şekilde anlamsızdır, zira bunun için ";" versiyonu zaten 
		vardır. Bu arada ";" ile kullanıdığında diziler parantez içine alınmazken "*" ile 
		kullanıldığında parantez içine alınırlar.</p>
	<pre class="formul">=SUMPRODUCT((B2:B11)*(C2:C11))/SUM(B2:B11)</pre>
	<h3>Klasik olmayan kullanım şekil</h3>
		<p>Şimdi fonksiyonumuz buraya kadar gereksiz görünebilir, zira "*" versiyonunun 
		yaptığı işi SUM 
	fonksiyonunu dizi formülü olarak kullanarak da gerçekleştirebilirdik. Hatta yukarıdaki 
		örneklerinden birinde yapmıştık. Bu örnekte kullanımı aşağıdaki gibi 
		olacaktır.:</p>
	<pre class="formul">{=SUM((B2:B11)*(C2:C11))/SUM(B2:B11)}</pre>
	<p><strong>NOT</strong>:Buradan itibaren aşağıda görülen SUMPRODUCT'ın "*" 
	verisyonlarının tamamı 
	SUM-Dizi formülü şeklinde de yapılabilir olup, tekrardan kaçınmak için 
	alternatifler arasında bu ayrıca gösterilmeyecektir.</p>
	<p>Hele hele <strong>SUMIF, COUNTIF </strong>ve bunların "S" ile biten çoklu 
	koşul türevleri çoğu durumda 
	SUM-Dizi formüllerini gereksiz bıraktığı gibi SUMPRODUCT'ı da gereksiz bırakmış 
	gibi düşünülebilir. Ancak yukarda belirttiğim gibi Excel'in dizilerle çalışma şeklini 
	ve SUMPRODUCT'ın kullanım şeklini anladıktan sonra bunun ne kadar gerekli bir fonksiyon olduğunu 
	göreceksiniz.</p>
	<p><span class="dikkat">DİKKAT</span>:SUMPRODUCT'ı tüm kolon seçimlerinde(Ör:A:A) 
	kullanmamaya çalışın, çoğunlukla hata döndürür. Sadece ilgili alanda(ÖR:A2:A20) 
	kullanın.</p>
	</div>
	<h2 class="baslik">Veritabanı(Database) fonksiyonları</h2>
	<div class="konu">
	<p>Şimdi dizi formüllerine kısa bir ara veriyoruz ve birçok durumda gerek 
	dizi formüllerine gerek SUMPRODUCT'a alternatif olmaları sebebiyle <strong>Veritabanı fonksiyonlarına 
	</strong>bakıyoruz. Bu konudan sonra hepsini birden tekrar ele alacağız ve en son da bir 
	karşılaştırma yapacağız.</p>
		<p>Veritabanı fonksiyonları, bir liste üzerinde çalışırken, <strong>birden çok 
		kritere </strong>göre toplam/adet/min/max aldırma gibi istatistiki işlemler 
		yapmayı sağlar.</p>
		<p>Genel syntax'ı şöyledir: <strong>Fonksiyon(veritabanı,&nbsp;field(işlem 
		yapılacak alan),&nbsp;kriter alanı)</strong></p>
		<p><strong>Veritabanı </strong>ifadesi gözünüzü korkutmasın. Bu aslında herhangi bir 
		hücre grubudur veya Table/Name şeklindeki özel bir hücre grubudur. 
		Veritabanı olarak gösterilen alanın ilk satırının başlık olması gerekir.</p>
		<p><strong>İşlem yapılacak alan</strong>: İlgili kolonun veritabanında kaçıncı kolon 
		olduğu veya doğrudan onun başlığıdır.(Veritabanı A kolonundan başlamıyorsa 
		mesela D kolonunda başlıyorsa ve işlem alanı G kolonundaysa 7 değil 4 
		girilir. DEF<strong>G</strong>HIJKLM)</p>
		<p><strong>Kriter </strong>alanında ne konacağını tahmin ediyorsunuz:Kriterler. Kriterler 
		tıpkı <strong>Advanced Filter </strong>kurallarında olduu gibi girilir. 
		Aynı satırdakiler <strong>VE</strong> olarak, farklı satırdakiler 
		<strong>VEYA</strong> olarak algılanır. Kriter alanı birbirine komşu 
		hücrelerden oluşmalıdır.</p>
		<p><strong>NOT</strong>:Bu fonksiyonlar joker elemanları destekler. Herhangi bir karakter için 
		<strong>"?"</strong>, birden çok karakter için <strong>"*"</strong> karakteri kullanılır. Ayrıca 
		case-sensitive değildirler.</p>
		<p>Şimdi aşağıdaki tablo üzerinden bir örnek yapalım.</p>
		<p><img src="../../images/dizidsum1.jpg"></p>
		<p>Kriter alanımız da aşağıdaki gibi olsun. Bu 3 hücre de combobox 
		şeklindedir ve olası değerleri içermektedir. (Kriterler sadece I ve J 
		kolonunda. K kolonundaki bilgiyi formü içinde ayrıca kullanacağız)</p>
		<p><img src="../../images/dizidsum2.jpg"></p>
		<p>Aradığımız şey: Marmara1 bölgesinin 2013 satışları toplamı olsun.</p>
		<p>Formülümüz şöyledir.</p>
		<pre class="formul">=DSUM(A1:F19;MATCH(K2;A1:F1;0);I1:J2)</pre>
		<p>Farkettiyseniz <strong>Field </strong>parametresini dinamik bir 
		şekilde hesapladık, yani bunu 
		kullanıcıya manuel girdirtmek yerine K2 hücresindeki combobox'tan 
		seçtirdik.</p>
		<p>Bu arada bu fonksiyonların joker karakterleri de desteklediğini söylemiştik. Kriter 
		alanını aşağıdaki gibi girersek, Marmara ile başlayan tüm bölgeleri 
		işleme sokar.</p>
		<p><img src="../../images/dizidsum3.jpg"></p>
		<p>Diğer önemli Veritabanı fonksiyonları <strong>DCOUNT, DAVERAGE, DMIN 
		VE DMAX</strong>'tır. Bunların D'siz versiyonlarını biliyorsanız ne 
		işe yaradıklarını&nbsp; anlamışsınızdır, bilmiyorsanız önce onlara 
		gözatmanızı tavsiye ederim.&nbsp;Tekrardan kaçınmak adına bunlarla 
		ilgili örnek yapmaya gerek duymuyorum.</p>
	</div>
	<h2 class="baslik">1 Boyutlu tablolarda çalışmak</h2>
	<div class="konu">
	<p>Aşağıdaki gibi matrisyel formda ama yatay ekseni <strong>tek </strong>kolondan oluşan tablolarla 
	çalışmak oldukça kolaydır. Bunlarda dikey eksen bir kolondan da birkaç 
	kolondan da oluşabilir. </p>
		<p>Bunlara, yatay eksende 1 kolon olduğu için 1 boyutlu tablolar 
	diyoruz. Aslında bunu sadece ben diyorum. MSDN dökümantasyonunda yer alan 
		bir sınıflandırma değil yani. Şimdi bu tür tablolarda ne tür işlemler 
	yapabiliyoruz, karşılaştırmalı olarak göreceğiz.</p>
	<p><img src="../../images/diziformul20.jpg"></p>
		<h3>Önemli Not (2 boyutlu tablolar için de geçerlidir)</h3>
		<p>Aşağıda göreceğiniz gibi, elde edilmeye çalışılan şey ve bunlara 
		ulaşılan yöntem sayısı çok fazla olabilir. Ben elimden geldiğince çok 
		farklı kombinasyonlara değinmeye çalıştım. Ör: Yöntem olarak SUMPRODUCT'ı 
		vermişimdir ama bunun ";" versiyonu da "*" versiyonu da var, hatta bir 
		de VEYA 
		kriteri için "+" operatörlü kullanımı var. Ayrıca çoklu kriterlerin 
		yatay eksenden oluşan bir alandan mı(ör:J1:L1), dikey eksenden oluşan 
		bir alandan mı(ör:J1:J3), yoksa formüle manuel mi(ör:{1;2;3}) girilmesi de 
		kombinasyon sayısını artırmaktadır. </p>
		<p>O yüzden tüm kombinasyonları buraya almak çok mantıklı olmayacaktır. 
		Bunun yerine yöntemlerde çeşitlendirme yapmayı tercih ettim. Mesela 
		SUMPRODUCT'ın çoklu kriter örneğini manuel girerken, başka bir fonksiyon 
		alternatifinde yatay eksenden aldım, başkasında ise dikeyden aldım, böylece gereksiz 
		tekrardan kaçınmaya çalıştım. Gerçi buradaki örneklerin bulunduğu<a href="../../Ornek_dosyalar/Formuller/diziformulleri.xlsx"> excel 
		dosyayı</a> indiriseniz orada bu tekrarların bazısını görebilirsiniz, ama 
		orada bile tüm kombinasyonlar(tahminimce 100ün üzerinde) yok. Bu sayfayı tamamen hatmettikten sonra 
		gereken kombinasyonları siz duruma göre oluşturuyor olabilmelisiniz.</p>
		<p>Ayrıca yukarda belirtildiği gibi SUMPRODUCT'ın alternatifi olarak SUM'ın 
		"*"'lı dizi formülünü ise kullanmaktan tamamen kaçındım, örnek dosyada 
		da bulamayacaksınız. Yukarıda sadece bir kere verdim. SUMPRODUCT'lı her case için 
		teorik olarak onu da 
		kullanabileceğinizi bilmenizi isterim(ama SUMPRODUCT varken çok da gerek 
		yok açıkçası)</p>
	<p>Son olark burada yapacağımız neredeyse tüm örneklerde kriterler <strong>
	<a href="DataMenusu_VeriDogrulama.aspx">Data Validation</a></strong>'la 
	yapılmış comboboxlar içinden seçiliyor olacakır. Zira kriter değiştirmek 
	istediğimizde formülün içinde bunu değişitrmek yerine kriterin dinamik 
	olarak bir hücreden beslenmesi daha profesyonelcedir.&nbsp; Bununla beraber 
	bazı örneklerde özellikle çoklu değer girilmesi gereken durumlarda manuel 
	kriter girişi de yapılacaktır.&nbsp;</p>
		<h3>A)Dikey eksendeki kolon(lar)da aranan şey 1 adetse</h3>
	<h4>1)Dikey eksende tek kolon var</h4>
	<p>Eğer, aranan değer bir kez geçiyorsa <strong>VLOOKUP</strong> veya 
	<strong>INDEX/MATCH</strong> ile arana 
	değer bulunurken, birden çok kez geçiyorsa <strong>SUMIF</strong> kullanılarak eşleşenlerin 
	toplamı alınır. Bunlarla ilgili örnekler
	<a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx">Lookup 
	fonksiyonları</a> ve
	<a href="FormulasMenusuFonksiyonlar_IstatistikiveMatematikselFonksiyonlar.aspx">
	istatistiki fonksyionlar</a> sayfalarında fazlasıyla yapılmıştır.</p>
	<h4>2)Dikey eksende çok kolon var</h4>
	<p>Dikey eksendeki aranan değerler bir kez geçiyorsa INDEX-MATCH'in dizi 
	formülü olmak zorundadır. Bunu da yine
	<a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx">Lookup 
	fonksiyonları</a> sayfasında çok kriterli vlookup olarak görmüştük.</p>
	<p>Eğer aranan kriterler çok kez geçiyorsa o zaman birkaç alternatifimiz 
	var. Aşağıdaki örnek üzerinden gidelim.</p>
	<p><img src="../../images/diziformul21.jpg"></p>
	<p>Bölge2'nin(F33 hücresinde) 2.dönem(G33 hücresinde) satışlarını 
	toplayalım. Formül anlaşılırlığı açısından aşağıda göreceğiniz üzere 
	sırasıyla şu <strong>Name </strong>tanımlamalarını yaptım: bölgeler, dönemler 
	ve tutarlar.</p>
	<pre class="formul">=SUMIFS(Tutarlar;bölgeler2;F33;dönemler;G33)  //2007 sonrası için SUMIFS
{=SUM(IF((bölgeler2=F33)*(dönemler=G33);Tutarlar))} //SUM+IF dizi formülü
=SUMPRODUCT(--(bölgeler2=F33);N(dönemler=G33);Tutarlar) //SUMPRODUCT ";" ile
=SUMPRODUCT((bölgeler2=F33)*(dönemler=G33)*Tutarlar) //SUMPRODUCT "*" ile
=DSUM(alanfordsum;3;F32:G33)</pre>
	<h3>B)Dikey eksendeki kolon(lar)da aranan şey 1'den çoksa</h3>
		<p>Dikey eksende birden çok şey aramak, konuşma dilinde "ve", formül 
		dilinde ise <strong>VEYA </strong>şeklinde algılanır ve eğer ayrı ayrı formüle girecekse 
		"+" işaretiyle girer. Ayrı ayrı girmeyip dizi(manuel veya hüre gurubundan) 
		olarak da girebilir; o zaman fakrlı yöntemler işletilir.</p>
		<h4>1)Dikey eksende tek kolon var</h4>
		<p>Aşağıdaki tablo üzerinden gidecek olursak, öncelikle dikey eksende iki 
		kolon olduğunu görüyorsunuz ancak biz tek kolon üzerinden işlem 
		yapacağımız için bunu tek kolonlu bir dikey eksen gibi düşünebilirsiniz.</p>
		<p><img src="../../images/diziformul25.jpg"></p>
		<p>Şimdi böyle bir tabloda Bölge1 ve Bölge2nin satılarını toplamak 
		istiyoruz. Birçok alternatif olmakla birlikte biz burada birkaçını 
		vereceğiz, kalanına Excel dosya içinden bakabilirsiniz.</p>
		<pre class="formul">=SUM(SUMIF(Tablo[Bölge];{"Bölge1";"Bölge2"};Tablo[Satış])) //Bölge kriteri manuel
=SUMPRODUCT((Tablo[Bölge]=G49:H49)*(Tablo[Satış]))
=DSUM(Tablo[#All];3;F54:F56)
=SUMPRODUCT(((Tablo[Bölge]=G49)+(Tablo[Bölge]=H49))*Tablo[Satış]) //Formül dilinde VEYA</pre>
		<p>Bu yöntemlerin genel değerlendirmesi aşağıda yapıldığı gibi, Excel 
		dosya içinde de her yöntemin artı ve eksikleri belirtilmeye 
		çalışılmıştır.</p>
		<h4>2)Dikey eksende çok kolon var</h4>
		<p>Şimdi ise dikey eksenimizde işlem yapacağımız iki kolon var, ve 
		ayrıca bu iki kolondan da iki kriter birden seçeceğiz. Gördüğünzü gibi 
		işler gittikçe zorlaşıyor.</p>
		<p><img src="../../images/diziformul27.jpg"></p>
		<p>İşler zorlaşıyor ama alternatifler tükenmiyor. Aşağıda yine 1-2 
		alternatif verilmiş, diğer alternatifler dosya içinde bırakılmıştır.</p>
		<pre class="formul">=DSUM(Tablo9[#All];3;F66:G68)
=SUM(SUMIFS(Tablo9[Satış];Tablo9[Bölge];{"Bölge1";"Bölge2"};Tablo9[Yıl];"&lt;2012"))
{=SUMPRODUCT((Tablo9[Bölge]=TRANSPOSE(F67:F70))*(Tablo9[Yıl]=TRANSPOSE(H67:H70))*(Tablo9[Satış]))}</pre>
		<p>Gördüğünzü gibi TRANSPOSE kullanmı olan formüllerde ana fonksiyonumuz 
		SUMPRODUCT olsa bile CSE işlemi yapmamız gerekmektedir.</p>
		<p>DSUM'da ise farkettiyseniz 2012'yi iki kere yazmamız gerekti. F67:G67 
		ile, Bölge1'in 2012'sini alırken, F68:G68 ile de Bölge2'nin 2012'sini almış 
		oluyoruz. Eğer G68'e 2012 yazmazsak Bölge1in 2012si ile Bölge2nin tüm 
		yıllarını toplardı. Bunları excel dosyada deneyip görebilirsiniz.</p>
		<p><strong>NOT:</strong>Ekli dosyadan da göreceğiniz üzere alternatif 
		yöntemler arasında kötü alternatifler de bulunabilir ve bunlar hiç 
		denenmemelidir. (Teorik olarak kullanılabilseler bile). Çünkü şuanda bile 
		2 bölge * 2 dönem=4 kombinasyon varken 3 bölge 3 dönem olduğunda 9 
		kombinasyon olacaktır ki, bu 9 tane manuel kriter girmek demek olacaktır. 
		İkili/üçlü sumifler, Sumproduct'ın veya görevi gören "+" operatörüyle 
		kullanımı bu kötü alternatifler arasındadır.</p>
	</div>
	<h2 class="baslik">2 Boyutlu matrisyel tablolarda çalışmak</h2>
	<div class="konu">
	<p>Aşağıdaki gibi 2 boyutlu matrisyel yapıdaki tablolarla da sık sık çalışıyorsanız 
	bu kısım tam size göre. 
	Bu tablolarda dikey eksen bir veya birkaç kolondan oluşabilirken yatay eksen ise 
	çok kolondan oluşur. Şimdi bu tür tablolarda ne tür işlemler 
	yapabiliyoruz, karşılaştırmalı olarak göreceğiz.</p>
	<p><img src="../../images/dizimatrisyel.jpg"></p>
		<p><strong>UYARI:</strong>1 boyutlu tablolara bakmadan direkt buraya geldiyseniz oradaki "önemli 
		notu" okumanızı tavsiye ederim.</p>
	<h3>A)Dikey Eksenin 1 kolondan oluştuğu durumlar</h3>
	<h4>1)Kriter:Dikey eksenden 1 değer, yatay eksenden 1 değer; dikey eksendeki 
	değer <span style="font-weight: normal; text-decoration: underline"><strong>
	bir kez </strong></span>geçiyor (Kesişim değeri)</h4>
	<p>Kesişim bulma örneklerini daha önce 
	<a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx#Kesisim">lookup fonksiyonlarını</a> incelerken 
	görmüştük. Şimdi onlara bir alterantif olarak da SUMPRODUCT geliyor. Tabi 
	SUMPRODUCT diğerlerinden farklı olarak çok satır ve sütunda kesişenlerin toplamını da 
	alabiliyor, onları da sonraki maddelere göreceğiz.</p>
	<p><img src="../../images/diziformul22.jpg">&nbsp;</p>
	<p>Yukarda M2'nin Ürün3'teki satış rakamını bulmak istiyoruz. Formülümüz şöyle 
	olacaktır.J8'de Mağaza seçimi, K8'de Ürün seçimi yapılıyor olsun. Ayrıca 
	formül okunurluğu adına Name'ler kullanılmıştır.</p>
	<pre class="formul">=SUMPRODUCT((Mağazalar=J8)*(Ürünler3=K8)*Tutarlar3)</pre>
	<p>İlk iki parantezli grubun önüne -- koyduktan sonra F9 yapınca görüntü 
	aşağıdaki gibi oluyor. Son 
	parantezli grupta ise değerler var, hepsi birbiriyle çarpılınca yanlızca 
	2.satır 3.sütundaki değer sıfırlanmadan kalır, o da aradığımız değerdir, 
	yani 532.</p>
		<p>{0;<span style="color: red">1</span>;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0}*{0\0\<span style="color: red">1</span>\0}&nbsp;</p>
		<p>Bir diğer alternatif de SUMIF+OFFSET+MATCH kombinasyonudur.</p>
		<pre class="formul">=SUMIF(Mağazalar;J8;OFFSET(Mağazalar;0;MATCH(K8;Ürünler3;0)))</pre>
		<p>Mağazalar kolonunda M2'yi arattırıyoruz, toplam işlemini hangi 
		kolonda yapacağımızı ise Mağaza kolonunu referans verip Ürün3'ün bu 
		referanstan sonraki kaçıncı 
		kolonda geçtiğini buldurarak dolaylı şekilde bulduruyoruz. Bunu da bu 
		sıra numarasını, OFFSET'in ikinci parametresi olarak vererek buldurmuş 
		oluyoruz.</p>
	<h4>2)Kriter:Dikey eksenden 1 değer, yatay eksenden 1 değer; dikey eksendeki 
	değer birden çok kez geçiyor. (Kesişenlerin toplamı)</h4>
	<p>Diyelim ki aşağıdaki bir tablomuz var. Bölge1'e ait Ürün3 satışları 
	görmek istiyoruz. (Dikey eksende Bölge-Mağaza olmak üzere 2 kolon olmasına rağmen biz sadece 
	Bölge kolonu üzerinden işlem yapacağımız için bu tablo bu kategoriye uygundur) </p>
	<p>Şimdi Alternatiflerimize bakalım.</p>
	<p><img src="../../images/diziformul24.jpg">&nbsp;</p>
	<p>Burada klasik SUMIF yapabiliriz. Tabi ürün adı değişken olacağı için 
	OFFSET-MATCH ikilisinden yararlanırız. Baz olarak Mağazayı alıp bunun kaç 
	kolon sağında olduğunu dinamik bir şekilde buldurabiliriz.</p>
	<p>J12'de Bölge, K12'de ürünlere ait Comboboxlar olduğu gözününde 
	bulundurulursa;</p>
	<pre class="formul">=SUMIF(bölgeler3;J12;OFFSET(Mağazalar;0;MATCH(K12;Ürünler3;0)))</pre>
	<p>Bir diğer alternatif de SUMPRODUCT kullanmak olacaktır</p>
	<pre class="formul">=SUMPRODUCT((bölgeler3=J12)*(Ürünler3=K12)*Tutarlar3)</pre>
	<p>Burada da yaptığımız normal kesişim formülü yazmak gibidir aslında, sadece birden 
	fazla eşleşme olduğunda sonuçlar toplanmış oluyor, o kadar.</p>
		<p>Ve tabiki yine Veritabanı fonksiyonları kullanılabilir.</p>
		<pre class="formul">=DSUM(B6:G24;MATCH(K12;B6:G6;0);J11:J12)</pre>
	<h4>3)Kriter:Dikey eksenden 1 değer, yatay eksenden çok değer; yatay 
	eksendeki kriterler yanyana değil(dikey eksendeki değer 1 veya daha çok kez 
	geçebilir)</h4>
	<p>2.maddedeki tabloyu kullanacağız. Bu tabloda Bölge2'nin Ürün2 ve ürün4 
	toplamını bulacağız. Böyle saçma şey olur mu demeyin, olur. Ör:Bankacılık 
	dünyasından bir örnek düşünecek olursak Vadeli TL ve Vadesiz TL olabilir,ve 
	biz Toplam TL mevduat bulmak 
	istiyor olabiliriz .</p>
		<p>Bu tabloda kriter alanımızın aşağıdaki gibi olduğunu gözönünde 
		bulunduralım.</p>
		<p><img src="../../images/diziformul29.jpg"></p>
	<p>Şimdi, hatırlayalım, konuşma dilindeki "<strong>Ürün2 ve Ürün4</strong>"'ün formül dili karşılğı 
	<strong>Ürün2 VEYA Ürün4 </strong>idi. <strong>VEYA </strong>operatörümüz de + işaretiydi. O halde 
	ilk formülümüz aşağıdaki gibi 
	olacaktır.</p>
	<pre class="formul">=SUMPRODUCT((bölgeler3=J17)*((Ürünler3=K17)+(Ürünler3=K18))*Tutarlar3)</pre>
	<p>Bir diğer alternatif de Ürün2 ve Ürün4 toplamından oluşan yardımcı kolon 
	oluşturmaktır ama bu çok kötü bir alternatiftir. Zira burdaki amacımız 
	yardımcı kolon oluşturmadan işlerimizi halletmektir, üstelik yardımcı kolonlar 
	da genelde daha anlamlı içeriğe sahiptirler; tüm ürünlerin toplamı gibi 
	mesela. İki 
	ürünün toplamından oluşan bir yardımcı kolon pek iyi bir fikir değil.</p>
		<p>Başka bir alternatif ise aşağıdaki gibi olabilir</p>
		<pre class="formul">{=SUMPRODUCT((bölgeler3=J17)*(Ürünler3=TRANSPOSE(K16:K19))*(Tutarlar3))}</pre>
		<p>Burada dikkat ettiyseniz ürünleri alırken sadece iki ürünü 
		seçemiyoruz. Eğer yatay kolondan ürün eşleştirmek istiyorsak <strong>tüm 
		ürünlerin sırasına ve sayısına uygun bir kriter dizesini </strong>formüle 
		sokmalıyız. Burda 4 ürün var ve sırası Ürün1,Ürün2,Ürün3 ve Ürün4 
		şeklinde. Biz de K16'da sanki Ürün1, K18'de de sanki Ürün3 varmış gibi 
		düşünerek 4 hücrelik bir alan seçmek durumundayız, bunlardan sadece K17 ve 
		K19'dakiler eşleşeceği için doğru sonucu elde etmiş oluruz. Bu sırada 
		kriter listeleri hazırlamak zahmetli olabileceği için uygulaması biraz zor 
		olabilir ama kriterleri tek tek girmemek adına da kolaylık sağlar.</p>
		<p>Dikkat edilecek bir diğer husus da, kriteler dikey eksende olduğu 
		için bunları TRANSPOSE ile çevirmemiz 
		gerektiğidir.</p>
		<p>Son yöntemimiz yine SUMPRODUCT ile. Bunda da kriterler elle girilir. 
		Eğer sabit bir fomrül olacaksa ve comboboxlardan değiştirme ihtiyacı 
		olmayacaksa bu yöntem de uygulanabilir.</p>
		<pre class="formul">=SUMPRODUCT(--ISNUMBER(MATCH(Ürünler3;{"Ürün2";"Ürün4"};0))*(bölgeler3=J17)*(Tutarlar3))</pre>
	<p>Bunda ürünler MATCH ile eşleşiyormu diye kontrol edilip, eşleşenlere 
	eşleşme sırası, eşleşmeyenlere N/A atanmış olur. Sonra bunlar ISNUMBER ile TRUE 
	ve FALSE değerlerine, son olarak da -- ile 1 ve 0'a dönüştürülür. Bunların hepsini aşama aşama aşağıda 
	görebilirsiniz.</p>
		<p>MATCH'li kısım--&gt;{#N/A\1\#N/A\2}<br>ISNUMBER'lu 
		kısım--&gt;{FALSE\TRUE\FALSE\TRUE}<br>--'li kısım--&gt;{0\1\0\1}</p>
		<h4>4)Kriter:Dikey eksenden 1 değer, yatay eksenden çok değer; yatay 
	eksendeki kriterler yanyana(dikey eksendeki değer 1 veya daha çok kez 
	geçebilir)</h4>
	<p>Komşu hücrelerin toplanması sözkonusu olduğunda bir VE işlemi yoktur, tek 
	bir alan(range) işleme girmektedir. Aslında 3.maddeden farklı bir yöntem 
	bulunmamaktadır. Ordaki 3 yöntem aynen uygulanabilir. Yani özeetle komşu olma durumuna 
	özgü ayrı bir yöntem bulunmamaktadır. </p>
		<p>Bu maddenin özel durumu olarak "tüm ürünler" düşünülebilir. Yani Bölge2'nin 4 
		ürün toplamını almak istiyoruz diyelim. Bunun için Yardımcı kolon olarak 
		en sağda bi toplamı kolonu oluşturulup basit SUMIF yapılabilir. Acil 
		çözümler için oldukça geçerlidir. Ancak kalıcı ve şık çözümler için yine 
		SUMPRODUCT'a ihtiyacımız vardır.</p>
		<p>J27'de Bölge comoboboxı olduğunu düşünürsek,</p>
		<pre class="formul">=SUMPRODUCT((bölgeler3=J27)*Tutarlar3)</pre>
	<p>Yeri gelmişken önemli bir hususu belirtmekte fayda var:<strong> SUMPRDOCUT'ın çok kolonlu durumlarda sadece"*"lı versiyonunun kullanılabilmesidir. ";"li versiyonu hata verir.</strong></p>
	<p>Bir diğer önemli husus da SUMIF'in çok kolonlu kullanımında hata 
	ver<span style="text-decoration: underline">me</span>mesi ancak yanlış sonuç 
	döndürmesi, çünkü sadece ilk kolonun sonucunu 
	döndürür.</p>
	<pre class="formul">=SUMIF(Table14[Yıl];2010;Table14[[Ürün1]:[Ürün4]]) //hatalı sonuç</pre>
	<h3>B)Dikey Eksenin çok kolondan oluştuğu durumlar</h3>
	<p>Dikey eksende birden çok kolondan seçim yapmak konuşma dilinde iyelik 
	eki, formül dilinde ise <strong>VE </strong>kriteri olup bununla 
	ilgili ne yapılacağını dizi formülleri bölümünde görmüştük. Yukardaki 4 madde 
	için 4 farklı versiyon yapılabilir ama biz burada sadece bir tanesini 
	yapacağız.</p>
	<p>Şimdi aşağıdaki tabloyu örnek alalım.</p>
	<p><img src="../../images/diziformul30.jpg"></p>
	<p>Bölge 2'nin 2014teki Ürün1 ve Ürün3 toplamını arıyoruz. Kriter bölgesi 
	aşağıdaki gibidir.</p>
		<p><img src="../../images/diziformul31.jpg"></p>
		<p>Burdaki 
	alternatiflerimiz şöyledir:</p>
	<pre class="formul">=SUMPRODUCT((bölgeler4=J38)*(yıllar4=K38)*((ürünler4=L38)+(ürünler4=L39))*tutarlar4)=SUM(IF(Table14[Bölge]="Bölge2";IF(Table14[Yıl]&lt;2013;Table14[[Ürün2]:[Ürün3]]))) //Çok kriter olursa çok sayıa içiçe if olacağı için yazımı zorlaşır
=SUMPRODUCT((bölgeler4=J38)*(yıllar4=K38)*(ürünler4=TRANSPOSE(L38:L40))*tutarlar4)</pre>
	<h3>C)Dikey Eksenden çok değer seçimi</h3>
	<p>Bu kategori için de sadece tek bir örnek yapılacak olup yukardaki 
	örneklere göre çeşitlendirme yapılabilir. Ekli dosyadan da diğer 
	alternatifleri görebilir, orada bulunmayan alternatifler üzerinde kendiniz 
	de çalışarak pratik yapabilirsiniz.</p>
		<p>Yukarıdaki son tabloyu örnek alalım. Burda Bölge1 ve Bölge2'nin 2014 
		toplamını bulmak istiyoruz. İki alternatifimiz şöyle olacaktır. 
		(Q49:Q50'de Bölge1 ve 2 var, R49'da 2014)</p>
		<pre class="formul">=SUMPRODUCT(--ISNUMBER(MATCH(bölgeler4;Q49:Q50;0))*(yıllar4=R49)*tutarlar4)
=SUM(SUMIFS(H34:H51;bölgeler4;Q49:Q50;yıllar4;R49)) //SUMIFS ile yardımcı kolon</pre>
		<p>Artık iyice kanıksamış olduğunuzu düşündüğüm için fomrülü tek tek 
		açıklamaya gerek duymuyorum.</p>
		<h3>D)Farklı kolonlarda VEYA işlemi uygulamak</h3>
		<p>İlginç bir durum da VEYA(formül diliyle) kriterinin farklı kolonlarda 
		uygulandığı durumlardır. Şimdiye kadar VEYA kriterini hep aynı kolon 
		üzerinde uyguladık; Bölge1 VEYA Bölge2, Ürün1 VEYA Ürün2 demek gibi.</p>
		<p>Aşağıdaki örneklerde ise farklı kolonlarda VEYA şartı aranacak.</p>
		<h4>1)Kesişmeyen VEYA'lar</h4>
		<p><img src="../../images/diziformul32.jpg"></p>
		<p>Bu tabloda Personel Tipi=Gişe olan veya Miy tipi BMIY olan kişilerin 
		4 üründeki toplam satışını arıyoruz. Bunda, önceki örneklere göre çok 
		fakrlı bir durum yok. VEYA koşulumuzu "+" operatörü ile kurgulayarak 
		formülümüzü yazalım.(I57'de Gişeci, J57de BMIY yazmakta)</p>
		<pre class="formul">=SUMPRODUCT(((perstip=I57)+(miytip=J57))*satış)</pre>
		<h4>2)Kesişen VEYA'lar</h4>
		<p><img src="../../images/diziformul33.jpg"></p>
		<p>Bu sefer iki sertifikadan herhangi birine sahip Miylerin toplam 
		satışlarını bulmak istiyoruz. Bunda da yardımcı kolonlar kullanılabilir, 
		hatta iki sertifkanın yanına bir kolon daha açılıp, "ikisi de E ise 1 
		değilse 0" denip, SUMIF uygulanabilir, ama bunlar bizim için çirkin 
		yöntemler. Biz daha şık olan alternatife bakalım. Ama öncesinde bir 
		üstteki yöntemin aynısını uygularsak nasıl hata yapacağımızı görelim.</p>
		<pre class="formul">=SUMPRODUCT(((cert1="E")+(cert2="E"))*certsatış)</pre>
		<p>Eğer iki operatörün toplandığı parantezi seçip F9 yaparsak sonucun 
		{2;1;2;2;0;2;0;0;2;2;1;0;1;1;2;0;0;1} olduğunu görürüz. Evet elimizde 
		0/1den başka 2'ler de olduğunu görüyoruz, çünkü aynı anda iki kriteri 
		sağlayan satırlar 2 döndürüyor. Bizim bunları 1'e döndürmemiz lazım. 
		Bunun için de belki bir önceki sayfada gördüğünüz ama belki hiç dikkat 
		etmediğiniz, ve belki de görüp "işime yaramaz dediğiniz" <strong>SIGN
		</strong>fonksiyonu yardımımıza koşar. SIGN ile sayıların işaretini alıyorduk, 
		pozitifler 1, negatifler -1, 0'lar 0 oluyor, tam da aradığımız şey.</p>
		<pre class="formul">=SUMPRODUCT(SIGN((cert1="E")+(cert2="E"))*certsatış)</pre>
	</div>
	<h2 class="baslik">Genel değerlendirme</h2>
	<div class="konu">
	<p>Şimdiye kadar gördüğümüz üzere belli bir amacı yerine getirmek için 
	çeşitli alternatifler var. Kimisi şık çözümler sunarken kimisi şık olmayan 
	ama daha hızlı çözümler sunabilmekte. Şimdi, hangi durumlarda hangi 
	alternatifleri kullanabileceğimizin bir özetini vermek istiyorum.</p>
	<ul>
		<li>Mümkün olduğunda Excel versiyonunuzun desteklediği yerel fonksiyonaları kullanın. 
		SUMIFS(2007 ve sonrası), MINIFS(2016 ve sonrası) gibi</li>
		<li>Ayrı bir kriter alanı yaratmak sıkıntı değilse bunu yaratın ve 
		<strong>Database fonksiyonlarını </strong>kullanın. Özelikle <strong>Data Validation
		</strong>ile 
		comboboxlardan yararlanacaksanız bu fonksiyonlar çok kullanışlı 
		olmaktadır. Çünkü hem daha 
		hızlı çalışırlar hem de yazılışları basittir. (Hız etkisini bir iki 
		hücreli çalışmalarda hissetmeyebilirsiniz ama çok fazla dizi formülünüz 
		varsa bunları Database fonksyionlarına çevirmenizi öneririm). Buların en 
		önemli eksikliği çok kolon üzerinde işlem yapamıyor oluşlarıdır. 
		Kriterler manuel girilecekse Database fonksiyonları kullanılamaz.</li>
		<li>Yatay eksenden çok kolon üzerinde işlem yapılacaksa <strong>SUMPRODUCT
		</strong>uygun alternatiftir.</li>
		<li>Database Fonksiyonu kullanılamayan durumlarda mümkünse <strong>
		SUMPRODUCT </strong>kullanın, SUM-dizi tarzıdaki dizi formüllerini terih 
		etmeyin.</li>
		<li>TRANSPOSE kullanımı veya bir şekilde dizi üreten diğer durumlarda 
		formül bitiminde CSE yapılacaksa SUMPRODUCT yerine dizi formüllerini 
		kullanabilirsiniz.</li>
	</ul>
	</div>
	
	<h2 class="baslik">Çeşitli örnekler</h2>
	<div class="konu">
		<h4 class="baslik">0 hariç en küçük sayı</h4>
		<div class="konu">
	<p>Yapmamız gerek şey, 0'dan büyük hücreleri dizi şeklinde elde edip bunlara 
	MIN işlemi uygulamak.</p>
		<p><img src="../../images/diziformul34.jpg"></p>
		<pre class="formul">{=MIN(IF(A2:J2&gt;0;A2:J2))}</pre>
	<p>Bunun bir altenratifi de SMALL ile dizi formülü yapmadan da aşağıdaki 
	gibi yazılabilir. 0'dan kaç tane var diye bakıyorum, çıkan sonucun bir 
	fazlası x olsun. x. küçük elemanı getir diyorum.</p>
		<pre class="formul">=SMALL(A2:J2;COUNTIF(A2:J2;0)+1)</pre>
		</div>
		<h4 class="baslik">En küçük 3 elemanın kendisi ve bunların toplamı</h4>
	<div class="konu">
		<p>
		<img src="../../images/diziformul34.jpg"></p>
		<p>
		Bu örnekte çoklu hücre seçimi yaparak 3 hücreye en küçük 3 elemanı 
		gireceğiz.</p>
		<pre class="formul">{=SMALL(A2:J2;COLUMN(INDIRECT(1&amp;":"&amp;3)))} //Kolon sıra numarası ile
{=SMALL(A2:J2;COLUMN(INDIRECT("A:C")))} //Kolon başlığı olan harflerle </pre>
		<p>
		Bu 3 elemanın toplamını ise ya bu 3 sonucu toplatarak buluruz veya tek 
		seferde yapmak istersek aşağıdaki formülü gireriz.</p>
		<pre class="formul">{=SUM(SMALL(A2:J2;COLUMN(INDIRECT("A:C"))))}</pre>
		<p>
		Farkettiyseniz değerleri ayrı ayrı yazdırıraen hem 1:3 hem de A:C 
		şeklinde girebiliyoruz ancak toplam aldırırken sade A:C versiyonu işe 
		yaramakta.</p>
				</div>
		<h4 class="baslik">Hatalı değer içeren alanlarda toplama yapmak</h4>
	<div class="konu">
		<p>
		Bildiğiniz gibi bir kolonda/satırda N/A gibi değerler varsa SUM 
		fomrülnün sonucu da N/A olmakta, aşağıaki durum çubuğunda da değer 
		görünmemektedir. Adetlerde sorun yok, nları direkt sayar, ama hatalı kaç 
		kayıt var bunu görmek için dizi formül yazmmız lazım. Bunları da 
		aşağıdaki gii </p>
		<pre class="formul">=SUM(IF(ISERROR(D9:D11);1)) //Hatalı kayıt sayısı
=SUM(IF(ISERROR(D9:D11);0;D9:D11)) //Hatasızların toplamı</pre>
		</div>
		<h4 class="baslik">Bir alandaki benzersiz değerleri saydırmak</h4>
	<div class="konu">
			<p>
		Yine yukarıdaki tabloyu kullanalım.</p>
		<p>
		<img src="../../images/diziformul34.jpg"></p>
		<p>Burada toplam 10 sayı var, bunlardan 157 iki kere geçiyor, o yüzden 
		benzersiz sayı adedi 9'dur. Bunu bulmak için izleyeceğimiz yol FREQUENCY 
		formülünü dizi formülü şeklinde kullanıp her rakamın kaç kez geçtiğini 
		bulmak, sonra bunların 0'dan büyük olup olmadığını kontrol edip 
		TRUE/FALSE döndürmek, en sonunda bu TRUE/FALSEları da -- ile 1/0'a 
		dönüştürüp toplamak olacaktır.</p>
		<p>Tabi burada FREQUENCY'nin işleyişini iyi bilmek gerekiyor. İstatistik 
		formüllerinde gördük ki, bu fonksiyon değerlerin sıklığını ele alırken, 
		bir değeri <strong>ilk kez gördüğü yerde </strong>onun frekansını yazar, 
		sonrakilerde 0 yazar.</p>
		<pre class="formul">{=SUM(--(FREQUENCY(A2:J2;A2:J2)&gt;0))}</pre>
		<p>Formülün çözümlemesi şöyle:</p>
		<p>* FREQUENCY kısmı: {2;1;1;0;1;1;1;1;1;1;0} dizisini döndürü<br>* Dizi&gt;0: 
		{TRUE;TRUE;TRUE;FALSE;TRUE;TRUE;TRUE;TRUE;TRUE;TRUE;FALSE} döndürür<br>
		*
		--'li kısım: {1;1;1;0;1;1;1;1;1;1;0} döndürür<br><strong>Sonuç</strong>:9</p>
		<p>Bu işlemin bir diğer alternatifini de bir alt örnekte göreceğiz</p>
				</div>
		<h4 class="baslik">Bir alandaki benzersiz değerleri saydırmak-2</h4>
	<div class="konu">
		<p>Diyelim ki şubelere mevduat spread hedefi vereceğiz. Yöneticiniz 
		sizden 800 şube için en fazla 8 çeşit spread hedefi vermenizi istemiş 
		olsun. Yani şubeler öyle gruplanalacak ki ens onunda bu 8 grup hedften 
		birinde yer alacaklar.</p>
		<p>Örneği basitleştirmek adına biz 20 şube baz alalım. İlk olarak da 
		aşağıdaki gibi bir çalışma yaptınız, şimdi kontrol etmek istiyoruz, kaç 
		çeşit hedef var diye.</p>
		<p><img src="../../images/diziformul35.jpg"></p>
		<p>Bunu yapmanın iki kötü yolundan biri Remove Duplicates yapıp, diğeri 
		de Pivot table uygulayıp saymak olacaktır. Kötüler çünkü ilkinde 
		isediğiniz sayıya gelene kadar her defasında işlemi yenilemeniz, 
		ikincisinde ise pivot tabloyu refresh etmeniz gerekir.</p>
		<p>İyi yöntemler ise bir hücreye formül yazmak olacaktır. FREQUENCY ile 
		dizi formülü yazmayı önceki örnekte görmüştük, tekrar yazmayacağız ancak 
		ekli dosyada bulabilirsiniz. Diğer dizi formülü ise aşağıdaki gibi 
		olacaktır</p>
		<pre class="formul">{=SUM(1/COUNTIF(B2:B21;B2:B21))}</pre>
		<p>Formülü çözümleyelim</p>
		<p>* COUNTIF'li kısım: {2;<span style="color: red"><strong>3</strong></span>;2;1;1;3;<span style="color: red"><strong>3</strong></span>;1;2;1;1;1;1;3;2;<span style="color: red"><strong>3</strong></span>;1;1;3;1}<br>
		*
		1/COUNTIF:{0,5;0,333333333333333;0,5;1;1;0,333333333333333;0,333333333333333;1;0,5;1;1;1;1<br>;0,333333333333333;0,5;0,333333333333333;1;1;0,333333333333333;1}<br>
		<strong>*
		Sonuç</strong>:14</p>
		<p>Formülün mantığı şöyle: İlgili alanı kendisiyle COUNTIF'e tabi 
		tutarak her değerden o alanda kaç tane olduğunu buluyoruz. Sonra farklı 
		değerlerin toplamının 1 olması için bunların 1'e bölünmüş halini 
		alıyoruz. Mesela 11,9'dan 3 tane var(kalın kırmızılı olanlar). Bunları 
		toplayınca 1 yapıyor. Diğerlerini de aynı şekilde toplayınca 1 yapar. 
		Tek olanlar zaten 1 olduğu için 1/1=1 yapıyoruz. Sonuç olarak hepsini 
		toplayınca da aradığımız sonuç olan 14'e ulaşıyoruz.</p>
				</div>
		<h4 class="baslik">Bir alandaki en uzun metin</h4>
	<div class="konu">
		<p class="baslik">Diyelim ki bir alanda girilmiş çeşitli açıklamalar 
		var. Siz bunlardan en detaylısını yani en çok karatkter içeren olanı 
		arıyorsunuz.</p>
		<p>
		<img src="../../images/diziformul37.jpg"></p>
		<p>Formülümüz aşağıdaki gibidir.</p>
		<pre class="formul">{=OFFSET(A2;MATCH(MAX(LEN(A2:A5));LEN(A2:A5);0)-1;0)}</pre>

			<p>Bu formülün çözümlemesi de şöyledir</p>
			<p>* Dizi1=LEN(A2:A5): {8;10;13;9}<br>* X=Max(Dizi1): 13 //Bunları 
			en büyüğü<br>* Y=MATCH(X;Dizi1;0): 3 //13'ün bu dizdeki sırası<br>* 
			OFFSET(A2;Y-1;0) //A2'den 2 satır aşağı<br><strong>* Sonuç</strong>:Açıklama 
			1231</p>
					</div>
			<h4 class="baslik">RANKIF ve RANKIFS</h4>
	<div class="konu">
		<p>Bunu daha önceden COUNTIFS ile yapmıştık, şimdi aynısını SUMPRODUCT 
		ile yapmakta.</p>
		<p>Yamak istediğimiz şey her bölgenin şubelerini hacimsel büyüklüğüne 
		göre boy sırasına dizmek.</p>
		<p><img src="../../images/diziformul38.jpg"></p>
		<pre class="formul">=SUMPRODUCT((bölge=A2)*(C2&lt;rakam))+1</pre>
		<p><strong>NOT</strong>:Bu yöntem de tıpkı COUNTIFSte olduğu gibi daha 
		ok koşul eklenerek RANKIF<span style="color: red"><strong>S</strong></span> 
		olarak da çalışabilir.</p>
				</div>
		<h4 class="baslik">En yüksek/düşük x Rakam Toplamı - En yüksek/düşük x 
		rakam hariç Toplam</h4>
	<div class="konu">
		<p>Aşağıdaki listeden 2 nolu şubenin Para çekme işlemindeki en düşük 3 
		işleminin toplamını almak istiyoruz.</p>
		<p><img src="../../images/diziformul39.jpg"></p>
		<p>Formülümüz şöyledir. (F2'de şube kodu olarak 2, G2'de de "Para çek" 
		yazıyor)</p>
		<pre class="formul">{=SUM(SMALL(IF(A2:A161=F2;IF(B2:B161=G2;C2:C161));ROW(INDIRECT(1&amp;":"&amp;3))))}</pre>
		<p>Çözümlemesi şöyle</p>
		<p>* Dizi1:IF(A2:A161=F2;IF(B2:B161=G2;C2:C161)) ile 2 nolu şubenin 
		Paraçek işlein olduğu listeyi elde ediyoruz, yani C2:C10 arası döner.<br>
		* Dizi2:ROW-INDRIECT'li kısım:{1;2;3}<br>* 
		SMALL(Dizi1;Dizi2):{263;290;481}<br>* <strong>Sonuç</strong>:1034</p>
		<p>Bir de en düşük/yüksek 3 işlem hariç toplam ne kadar diye bakalım. 
		(F7=2, G7=Para çek)</p>
		<pre class="formul">=SUMIFS(C:C;A:A;F7;B:B;G7)-SUM(SMALL(IF(A2:A161=F7;IF(B2:B161=G7;C2:C161));ROW(INDIRECT(1&amp;":"&amp;3))))-SUM(LARGE(IF(A2:A161=F7;IF(B2:B161=G7;C2:C161));ROW(INDIRECT(1&amp;":"&amp;3))))</pre>
		<p>Burda yapılan da SUMIFS yaparak full toplam almak ve ondan bir 
		üstteki formülü ve bunun LARGE'lı versiyonunu çıkarmaktır.</p>
		<p>Gördüğünüz gibi oldukça uzun bi formül oldu bu. Açıkçası yazması bile 
		meşakkatli bu formülün çözümlemesi hiç girmek istemiyorum, bunu bu 
		noktada artık sizin yapabileceğinizi düşünüyorum. Bu arada bunun bir de 
		SUMPRODUCT'lı versiyonu var, ona ekli dosyadan bakarsınız. </p>
		<p>Evet, uzunca bir konunun sonuna geldik. Umarım faydalı olmuştur. Her 
		konuda olduğu gibi bunda da tam kavrayış için bol tekrar ile 
		pekiştirilmesi gerekmektedir.</p>
		</div>
</div>
</asp:Content>
