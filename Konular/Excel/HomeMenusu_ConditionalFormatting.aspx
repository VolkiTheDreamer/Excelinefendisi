<%@ Page Title='' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Home Menüsü'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Conditional Formatting</h1>
	<p>Bu bölümde Conditional Formatting(Bundan sonra CF olarak geçecek) işlemlerini baştan anlatmayacağım ve genel olarak ne işe yaradığını bildiğinizi varsayıyorum. Şimdi birkaç örnekle çeşitli püf noktalarına 
ve az bilinen kullanım şekillerine bakalım.</p>
	<p>Öncesinde, CF'ı sadece <strong>Greater/Less than</strong> veya <strong>Top/Bottom 10</strong> şeklinde 
	değil, aşağıdaki diğer özellikleri de içerecek şekilde kullanmanızı 
	öneririm. Mesela, mükerrer(duplike) kayıtlarınızı bulmak için başka bir 
	sütuna COUNTIF formülü yazıp oraya "&gt;1" filtresi uygulamak yerine doğrudan 
	Duplicate Values CF'ını uygulayabilir ve ilgili renge filtre koyabilirsiniz. 
	Veya ortalamanın altında/üstünde kalanları filtrelemek için bir yere ortalama 
	formül yazıp onunla karşılaştırma yapmanıza gerek yok, bunu direkt yapmanızı 
	sağlayan seçenekler de var. Özetle diğer seçenekleri de biraz kurcalayın 
	derim.</p>
<img src="/images/homeCF10.jpg">


	<h2 class="baslik">Yüzdesel eşikler belirleme</h2>
	<div class="konu">
	<p>
	CF'da en sık karıştırılan konulardan biri, belirli yüzdesel eşiklerin 
	altını/üstünü formatlı göstermektir. Malesef 
	<span style="text-decoration: underline">bana göre</span> bu konuda Microsoft hatalı 
	yönlendirmede bulunmakta, daha doğrusu gerekli yönlendimeyi 
	yapmamaktadır. İşin doğrusunu öğrenene kadar bayağı bir araştırma yapmanız 
	gerekmektedir. </p>
		<p>
		Örneğin aşağıdaki tabloda D kolonuna CF uygulayacacağız. Diyeceğiz ki 
		%100 ve üzerini en iyi olarak(yeşil ikon), %90-100 arasını orta(sarı 
		ikon), 90 altını da kötü(kırmızı ikon) olarak göster.</p>
		<p>
		<img src="/images/HomeCFPerc1.jpg"></p>
		<p>
		Aşağıdaki ikon stini uyguladım</p>
		<p>
		<img src="/images/HomeCFPerc2.jpg"></p>
		<p>
		CF, otomatik olarak 33&nbsp; ve 67 şeklinde ayarladı. Ben gittim bunu 90 
		ve 100 olarak değiştirdim. </p>
		<p>
		<img src="/images/HomeCFPerc3.jpg"></p>
		<p>
		Sonuç şöyle:</p>
		<p>
		<img src="/images/HomeCFPerc4.jpg"></p>
		<p>
		Gördüğünüz gibi %100 ve üzerinde 6 bölge olmasına rağmen sadece birini 
		yeşil yaptı, 90-100 arası 1 bölge olmasına rağmen hiç sarı yapmadı. 
		Neden böyle oldu?</p>
		<p>
		Çünkü Excel, buradaki verilerin en düşüğünü %0, en yükseğini %100 
		varsayarak yeni bir değer hesaplıyor ve onun üzerinden CF uyguluyor. 
		Zaten ilk başta %33 ve %67 diye otomatik ayırmasının sebebi de bu. 
		Değerleri %0 ve %100 arasında olacak şekilde diziyor, 3 dilim 
		belirliyor, ve rakamları bu 3 dilim içine yerleştiriyor. 
		Uygulanan formül şöyle:</p>
		<pre class="formul">=(D2-MIN(D:D))/(MAX(D:D)-MIN(D:D))</pre>
		<p>Buna göre hesaplanmış HG%ler E kolonundaki gibi oluyor ve bu 
		değerlere göre CF uygulandığında da sonuç gayet normal(!). %100 ve üzerinde 
		sadece 1 bölge var, 90-100 arasında ise hiç yok.</p>
		<img src="/images/HomeCFPerc5.jpg">
	<p>Bizim istediğimiz ise böyle birşey değildi. O yüzden <strong>Percent</strong> ifadesinin bizi yanıltmasına izin vermeyelim ve
	<strong>Number</strong> seçeneğini aşağıdaki gibi ayarlayalım.</p>
	<p><img src="/images/HomeCFPerc6.jpg"></p>
	<p>Sonuç:</p>
	<p><img src="/images/HomeCFPerc7.jpg"></p>
	</div>

	<h2 class="baslik">Başka hücrelere göre göreceli referansla formatlama</h2>
	<div class="konu">
	<p>
	Sıklıkla, formatlama işlemi komşu veya başka hücrelere göre yapılmaktadır. 
	Mesela format uygulayacağımız&nbsp; hücre bu yılın artış oranını, bir yan 
	hücre ise geçen yılın artış oranını veriyordur. Bu&nbsp; yılın artış oranı 
	geçen yılın artış oranından büyükse/küçükse şöyle şöyle formatla 
	diyebiliriz. Ancak burda bazı nüanslar var. Şimdi bunlara bakalım.</p>
		<p>
		CF'da karıştırılan, anlaşılması zor olan bir konu da mutlak ve göreceli başvurulardır, 
		ki bu konu formüle dayalı CF'in can damarını oluşturmaktadır. Eğer bu 
		konu tam anlaşılamazsa CF de verimli ve etkin kullanılamaz. 
		Mutlak/Göreceli başvuru tipleri için <a href="/Konular/Excel/FormulasMenusuDiger_PufNoktalari.aspx#basvurutip">buraya </a>
		bakabilirsiniz. <span style="color: red"><strong>Özetle CF dilinde "=$A$1", "=$A1", "=A$1" ve "=A1" tamamen farklı sonuçlar 
		üretir. Örnek üzerinden durumu açıklayalım.</strong></span></p>
		<p>
		Şimdi aşağıdaki tabloda <strong>Highlight Cell Rules&gt;Greater Than</strong> 
		dedik ve değer olarak da E2 hücresini seçtik. Default olarak mutlak 
		başvuru yazıldı. Buna göre seçilen alandaki tüm hücrelerin E1'in 
		değerinden yani %81den büyük olmasına baktı. Mutlak başvuruda, CF 
		uygulanan hücrelerin karşılaştırma hücresine göre konumu önemsizdir ve 
		sabittir. Yani bu örnek için tüm hücreler için E1 hücresiyle 
		karşılaştırma yapılır.</p>
		<p>
		<img src="/images/homeCF11.jpg"></p>
		<p>
		Mutlak başvuru yerine göreceli başvuru yaparsak durumun nasıl 
		değiştiğini görelim. Bu sefer D3 hücresini E3 ile, D4'ü E4 v.s ile 
		karşılaştırır.</p>
		<p>
		<img src="/images/homeCF12.jpg"></p>
		<p>
		Burda yapılacak bir hata da şu olabilir. İlk CF'ı yaptınız ve mutlak 
		başvuru kullandığınızı farkettiniz, bunu düzeltmek yerine mevcudun üzerine 
		ikinci bir CF uygulayıp orada göreceli yaparsanız etkisi olmaz, çünkü 
		ilk CF ikincisini ezmiş olur. Zaten Rule Managera girince 2 tane 
		CF olduğunu da görebilirsiniz. </p>
		<p>
		<a href="https://support.office.com/en-us/article/Manage-conditional-formatting-rule-precedence-063cde21-516e-45ca-83f5-8e8126076249">
		Burada</a> ve
		<a href="https://www.extendoffice.com/documents/excel/2529-excel-conditional-format-stop-if-true.html">
		burada</a>, birden fazla CF'ın öncelik işleyişi hakkında detaylı bilgi 
		bulunmaktadır.</p>
		<p>
		<img src="/images/homeCF13.jpg"></p>
		<p>
		<span class="dikkat">Dikkat</span>:Göreceli başvuruları, Icon setleri ve 
		Color Scale formatingle kullanamazsınız. Bunlarda nasıl biz çözüm 
		uygulanacağını bir sonraki örnekte bulabilirsiniz.</p>
	
	</div>

<h2 class='baslik'>Icon setlerde göreceli başvuru kullanma(VBA içerir)</h2>
<div class='konu'>
<p>Aşağıdaki gibi bir tablomuz olsun. Bu, bir bankada çeşitli kalemlerde X bölge müdürlüğü ile bankanın toplam performansını karşılaştıran bir tablodur.</p>
	<p>	<img alt="" src="/images/home_conditional1_1.jpg"></p>

	<p>Bu tabloda Bölge Büyüme kolonu üzerinde, Bankadan büyükse yeşil bir tick 
	işareti olsun, sarı ve kırmızı ikon olmasın istiyoruz diyelim. Çünkü 
	yöneticimiz "çok renkli olunca kafam karışıyor, sadece iyileri görmek 
	yeterli" demiş olsun. Bunun için <strong>Conditional Formatting&gt;Icon 
	Sets&gt;Indicators</strong>'ten ilk seçenek tıklandığında karşımıza aşağıdaki ekran 
	gelir.</p>

	<img alt="" src="/images/home_conditional1_2.jpg">

<p>	Bu ekranda aşağıdaki gibi sadece ilk kutu için ilgili hücreyi gireriz ve Type kutusuna da Number gireriz. Diğer iki kutuyu "No Cell Icon" yaparız.</p>

	<img alt="" src="/images/home_conditional1_3.jpg">

<p>OK dedikten sonra da "Applies to" kutusunu da hangi hücrelere 	uygulayacaksak 
onları seçeriz. </p>

	<img alt="" src="/images/home_conditional1_4.jpg">

	<p>Ancak iki üstteki resimde farkettiyseniz yeşil tick işareti koyduğumuz 
	alan için Value kutusuna $I$6 şeklinde mutlak başvuru girişi yaptık, zira 
	göreceli başvuryu(I6) Icon Setlerde girmeye izin vermiyor. Bu nedenle, Conditional kutusunu kapattıktan sonra ilgili hücrelere tek tek girip bu 
	Value alanını değiştirmemiz lazım. Ör:H7 hücresine gidip bunun için Value 
	alanını $I$7 yapmak gibi. Ancak değiştireceğiniz hücre sayısı çoksa ve 
	bunları tek tek değiştirmek istemiyorsanız, aşağıdaki VBA kodu ile de bunu 
	başarabilirsiniz(Makro bilmenin avantajları...)</p>
	
<pre class="brush:vb">
Sub conditional_toplu()
'bu makroyu recorder ile kaydettim, sadece döngüyü ve Value kısmını elle değiştirdim
For Each c In Selection
    c.Select
    c.FormatConditions.Delete 'mevcuttaki conditionı silelim

    c.FormatConditions.AddIconSetCondition
    c.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    c.FormatConditions(1).IconCriteria(1).Icon = xlIconGreenCheckSymbol
    With c.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueFormula
        .Value = "=" & c.Offset(0, -2).Address 'bu kısmı kendim değiştirdim
        .Operator = 7
        .Icon = xlIconNoCellIcon
    End With
    With c.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 1000000000000#
        .Operator = 7
        .Icon = xlIconNoCellIcon
    End With
Next c

End Sub
</pre>
<p>
Sonuç aşağıdaki gibi olacaktır.</p>
	<p>
	<img alt="" src="/images/home_conditional1_5.jpg"></p>	
	
</div>	



<h2 class='baslik'>İki kritere göre iki farklı format uygulamak</h2>
<div class='konu'>

<p>	İkinci örneğimiz ilk örneğin hemen hemen aynısı, ancak bu sefer bir de pazar 
	büyümesi ile karşılaştırma kolonu var. Bundan sonraki kısım teknik bilgiden 
	ziyade, çözüm bulma beceresi ve biraz pratiklikle alakalı. Giriş sayfamda belirttiğim gibi bu sitede bu tür 
	çözüm önerilerini de buluyor olacaksınız. Şimdi diyelim ki yöneticiniz, bölgeyi banka ve pazarla karşılaştırmanızı 
	istemiş olsun. Bir hücrede aynı anda hem bölgeye hem pazara göre <span style="text-decoration: underline">aynı 
	türde</span> conditional formatting uygulayamıyoruz. Yani aynı anda hem 
	bankaya hem pazara göre icon set formatı uygulayamayız. Bu durumda önünüzde iki seçenek(benim 
	aklıma gelen) var,</p> 
	
        <p><strong>1. seçenek: </strong>Bölge kolonunda, bankaya göre karşılaştırma için icon set 
	uygularken, pazara göre karşılaştırma için farklı bir formating 
	uygulanabilir(arkaplan rengini değiştirmek gibi)</p>

	<p><strong>2.seçenek: </strong>Conditional formattingi bölge kolonuna değil, banka ve pazar 
	kolonlarına uygulamak. Bölge, bankadan iyiyise banka kolonuna yeşil tick, 
	bölge pazardan iyiyse pazar kolonuna yeşil tick konur.</p>

<p>Ben örnek olsun diye ikisini birden uyguladım, tercih size kalmış. Tabiki ikisini aynı anda kullanmak anlamsız olur, ya birini ya 
	diğerini seçmeniz gerekecektir.</p>
	<img alt="" src="/images/home_conditional2_1.jpg">

</div>

<h2 class="baslik">Formüllü CF Örnekleri</h2>
<div class="konu">
	<p>CF'i, girdiğimiz bir formülün doğru olması durumunda da uygulayabiliriz. 
	Yani girdiğimiz formül TRUE/FALSE döndürmelidir. Mesela aşağıdaki formül ile 
	Sayı olan tüm hücreler formatlanacaktır.</p>
	<p>
	<img src="/images/HomeCF16.jpg"></p>
	<p>
	Aşağıda diğer formüllü örnekleri bulabilirsiniz. Ancak yukarda belirttiğimiz gibi mutlak ve  göreceli başvuru tiplerinin ne olduklarını iyice içselleştirdikten sonra devam etmenizi öneririm. Zira bu örneklerin hepsinde mutlak ve göreceli başvuruların farklı kombinasyonları kullanılacak olup formüller biraz karmaşıklaşacaktır.</p>

<h4 class="baslik">Mükerrer kayıtları bulma</h4>
<div>
<p>Standart CF seçenekleri arasında "Duplicate Values" olduğunu biliyoruz. 
Bununla mükerrer <strong>kayıtları</strong>(tüm satırı aynı) değil, sadece seçilen kolondaki 
mükerrer <strong>datayı</strong> bulmuş oluyoruz. Halbuki veritabanı dilinde 
<span style="text-decoration: underline"><strong>Kayıt</strong></span> demek, 
tüm bir satıra ait data demek. Örneğin aşağıdaki kayıtlar mükerrer iken,</p>
	<p><img src="/images/CFduplike1.jpg"></p>
<p>Şu <strong>kayıtlar</strong> mükerrer değildir. (Sadece Müşteri No <strong>
datası</strong> mükerrerdir.)</p>

	<p><img src="/images/CFduplike2.jpg"></p>
	<p>Peki mükerrer kayıtları tüm satır şeklinde nasıl buluruz. Tabiki bir 
	formülle. Hemen bakalım.</p>
	<p><img src="/images/homecfmuk1.jpg"></p>
	<p>Countifs'in 1. ve 3. parametresini mutlak başvurulu yaptım, sürekli aynı 
	yerde&nbsp; baksın diye. 2.ve 4.parametresini ise yarı mutlak yarı göreceli yaptım, kayarak 
	ilerlesin diye, ama sadece satırda kaysın, sürunda değil. Böylece örneğin 5.satırı kontrol ederken A2:A17 arasında 
	14023ten bi tane daha varmı diye bakacak ve aynı anda B2:B17 arasında 
	14023ün Kredi Kartı bi tane daha varmı diye bakmış olacak. </p>
	<p>Tüm satır renklendirmesi istediğim için tüm alanı seçiyorum.</p>
	<p><img src="/images/homecfmuk2.jpg"></p>
	<p>Gördüğünüz gibi en alttaki 18952 nolu müşteri de mükerrer olmasına rağmen 
	iki farklı ürün için datası bulunduğu için kayıt anlamında mükerrer değildir. 
	Sadece 15375 nolu müşteri mükerrer çımıştır.</p>
	<p><img src="/images/homecfmuk3.jpg"></p>
	<p>NOT:C kolonundaki alan hiç değişmediği için bunu formüle dahil etmedim.</p>
</div>


<h4 class="baslik">Mükerrer kayıtlardan sadece ilkini formatlama</h4>
<div>
	<p>Eğer mükerrer kayıtların tamamıyla değil de sadece ilkiyle 
	ilgileniyorsak, tarama sırasında kontrol bölgesinin de kayarak ilerlemesini 
	sağlamalıyız. Mesela aşağıdaki görüntüde 10.satırın tarama bölgesi A10:A25 
	olur. Neden? Çünkü A2:A17 olarak başladık, ama satırı göreceli verdik. 
	Dolayısıyla aşağı doğru indikçe satırlar da 1er 1er kayar. Yani ilk kayıt 
	olan 11670 için A2:A17'ye bakarken ikincisi için A3:A18'e bakar ve böylece 
	gider. Dolayısıyla mükerrer kaydı bulduktan sonra bi aşağıya kaydığı için o 
	kayıttan bidaha bulamamış olur. Bu örnekte 11.satırdaki 15375 için tarama 
	bölgesi A11:A26 olduğu için onun aynısına rastlayamamıştır ve 
	işaretlememiştir.</p>

<p><img src="/images/homecfmuk4.jpg"></p>
</div>

<h4 class="baslik">Mükerrer kayıtlardan ilki dışındakileri formatlama</h4>
<div>
<p>Eğer mükerrer kayıtlardan, ilki dışındakilerle ilgilenmiyorsak bundaki tarama mantığı "İlk başı sabitle, sonra hücrenin kendisyle birlikte aşağı doğru kay" şeklinde olmalı. Yani ilk kayıt olan 11670'i A2:A2de arar, ikinci kayıt olan 11693ü A2:A3te, 
böyle ilerlerken 10.kayda geldiğinde 15375i A2:A10'da aradı ve 1 tane buldu, 
dolayısıyla işaretlemedi, 11.kayda geldiğinde 15375i A2:A11de aradı ve 2 tane 
bulduğu için bunu işaretledi. Aşağılarda 3. bi eşleşme olsaydı onu da 
işaretlerdi.</p>
	<p><img src="/images/homecfmuk5.jpg"></p>
</div>

<h4 class="baslik">Belli bir kategorideki rakamların toplamı negatif/poziti mi kontrolü</h4>
<div>
<p>Şimdi diyelim ki şöyle bir listemiz var. </p>
	<p>
	<img src="/images/homeCFornek50.jpg"></p>
<p>Bu listede İndirim+İlave toplamı negatif mi diye kontrol edelim, toplamı negatif olan kayıtları işaretleyelim. 
Tabi bunun yapmanın bir yolu da PivotTable yapıp, küçükten büyüğe sıralayıp en üstte negatif var mı diye bakmak olabilir, 
veya başka yöntemler de kullanılabilir. Her zaman dediğimiz gibi, o an ihtiyacınızı hangi yöntem çözüyorsa onu kullanın. Biz şimdi
CF yöntemini kullanacağız.</p>

<p>Bunun için formüle dayalı CF uygular ve şu formülü yazarız:</p>
<img src="/images/homeCFornek51.jpg">

<p>ve A2:C17 alanına uygularız:</p>
<img src="/images/homeCFornek52.jpg">

<p>Sonuç:</p>	

<p><img src="/images/homeCFornek53.jpg"></p>
</div>


	<h4 class="baslik">Bölge ortalamasının altında kalan şubeleri formatlama</h4>
<div>
<p>Şimdi diyelim ki şöyle bir listemiz var. Burada bölgenin ortalama HG% 
oranından düşük HG%'li şubeleri renklendirmek istiyoruz ama şubenin HG%si 
%100'ün üzerinde kalıyorsa renklenmesin. Mesela Bölge1'in HG% ortalaması %104 
olup Şube2'nin renklenmemesi gerekiyor.</p>
	<p>
	<img src="/images/cfavgif1.jpg"><p>
	Formülümüzü aşağıdaki gibi yazıp ilerleyelim.<p>
&nbsp;<img src="/images/cfavgif2.jpg"><p>
	Uygulanacak alanı da seçelim.<p>
	<img src="/images/cfavgif3.jpg">&nbsp;	

<p>Sonuç aşağıdaki gibi olacaktır.</p>
	<p><img src="/images/cfavgif4.jpg"></p>
</div>



</div>
</asp:Content>
