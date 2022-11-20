<%@ Page Title='InsertMenusu Grafikler' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Insert Menüsü'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Grafikler</h1>
<p>Grafikler elimizdeki ham veriyi görselleştirme ve daha hızlı anlayıp yorumlamamızı sağlayan en güçlü araçlardandır. 
Hatta Excelle o kadar bütünleşmiştir ki, birçok Excel kitabı/sitesi logosunda 
veya giriş görselinde bir grafik kullanır.</p>
	<p>Bu kadar bilindik oldukları 
ve bu siteyi inceleyenlerin temel grafik bilgilerine sahip olduklarını 
düşündüğüm için bunlar hakkında genel bilgileri burada vermeyeceğim. Onun yerine 
birkaç püf noktası ve önemli birkaç grafik türüne değineceğim.</p>
	<p>Sayfa boyunca işlediğim örnekleri içeren dosyayı
	<a href="../../Ornek_dosyalar/grafikler.xlsx">buradan</a> indirebilirsiniz.</p>

<h2 class="baslik">Şık Grafikler için ipuçları</h2>
<div class="konu">
<p>
Bu ilk kısımda gerek kendi deneyimlerimden, gerekse grafik gurularının 
tavsiyelerinden ortaya çıkan tavsiyeleri bulacaksınız. </p>
	<p>
	Grafikleri, elmizdeki datanın dış dünyaya açılan kapısı olarak görebiliriz. 
	Bu nedenle onları ihmal etmemek ve gereken önemi 
vermek lazım gelir. Zira binbir zorlukla ürettiğiniz(belki uzun bir SQL yazdınız, 
Excel içine gömdünüz, Excelde bunu otomatize etmek için yine uzunca bir makro 
yazdınız v.s) datayı zayıf bir şekilde sunarsanız emekleriniz heba olabilir. 
	Özetle işin pazarlamasını iyi yapmamız gerekiyor. Bu, aynı zamanda 
	insanların sizi sadece "Datacı" etiketiyle yaftalamalarını da engellemiş 
	olur :)</p>
	<p>
	Şimdi tavsiyelere bakalım.</p>
	<h3>Doğru grafik türü</h3>
	<p>Öncelikle doğru grafik türünü seçerek başlayalım. Bu, çoğu durumda çubuk veya çizgi grafik olacaktır. Kategorilerin 
	birbirine göre durumu daha çok çubuk grafikle gösterilirken trendler ise 
	daha çok çizgi grafik olarak gösterilmektedir.</p>
	<p>Aşağıda bölgelerin hacimsel grafiği bulunmakta.</p>
	<p><img src="/images/insertgrafik20.jpg"></p>
	<p>Trend örneği ise şöyle olacaktır.</p>
	<p><img src="/images/insertgrafik21.jpg"></p>
	<p>Ancak her trend grafiği illa çizgi grafik olmak zorunda değil. Bu örnekte 
	sürekli üzerine eklenen(arada istisnalar olabilir, sezonsallık v.s nedeniyle 
	düşüş) türde bir trend sözkonusu. Ancak aşağıdaki gibi günlük satış trendi 
	dalgalanan bir yapıda olabileceği için bunu çubuk olarak izlemek daha uygun 
	olacaktır. Zira buna bir de aşağıda göreceğimiz gibi ortalama çizgisi de 
	eklemek istenmesi muhtemeldir, iki çizgi de genelde çok şık durmaz.</p>
	<p><img src="/images/insertgrafik22.jpg"></p>
	<p>Bazen de Çubuk ve Çizginin bir kombinasyonu da gösterilebilir. 
	<a href="#combografik">Aşağıda</a> 
	bunun nasıl yapıldığı ayrıca gösterilecektir.</p>
	<p>Eğer karşılaştırma yapacağınız bir data kümesi varsa ve bu küme çok 
	kalabalık değilse, ve bunlardan özellikle biri diğerlerinden önemli ölçüde 
	büyükse Pie grafikler de seçilebilir.<br></p>
	<img src="/images/insertgrafik1.jpg"><h3>Sıralama şekli</h3>
	<p>Eğer, kategorideki eleman sayısı çok fazlaysa alfabetik gösterim yerine 
	rakama göre sıralanmış şekilde gösterim daha uygun olabilir.</p>
	<p><img src="/images/insertgrafik2.jpg"></p>
	<h3>Arkaplan çizgilerini kaldırın</h3>
	<p>Arkaplan çizgileri(Gridlines) çok zaruri değilse kaldıralım.</p>
	<p><img src="/images/insertgrafik3.jpg"></p>
	<p>Grafiklerimiz bu yukardaki kadar basit olmayacaktır, bunlara ortalama 
	çizigisi, trend çizgisi gibi diğer unsurlar da eklendiğinde arkaplan 
	çizgileri iyice kalabalık olur, o zaman mutlaka kaldırın.</p>
	<h3>Aşırı formatlamadan uzak durun</h3>
	<p>Aşırı renklendirme, arkaplan deseni ekleme, 3D efektleri abartma gibi 
	yolalra başvurmayın. Bunları dozunda 
	kullanmak iyi bir fikirken aşırı kullanım grafiğinizi, makyajı abartmış 
	kokoş bir kadına benzetebilir. Özetle, Grafiğinizin güzelliğini aşırı 
	makyajla berbat etmeyin.</p>
	<p>Aşağıdaki gibi orta karar bir formatlama yeterli olacaktır.</p>
	<p><img src="/images/insertgrafi4.jpg"></p>
	<p>Öne çıkarmak istediğiniz grafiğin rengini ve ağırlığını da daha çarpıcı 
	hale getirebilirsiniz. Aşağıda iki eksenli grafikteki gibi.(Farkettiyseniz 
	ben kırmızı çizgiyi biraz abarttım, işte bu aşırı makyajdır, bundan bir ton 
	daha ince olabilir)</p>
	<p><img src="/images/insertgrafi5.jpg"></p>
	<h3>Legend kullanmalı mı kullanmamalı mı</h3>
	<p>Tek boyutlu bir grafikse kesinlikle kullanmayın. Gerekirse Başlık içinde 
	boyut ismini geçirin. İki boyutlu grafiklerde özellikle combo(2 eksenli) 
	grafikse neyin ne olduğu belliyse yine Legend kullanmanıza gerek yoktur.</p>
	<p>Mesela aşağıdaki grafikte sol eksen büyük rakamlara aittir ve bunlar 
	kredi kullandırım tutarını gösterirken, sağ eksen kullandırım adedini 
	gösterir ve çizgi ile gösterilir. Genelde Büyük rakamlar çubukla, küçük 
	rakamlar çizgi ile gösterilebilir.(Bu bir norm değildir, sadece bizim 
	kurumdaki genel kullanım şeklidir, sizin dünyanızda tersi durum da olabilir)</p>
	<p><img src="/images/insertgrafik6.jpg"></p>
	<h3>Gerekli yerlerde not kutuları kullanın</h3>
	<p>Soru işareti olabilecek yerlerde(aşırı dalgalanma gibi) gerekli açıklamaları mümkünse grafik 
	üzerinde gösterin. Böyle birden fazla durum varsa grafik üzerinden göstermek 
	yerine, grafik altında dipnot olarak gösterilebilir.</p>
	<p><img src="/images/insertgrafik7.jpg"></p>
	<h3>Ölçeklendirmeye dikkat edin</h3>
	<p>Aşağıdaki grafiğe baktığımızda sanki günler arasında çok dalgalanma 
	varmış gibi görünüyor. Bu, dataya çubuk grafik şablonu uyguladığımda çıkan 
	otomatik grafikti. Y eksenine bakacak olursanız en küçük değer 2.940.000'tan 
	başlıyor.</p>
	<p><img src="/images/insertgrafik8.jpg"></p>
	<p>Halbuki başlangıç noktasını 0'a çekersek o kadar da bir dalgalanma 
	olmadığını görebiliriz.</p>
	<p><img src="/images/insertgrafik9.jpg"></p>
	<p>Hangi durumlarda başlangıç değerinin 0, hangi durumlarda ise daha yukarda 
	bir değer olacağını iyi netleştirin.</p>
	<h3>İnsanların kafasını yana yatırtmayın</h3>
	<p>Eksenlerdeki kategori isimlerinin uzun olduğu durumlarda eğimli gösterim 
	şeklini tercih etmeyin. Mesela aşağıda grafik çok şık görünmüyor.</p>
	<p><img src="/images/insertgrafik10.jpg"></p>
	<p>
	Alternatif olarak şunu da denemeyin(Her ne kadar yukarıdaki bazı örneklerde 
	bu şekilde kullanmış olsam da)</p>
	<p>
	<img src="/images/insertgrafik11.jpg"></p>
	<p>
	Onun yerine biraz daha basit düşünüp grafiği yatay çubuklar şekline 
	dönüştürebilirsiniz. Duruma göre Axis'te <strong>Dates in reverse order</strong> 
	seçeneğini işaretlemeniz gerekebilir. </p>
	<p>
	Tabi bu seçeneği kullandığınızda da grafiğin aşağı doğru çok fazla 
	uzamamasına dikkat etmelisiniz.</p>
	<p>
	<img src="/images/insertgrafik12.jpg"></p>
</div>

<h2 class="baslik">Grafik Örnekleri</h2>
<div class="konu">
<h3 id="combografik">İki eksenli grafik çizimi</h3>
	<p>İki eksenli grafik yapmak 2013 öncesi versiyona kadar biraz 
	uğraştırıcıydı. Hani ilk bakışta nasıl yapılacağını bilemediğniz, mulaka 
	araştırmanız gereken konulardan biriydi. Ama sağolsun Microsofttaki 
	geliştiriciler her yeni sürümde işleri biraz daha basitleştiriyor ve ilave 
	araştırma yapmadan kullanılabilecek tarzda araçlar geliştiriyorlar. 2 
	eksenli grafikler de bulardan biri. Bunun için tek yapmanız gereken mevcut 
	alışkanlıklarınızın dışına çıkmak ve "ne yenilikler varmış" diye yeni 
	versiyonu şöyle bi kurcalamak. </p>
	<p>Excel 2013 kullandığı halde aşağıdaki bu grafik türünü (2 eksenli)&nbsp; 
	görmeyen veya görüp de içine bakmayan çok kişi olduğuna eminim, hatta 
	bazılarını tanıyorum bile :)</p>
	<p><img src="/images/insertgrafikcombo.jpg"></p>
	<p>Şimdi şöyle bir data kümemiz olsun.</p>
	<p><img src="/images/insertgrafikcombo2.jpg"></p>
	<p>Tek yapmanız gereken hangi alanın çubuk hangisinin çizgi(veya başka 
	kombinasyonlar da olabilir, ama en yaygını budur) ve bunlardan hangisinin 
	2.eksen olacağını seçmek.</p>
	<p><img src="/images/insertgrafikcombo3.jpg"></p>
	<p>Sonuç aşağıdaki gibi olacaktır.</p>
	<p><img src="/images/insertgrafikcombo4.jpg"></p>
	<p>İki eksen geldiğinde grafiğin orta alanı iyice daralır, o yüzden 
	eksenleri uygun metrik dilimde göstermek akıllıca olacaktır. Mesela bu 
	örnekte 1.ekseni milyonlar halinde gösterelim.</p>
	<p>Bunun için bu eksene sağ tıklayıp <strong>Format Axis</strong> diyoruz ve
	<strong>Display Units</strong> bölümünü Millions olarak değiştiriyoruz.</p>
	<p><img src="/images/insertgrafikcombo5.jpg"></p>
	<p>Ve sonuç:</p>
	<p><img src="/images/insertgrafikcombo6.jpg"></p>
	<p>Not:İsterseniz sağ ekseni de 1000ler şeklinde gösterebilirsiniz.</p>
	<h3>Ortalama çizgisi ekleme</h3>
	<p>Microsoft geliştiricileri şüphesiz her yeni versiyonda güzel şeyler 
	ekliyorlar ancak anlayamadığım şekilde grafiklere ortalama çizgisi ekleme 
	seçeneğini hala yapmadılar. Bunun için kendi alternatif yöntemlerimizi 
	geliştirmemiz gerekiyor. </p>
	<p>Ortalama çizgisini tek eksenli grafiklerde ekleyebileceğimiz gibi çift 
	eksenlilere de ekenebilir. Gelin biz bir üstteki çift eksenli grafiğe 
	ekleyelim.</p>
	<p>Yapacağımız şey, datamıza yeni bir kolon ekleyip ilgili sütunun 
	ortalamasını almak olacak.</p>
	<p><img src="/images/insertgrafikcombo7.jpg"></p>
	<p>Sonra grafiğimizin kaynak datasını D kolonuna uzayacak şekilde 
	genişletelim.</p>
	<p><img src="/images/insertgrafikcombo9.jpg"></p>
	<p>O da ne! İsteğimiz şey olmadı. Tamam, Ort Tutar kolonu istediğimiz gibi 
	çizgi şeklinde geldi ama yanlış eksende duruyor.</p>
	<p><img src="/images/insertgrafikcombo10.jpg"></p>
	<p><strong>Change Chart Type </strong>diyoruz. Ortalama Tutarımız, eksen 
	türü olarak 2.eksende seçiliydi, bundaki tick işaretini kaldırıyoruz. Bu 
	kadar.</p>
	<p><img src="/images/insertgrafikcombo11.jpg"></p>
	<p>Şimdi tamamdır.</p>
	<p><img src="/images/insertgrafikcombo12.jpg"></p>
	<p>Çizeceğimiz çizgi illa ortalama olmak zorunda değildir. Belirlenen herhangi bir 
	eşik değer 
	de grafiğe aynı şekilde eklenebilir. Mesela günlük 3 mio Kullandırım altında 
	kalınan günleri görmek isterseniz, Ortalama formülü yazdığınız kolona direkt 
	3.000.000 yazabilrsiniz. Veya hem ortalama hem de eşik değeri ikisni de 
	kullanabilirsiniz.</p>
	<p><img src="/images/insertgrafikcombo13.jpg"></p>
	<h3>Pareto Grafiği ile kümülatif toplam görme(2016)</h3>
	<p>Bir başka faydalı geliştirme örneği daha. 2016 öncesinde bir pareto 
	grafiği çizmek için datayı sıralamak ve ilgili datanın yanında iki yardımcı 
	kolon açmak&nbsp; gerekiyordu. Biri kümülatif toplamı, diğeri kümülatif 
	yüzdeyi göseren iki kolon. Sonra bunları iki eksenli grafik haline getirmek 
	gerekiyordu. 2016 versiyonunda bu grafiği tek seferde çizdirebiliyoruz.</p>
	<p>Bu arada pareto grafiği de nedir derseniz kısa bir açıklama yapalım, daha 
	detaylı bilgiyi küçük bir google araştırmasıyla bulabilirsiniz. Pareto 
	teorisine göre herhangi bir aktivitede bir sonucun %80sini, o aktiviteye 
	katılanların %20si oluşturur. Mesela bankacılık örneğinde gidersek, bir 
	şubenini müşterilerinin %20si, o şubenin karının %80ini oluşturur. (Bilgi:bu 
	durum o şube için iyi bişrşey değildir, zira müşteriler tabana yaygın 
	durumda değildir. En büyük müşterinin çıkmasıyla şubenin karlılığı altüst 
	olabilir.) Perakende tarafından örnek verecek olursak, bir marketler 
	zincirinin karının %80ini, marketlerin %20si oluşturmaktadır. Tabi bu 80-20 
	kuralı işin teorisi. Pratikteki durumu görmek için bu grafiği çizdirmemiz gerekiyor. Ayrıca banka şubesi örneğinde olduğu gibi çıkan 
	değerin iyi mi kötü mü olduğunu ayrıca yorumlamak gerekiyor. </p>
	<p>Şimdi önce datamıza bakalım. Bir banka bölge müdürlüğünün şubelerinin 
	belirli bir kalemdeki rakamları aşağıda gibi olsun. </p>
	<p><img src="/images/insertgrafikpareto1.jpg"></p>
	<p><strong>Recommended Charts</strong> içinde en altta Paretoyu seçiyorum.</p>
	<p><img src="/images/insertgrafikpareto2.jpg"></p>
	<p>Bu grafik türüne <strong>All charts </strong>içinden, <strong>Histogram</strong> 
	türündeki grafiklerden ikincisini seçerek de ulaşabilirsiniz.</p>
	<p><img src="/images/insertgrafikpareto3.jpg"></p>
	<p>Grafiğimiz Önizlemedeki gibi aynen gelir. Farkettiyseniz şubeleri boy 
	sırasına dizmediğim halde grafiğimiz boy sırasına göre geldi.</p>
	<p><img src="/images/insertgrafikpareto4.jpg"></p>
	<p>Çizgi olarak görünen şey, kümülatif yüzdesel değerleri ifade eder ve sağ 
	eksen üzerinden okunur. Burada %80lik dilime gelen grubu görebilirsiniz. 
	Ancak spesifik olarak tam neresi %80in altında kalıyor diye görmek 
	isterseniz bunun için eski yöntemle Pareto grafiği çizmemiz gerekiyor. Zira 
	2016 sürümünde malesef %80 çizgisi ekleme diye bir seçenek yok. Sanırım bunu 
	da bir sonraki versiyona sakladılar. </p>
	<p>Şimdi bir de eski yöntemle grafiğimizi oluşturalım, böylece 2016 
	versiyonu kullanmayanlara da Pareto grafiğini çizme yöntemini göstermiş 
	olalım.</p>
	<p>İlk olarak datamızı boy sırasına sokalım ve sonra Kümülatif Toplam ile 
	Kümülatif Yüzde kolonlarını oluşturalım.</p>
	<p><img src="/images/insertgrafikpareto5.jpg"></p>
	<p>Datamız hazır olduğuna göre şimdi Combo grafik hazırlayabiliriz. 
	Kümülatif Hacim datasına grafikte ihtiyacımız yok, ister bu kolonu gizleyip 
	grafiğiniz oluşturun, isterseniz grafik oluştuğunda bunun çubuklarına 
	tıklayıp silin.&nbsp;(Tabiki Kümülatif yüzde kolonunu tek bir formülle de 
	yazabilir ve Kümülatif yüzde kolonuna hiç ihtiyaç duymayabilirdiniz)</p>
	<p><img src="/images/insertgrafikpareto6.jpg"></p>
	<p>Son olarak bu grafiğe bir de eşik değer olarak %80 kolonunu ekleyeceğiz. 
	Bunu da datamıza yeni kolon ekleyerek yapacağız. Sonra da Grafiğimizi seçip
	<strong>Select Data</strong> der ve yeni alanımızı seçeriz(Gerekmesi 
	durumunda eksenleri tekrar ayarlarız)</p>
	<p><img src="/images/insertgrafikpareto7.jpg"></p>
	<p>Ve işte grafiğimizin nihai hali aşağıdaki gibidir. Gördüğünüz gibi 
	5.şubeden itibaren %80lik dilimi yakalıyoruz. Yani bu durumda şubelerin 
	%41'i(5/12si) toplam hacmin %80ini üretiyormuş.</p>
	<p><img src="/images/insertgrafikpareto8.jpg"></p>
	<h3>Stacked Column/Bar ile değerleri sağlı sollu(aşağı yukarı) görme</h3>
	<p>Aşağıdaki gibi biri pozitif diğeri negatif olan iki veri kümeniz olsun. 
	Bunları birarada göstermenin en iyi yöntemi <strong>Stacked</strong> grafik 
	türleridir.</p>
	<p><img src="/images/insertgrafikstacked1.jpg"></p>
	<p>Recommended Charts altında da çıkan olan Stacked Bars'a ilk önerimiz 
	olarak bakalım.</p>
	<p><img src="/images/insertgrafikstacked2.jpg"></p>
	<p>Bölge isimlerini en solda görmek için bu ekseni seçip <strong>Format Axis</strong> diyelim 
	ve Label Position olarak Low ayarlayım. </p>
	<p>&nbsp;<img src="/images/insertgrafikstacked3.jpg"></p>
	<p>Bu arada bölgeleri Bölge1 yukarda olacak şekilde dizmek istersek, Format 
	Axis içinde <strong>Categories in Reverse Order</strong> seçeneğini 
	işaretleriz.</p>
	<p>Bunun Çubuk versiyonu da aynı şekilde kullanılabilir. Kategori sayısı 
	arttıkça Çubuk yerine Bar kullanımı daha uygun olacaktır, zira aşağıdakinde 
	olduğu gibi Bölge isimlerini tam okumak için kafalar hafif yana yatmak 
	durumunda kalıyor.</p>

       <p>Bu grafik türlerinin ikisinin de %100 versiyonları da var. Bu, eksenleri 
	mutlak değere göre değil, iki değerin mutlak büyüklüklerini %100 olacak 
	şekilde bölümlere ayırmak anlamına geliyor. Aşağıda Stacked Column'nın %100 
	versiyonu görünüyor.</p>
	<p><img src="/images/insertgrafikstacked4.jpg"></p>
	<p>Bir diğer alternatif de negatifleri pozitif yaptıktan sonra Stacked 
	grafik uygulamak. Benim tercihim, pozitif ve negatifin ayrı ayrı olduğu 
	versiyonlardır, zira iki pozitif olduğunda sanki beyin olayı anlamak için 
	ekstradan çabaya giriyor gibi geliyor bana. İşte bu da aşağıda.</p>
	<p><img src="/images/insertgrafikstacked5.jpg"></p>
	<p>İki pozitifli datayı Stacked %100 yapınca hepsinin uzunluğu sabit olduğu 
	için bölgelerin birbirine göre durumu daha iyi karşılaştırılabiliyor. Bu da 
	seçenekler arasında olabilir.</p>
	<p><img src="/images/insertgrafikstacked6.jpg"></p>
	<h3>XY(Scatter) grafiği ile dağılım yoğunluğunu görme</h3>
	<p>Aşağıdaki gibi bir data kümeniz var ve siz hem Hedef Gerçekleştirme 
	Oranı(HGO%) hem de geçen yıla göre artış oranı küçük bölgeleri yakın 
	izlemeye almak istiyorsunuz diyelim. Bununla beraber diğer 3 kombinasyonu da 
	görmek istiyorsunuz, kim nerede diye. Bunu basit bir IF formülü ile de 
	yapabilirsiniz, ancak bu yöntemde 2 soru karşımıza çıkar. 1- Hangi spesifik 
	değer altını/üstünü hangi kategoriye koyacaksınız, bunu seçmek zor olabilir 
	2-Seçtiniz diyelim, bunları görsel olarak görmenin rahatlığını elde 
	debilecek misiniz?Çok büyük ihtimalle hayır.</p>
	<p><img src="/images/insertgrafikxy1.jpg"></p>
	<p>İşte böyle durumlarda XY(Scatter) Grafikleri en uygun çözümü sunar. 
	Recommended Charts içinden seçelim kendisini.</p>
	<p><img src="/images/insertgrafikxy2.jpg"></p>
	<p>Bu haliyle çok yavan, yine de üç aşağı beş yukarı yoğunlaşma bölgelerini 
	görebiliyorsunuz, ama daha işimiz var, şimdi biraz bu grafiğe çeki düzen 
	verelim.</p>
	<p>Öncelikle her iki eksendeki min ve max noktalarını iyi belirleyelim ki, 
	boş alanlar gereksiz yer kaplamasın. Bunun için X eskenine sağ tıklayıp
	<strong>Axis Options</strong>'a girip <strong>Bounds&gt;Minimum </strong>
	değerini 0,6 yapıyorum. Bu arada arka plandaki grid çizgilerini 
	kaldırıyorum.</p>
	<p><img src="/images/insertgrafikxy3.jpg"></p>
	<p>Şimdi Y ekseninin X eksenini tam ortasından kesmesini sağlayacağım, aynı 
	zamanıda X ekseni de 0 çizgisinden değil, Y ekseninin min ve max 
	değerlerinin ortasından geçsin isityorum.(İlla orta noktalardan geçmek 
	zorunda değil, isterseniz belirlediğiniz referans değerlerden de geçebilir) </p>
	<p>Önce Y'yi halledelim: 0,6 ile 1,60'ın ortası 1,10. Şimdi burası biraz 
	karışık gelebilir. Y eksenini ayarlıyoruz dedik ama beklediğinizin aksine Y 
	eksenini değil <strong>X eksenini </strong>seçip <strong>Axis Options
	</strong>deriz ve Vertical(yani Y) eksenin geçtiği noktayı Automaticten Axis 
	Value=1,1'e değiştiririz.</p>
	<p>&nbsp;<img src="/images/insertgrafikxy4.jpg"></p>
	<p>Hemen arkasından Y eksenini seçip <strong>Label Position </strong>olarak
	<strong>Low </strong>deriz, ve Fill&amp;Line alt sekmesinden de Line formatını
	<strong>Solid Line </strong>seçip rengini de siyah olarak belirleyelim.</p>
	<p><img src="/images/insertgrafikxy5.jpg"></p>
	<p>ve bu haliyle grafiğimiz biraz daha istediğimiz şekle bürünmüş oldu.</p>
	<p><img src="/images/insertgrafikxy6.jpg"></p>
	<p>Şimdi de X eksenini biraz yükseltelim. -20 ile +40'ın ortası +10dur. Bu 
	sefer Y eksenini seçip, Eksenin geçtiği nokta olarak 0,1 değeri girilerek, X 
	eskeni için <strong>Label Position'ı Low </strong>ve <strong>Solid Black
	</strong>line seçimleri yapılır. Bu arada Y ekseni için max değeri %50den 
	%40a da çekelim, onu unutmuşuz, yukarda %10luk alan boş yere işgal olmasın.</p>
	<p>Son olarak otomatik gelen başlığı da daha anlamlı hale getirelim. </p>
	<p><img src="/images/insertgrafikxy7.jpg"></p>
	<p>Vee işlem tamamdır.</p>
	<p>Şimdi aklınıza şöyle bir soru gelebilir. Bu eksenlerin ikisi de yüzdesel, 
	bunlardan hangisi hangisi? Grafiğe <strong>Axis Title </strong>ekleyerek bu 
	sorunu çözebileceğiniz gibi grafik alanınızı daraltmamak adına akıl yürütme 
	yoluna da başvurabilirsiniz. Siz bilirsiniz ki HG% oranı genelde %100 
	etrafındadır ve asla negatif olamaz; artış oranları ise hem negatif olabilir 
	hem de göreceli daha düşük değerlerdir. Bu bilgiyle Y ekseninin artış 
	ekseni, X ekseninin de HGO ekseni olduğunu anlarsınız. Tabi bazı durumlarda 
	bir eksen mutlak değer bir eksen yüzdesel olur, ki böye durumlarda neyin ne 
	olduğu zaten bellidir.</p>
	<p>Buraya istenise 4 adet textbox konarak bölmelerin açıklaması ve önem 
	sırasına göre numaralandırılması yapılabilir.</p>
	<p><img src="/images/insertgrafikxy8.jpg"></p>
	<p>Peki hangi yuvarlak nokta hangi Bölgeyi ifade ediyor. İşte bu sorunun 
	cevabı Excel 2013e kadar verilemiyordu. 2013 ile birlikte bu sorun da 
	çözüldü. Bunun için öncelikle grafiğimize Data Labels eklememiz lazım.</p>
	<p><img src="/images/insertgrafikxy9.jpg"></p>
	<p>Sonrasında bu Data Labellara tıklarayarak, Value From Cells seçimi 
	yapılır, istenirse Y değerlerinin gösterilmemesi de sağlanır, zira ne kadar 
	çok veri o kadar karmaşa demektir.</p>
	<p><img src="/images/insertgrafikxy10.jpg"></p>
	<p>Alan olarak bölge isimlerinin olduğu yer seçilir.</p>
	<p><img src="/images/insertgrafikxy11.jpg"></p>
	<p>Nihai grafiğimiz aşağıdaki gibi olacaktır</p>
	<p><img src="/images/insertgrafikxy12.jpg"></p>
	<h4>Dinamik Seçimler</h4>
	<p>Her bölgenin şubeleri için böyle bir çalışma yapmanız gerekti diyelim. 
	İşler o zaman biraz karmaşıklaşır. Bi kere her bölgenin min/max değerleri 
	farklı olacağı için her bölge seçiminde ona göre min max ve orta noktaların 
	ayarlanması için VBA kullanarak makro yazımı gerekir. Bu arada esas soru 
	seçimin ve seçilen bölgeye göre şubelerin gelme işinin nasıl yapılacağıdır. 
	Bu işlem, Data Validation ile bölge seçimi ve Dinamik Named Range ile&nbsp; 
	şubelerin listelenmesi şeklinde olabileceğini gibi, daha basit olarak Table 
	üzerinde Slicer kullanımı ile de sağlanabilir, ancak bunun için en az 2013 
	versiyonu gereklidir. Buna ait bir örneği
	<a href="../VBAMakro/Ileriseviyekonular_PivotTableChartveSlicernesneleri.aspx#OrnekUygulama">
	bu sayfada</a> bulabilirsiniz.</p>
	<h3>Treemap ve Sunburst ile hiyerarşik grafik oluşturma(2016)</h3>
	<p>Veri kümenizi hiyerarşik bir bakış açısıyla incelemek istiyorsanız 2016 
	ile gelen Treemap ve Sunburst grafik seçenekleri bu iş için uygundur. 
	Özellikle bir bölgenin belli bir kalemdeki en büyük bir iki şubesi kimlermiş 
	diye görmek istediğinizde ama bunu tek seferde tüm bölgeler için yapmak 
	istediğinizde idealdirler.</p>
	<h4>Treemap</h4>
	<p>Bu grafik türünde <strong>sadece 2 boyut</strong> bulunur:Ana ve alt kategori(Bölge ve 
	şube gibi). Alt kategoriler birbirlerinden bulunduklar dikdörtgenin 
	büyüklüklerine göre kolayca ayrıştırılabilir ve boy sırasına göre büyük 
	dörtgenlerden küçük dörtgenlere doğru dizilirler.</p>
	<p><img src="/images/insertgrafikhiyerarşik1.jpg"></p>
	<p><img src="/images/insertgrafikhiyerarşik2.jpg"></p>
	<p>Bölge isimlerini sol üstte görmek yerine başlık olarak da görmek 
	mümkündür. Bunun için grafiğe sağ tıklayın, <strong>Format Data Series
	</strong>deyin, <strong>Series options </strong>altında <strong>Label 
	Options'</strong>ı <strong>Banner </strong>olarak seçin. </p>
	<p><img src="/images/insertgrafikhiyerarşik3.jpg"></p>
	<p>Banner ilk başta gri idi, gri genelde pasif alanların rengi olduğu için 
	ben her bir bölgeyi seçerek ayrı ayrı başlıklarıyla birlikte renklendirdim.</p>
	<p>Değerleri de grafikte görmek isterseniz labellara sağ tıklayıp <strong>
	Format data label</strong> dedikten sonra <strong>Value </strong>seçeneğini 
	işaretleyin. Değerleri sadece alt kategoride gösterdiğine dikkat edin.</p>
	<p><img src="/images/insertgrafikhiyerarşik4.jpg"></p>
	<h4>Sunburst</h4>
	<p>Eğer, kategori altında birden fazla seviyede alt kategori varsa Treemap 
	yerine Sunburst kullanmakta fayda var, zira Treemap her zaman ana kategori 
	ile en sondaki alt kategoriyi ele alır, aradaki diğer seviyeleri eler. </p>
	<p>Data kümemiz şöyle olsun:</p>
	<p><img src="/images/insertgrafikhiyerarşik5.jpg"></p>
	<p>Sunburst uygulanınca;</p>
	<p><img src="/images/insertgrafikhiyerarşik6.jpg"></p>
	<p>Kategorileri tek tek seçerek odağınızı keskinleştirebilirsiniz.</p>
	<p><img src="/images/insertgrafikhiyerarşik7.jpg"></p>
	<p>Bir seviye daha inelim, Ürün1e tıkladım.</p>
	<p><img src="/images/insertgrafikhiyerarşik8.jpg"></p>
	<p>Bi tane daha inelim, Şube7'ye tıklayalım, ayrıca üzerinde bekleyerek 
	değerini de görelim.</p>
	<p><img src="/images/insertgrafikhiyerarşik9.jpg"></p>
	<p>Rakamları grafiğin içinde de gösterebilirsiniz ancak kategori sayısı çok 
	olduğu için zaten iyice sıkışk bir alanımız vardır, bunun yerine yukarıdaki 
	gibi alanların üzerine gelerek görme yöntemini tercih etmenizi öneririm. Ama 
	ille de grafikte görünsün isterseniz Treemaps'te yaptığımız gibi Labellara 
	sağ tıklayarak <strong>Value</strong> kutusunu işaretleriz. Yine Treemaps'te 
	olduğu gibi sadece en alt seviyede rakam gösterildiğine dikkat edin.</p>
	<p><img src="/images/insertgrafikhiyerarşik10.jpg"></p>
</div>
</asp:Content>
