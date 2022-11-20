<%@ Page Title='Tablo' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>

<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Home Menüsü'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Tablolar</h1>

<p>Tablolar(Eski adıyla Liste) gücünü ve önemini geç keşfettiğim araçlardan birisidir. 
Naçizane benim görüşüm MIS'le ilgilenen kişilerin bunları mutlaka bilmesi ve 
kullanması gerektiğidir.</p>

<p>Tablolar belirli bir dış veri kaynağını içeriyor olabilecekleri gibi düz listeleri de 
Tablo haline getirebilir ve Tabloların gücünden faydalanabiliriz. Biz burada düz 
listelerin nasıl Table haline getirileceğini göreceğiz. Dış veri kaynaklarını
<a href="DataMenusu_BaskaVeriKaynaklariilecalismak.aspx">başka bir bölümde</a> ele alıyor olacağız.</p>
	<h2>Nasıl oluşturulur</h2>
	<p>Home menüsünden <strong>Format as Table </strong>düğmesi ile bir listeyi çok sayıdaki 
	formatlı tablolardan birine dönüştürebilirsiniz. Aşağıdaki resimde bunlardan 
	bazılarını görebilirsiniz.</p>
	<p><img src="/images/inserttable1.jpg"></p>
	<p>Gördüğünüz üzere tablo formatları genelde zebra desenlidir, yani bi açık 
	bi koyu. Ve işin güzel tarafı bu format, yeni data eklense veya aradaki 
	kayıtlar silinse dahi korunmaya devam eder. Mesela ben aşağıdaki tablodan 
	Anadolu2 bölgesini silersem Başkent1'in olduğu altıncı satır otomatikman 
	beyaz ve sonrakiler de bi mavi bi beyaz olarak güncellenir. Manuel zebra 
	deseni uygulanmış bir 
	listede ise Anadolu1 de Başkent1 de mavi olacak ve desen bozulmuş 
	olacaktı.</p>
	<p>
	<img src="/images/inserttable3.jpg"></p>
	<p>Bu da Anadolu2'nin silinmiş halidir.</p>
	<p>
	<img src="/images/inserttable4.jpg"></p>
	<p>Data alanını Table haline getirmeden önce kontrol edilmesi gereken bazı 
	noktalar bulunmaktadır.</p>
	<ul>
		<li>Listenin ilk satırında benzersiz kolon başlıkları olmalıdır ve boş 
		hücre olmamasına dikkat edilmelidir.</li>
		<li>Listede tamamen boş bir satır olmamasına özen gösterilmelidir.</li>
		<li>Liste, sayfadaki başka veri kümelerinden en az bir boş satır ve bir 
		boş kolon ile ayrı olmalıdır.</li>
	</ul>
	<h2>Ne işe yarar?</h2>
	<p>Tabloların normal listelere göre 5 ana üstünlüğünden bahsedebiliriz. 
	Bunlar;</p>
	<ul>
		<li>Format avantajı</li>
		<li>Name davranışı</li>
		<li>Pivot(Özet) Tablo kaynağı olması</li>
		<li>Slicer kullanımı</li>
		<li>Formül kolaylığı</li>
	</ul>
	<p>Bunlardan Format faydasını yukarda görmüştük. Şimdi diğer 4 maddeye 
	bakalım.</p>
	<h3>Name davranışı</h3>
	<p>Tablolar, sadece belirli bir veri kümesinin formatlanması demek değildir. Aynı 
	zamanda özel bir alan oldukları için bir Named Range gibi davranırlar ve birçok 
	durumda(VBA dahil) <strong>Name</strong> olma avantajlarını kullanırlar.</p>
	<p>Mesela aşağıdaki formülle bu tabloda ne kadar kayıt olduğunu 
	saydırabilirsiniz. (NOT:Table isimleri sadece DSUM gibi Veritabanı fonksiyonlarında kullanılamazlar)</p>
	<pre class="formul">=ROWS(Table1)</pre>
	<p>Veya aşağıdaki gibi bir VBA kodu içinde Named Range 
	olarak kullanılabilirler.</p>
	<pre class="brush:vb">Sub tableornek()
  Range("Table1").Select
End Sub</pre>
	<h3>Özet tablo kaynağı</h3>
	<p>Tablolar, Özet tablolara veri kaynağı&nbsp;teşkil edebilirler. Üstelik normal data kaynaklarından farklı olarak tek sefer 
	tanımlanmaları yeterlidir. Böylece veri listesine yeni data eklendikçe Özet 
	tabloyu güncellemek için veri kanağını genişletmenize gerek kalmamaktadır. Bu konuyu
	<a href="InsertMenusu_PivotTable.aspx#datasource">ÖzetTablolar</a> 
	bölümünde gördüğümüz için burada ayrıca detaylara girilmeyecektir.</p>
	<p><img src="/images/inserttable6.jpg"></p>
	<h3>Slicer kullanımı</h3>
	<p>Normal data listelerinde <strong>Slicer</strong> türü filtreler 
		uygulanamazken Table'lara 2013 versiyonundan itibaren Slicer da 
		uygulanabilmekte ve böylece data kümeniz <strong>Dashboard</strong> ekranları hazırlamaya 
		daha uygun hale gelmektedir. Slicerlar hakkındaki detaylı bilgiye
	<a href="DataMenusu_SiralamaveFiltreleme.aspx#Slicer">buradan</a> 
		ulaşabilrisiniz.</p>
	<p><img src="/images/inserttable7.jpg"></p>
	<h3>Formül kolaylığı</h3>
	<p>Table'larda hem Table içi formül yazmak, hem de başka bir yerden formül 
	başvurusunda bulunmak çok kolaydır. </p>
	<p>Önce Table içine formül yazmaya bakalım. Aşağıdaki gibi bir listede G 
	kolonuna sonradan Barem diye bir kolon ekedim. Şimdi buraya bir formül 
	yazarak Tutar'a göre çeşitli baremler(gruplar) oluşturacağız.</p>
	<p><img src="/images/inserttable8.jpg"></p>
	<p>Formül yazmaya başladığımda seçtiğim hücreler D2 gibi görünmek yerine "@" 
	işaretini takiben kolon başlıkları şeklinde görünürler.</p>
	<p><img src="/images/inserttable9.jpg"></p>
	<p>Formülü yazıp Enter'a basılınca formül otomatik aşağı iner.</p>
	<p><img src="/images/inserttable10.jpg"></p>
	<p>Burada önemli bir husus var: Eğer tablomuz bir veritabanından beslenen 
	bir tablo ise sonraki refreshlerde yeni gelen data için formüllerin aşağı 
	inme sorunu yaşanabilmektedir. Böyle olmaması için gereken bazı ayarlamalar 
	var, ona <a href="DataMenusu_BaskaVeriKaynaklariilecalismak.aspx#properties">dış data</a> bölümünde ayrıca değineceğimiz için burada girmiyoruz.</p>
	<p>Bu arada, çok isteyeceğinizi sanmam ama olur da formüllerin otomatik 
	olarak aşağı inmesini istemiyorsanız <strong>Options&gt;Proofing&gt;AutoCorrect 
	Options</strong> altında aşağıdaki işaretli ticki kaldırabilirsiniz.</p>
	<p><img src="/images/inserttable16.jpg"></p>
	<p>Başka bir yerden tablomuza formülle de başvurulduğunda aynı rahatlığı 
	görebiliriz. Aşağıdaki resimde gördüğünüz gibi formülde D:D gibi kolon 
	harfleri yerine kolon başlıkları yer almakta ve bu da sonradan yapılacak 
	formül kontrolü, bakım, düzeltme gibi işlemleri kolaylaştırmaktadır.</p>
	<p><img src="/images/inserttable11.jpg"></p>
	<p>Burda tabi dikkat edilmesi gereken nokta, kolonu seçerken kolon harfine 
	değil, ilgili Table başlığına tıklayarak seçmektir. </p>
	<p>Mesela şu seçim şekli yanlışkken,</p>
	<p><img src="/images/inserttable12.jpg"></p>
	<p>KANAL'ın üzerindeyken çıkan siyah ok varken seçmek doğrudur.</p>
	<p><img src="/images/inserttable13.jpg"></p>
	<p>En temizi ise, ekranı bir iki satır scrolldown yapıp Gri renkli Tablo 
	başlıklarının çıkmasını sağlayıp o şekilde seçmektir.</p>
	<p><img src="/images/inserttable14.jpg"></p>
	<p>Bu arada sadece başlık hücreleri seçildiğinde de aşağıdaki gibi bir ifade 
	yazar. "Yani Table1'in Başlıklarından TUTAR". </p>
	<p><img src="/images/inserttable17.jpg"></p>
	<p>Bunu da özellike başlıklarda çeşitli kategorilerin olduğu durumlarda, 
	OFFSET ve MATCH formülüyle birlikte kullanabilirsiniz. Aşağıdaki örnekte bir 
	kesişim noktası bulunmaktadır. Kesişim formüllerine
	<a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx">Lookup Formülleri</a> 
	bölümünde detaylıca değinilecektir.</p>
	<p><img src="/images/inserttable18.jpg"></p>
	<p>NOT:Table'lara yapılan başvuruların daha etkin kullanımı için MSDN'deki
	<a href="https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e?ui=en-US&amp;rs=en-US&amp;ad=US&amp;fromAR=1">
	bu sayfaya</a> da göz atmak isteyebilirsiniz.</p>
	<h2>Table Menüsü</h2>
	<h3>Table'ı normal alana çevirme</h3>
	<p>Bir nedenle data kümenizi Table olmaktan çıkarıp normal bir alana 
	çevirmek istediğinizde Table menüsünden <strong>Convert to Range</strong> 
	demeniz yeterli. Bunu yaptığınızda listeniz, format olarak hala Table formatını korumuş 
	görünse de, aradan bir satır sildiğinizde otomatik zebra deseninin devam 
	etmediğini görürsünüz.</p>
	<h3>Diğer</h3>
	<p>Tablo oluşturduğunuzda ona Table1 gibi otomatik bir isim verilmektedir. Ancak siz bunu isterseniz aşağıdaki ilk dairedeki gibi değiştirebilirsiniz.</p>
	<p><img src="/images/inserttable15.jpg"></p>
	<p>Tablolara yukardaki gibi Diptoplam satırı da ekleyebilirsiniz(ikinci 
	dairedeki seçim ile), ve bunları 
	sadece Toplam şeklinde değil Min/Max/Ortalama gibi diğer grup fonksiyonları 
	şeklinde de kullanabilirsiniz.</p>
	<p>Table menüsünde yine diğer ilgili menülerden de ulaşabileceğiniz Özet 
	Tablo, Slicer, Remove Duplicates(Mükerrerleri kaldırma) toollarına da erişebilirsiniz.</p>
	<p>&nbsp;</p>


</asp:Content>
