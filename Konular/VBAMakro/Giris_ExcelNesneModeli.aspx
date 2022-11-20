<%@ Page Title='Giris ExcelNesneModeli' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Excel Nesne Modeli</h1>
    <h2 class="baslik">Giriş</h2>
    <div class="konu">
	<p>Visual Basic(VB) gibi nesne yönelimli(İngilizce tabiriyle object oriented(OO)) dillerde <strong>nesneler</strong> ve 
	bu nesnelerin özellikleri(properties), eylemleri(methods) ve olayları(event) bulunur. Gerçi VB tam anlamıyla bir OO dil değildir; bu konsepti destekler ama tam bir OO dil olabilmesi için tüm gereken kriterleri(Encapsulation, Abstraction, Inheritance ve Polymorphism) karşılamaz. Dolayısıyla VBA de tam bir OO dil değildir. Ancak biz burada OO konseptinin detaylarıyla ilgilenmekten ziyade genel anlamda nesne kavramı ve nesnelerin üyeleri üzerinde duracağız.</p>

	<p>Gündelik hayattan örnek verecek olursak pencere bir nesnedir. Pencerenin 
	kulpu bir alt nesnedir, kulpun yapıldığı madde ve rengi ise bir 
	özelliktir, açılmak ise onun bir eylemidir.(İngilizcesi method olduğu ve 
	programcılıkta metod kelimesi daha çok kullanıldığı için bundan sonra eylem yerine 
	metod terimini kullanacağım.)</p>
	<p>Şimdi bu gündelik hayattaki pencere kulpuyla ilgili bir örnek yapalım.</p>
	<p><span>Eğer Pencerenin Kulpu alüminyumsa&nbsp;</span>pencere kapansın, ahşapsa açılsın, başka bir maddeyse ne 
	yapılacağına kullanıcı karar versin.</p>
	<pre class="brush:vb">
Sub Pencere()
If pencere.kulp.malzeme=alüminyum then 
  	Pencere.kapat
ElseIf pencere.kulp.malzeme=ahşap
   	pencere.açıl
Else
  	Msgbox("kararı kullanıcı versin")    	
End If
End Sub     </pre>

	<p>Pencerenin açılma olayına da bir kod ekleyebilirsiniz, şöyle ki:</p>
	<p>Eğer pencere açılırsa radyatörün peteklerini kapat, ki bunun da kodlaması 
	şuna benzer birşey olacaktır.</p>

	<pre class="brush:vb">Sub pencere_afteropen()
  radyator.statu=off
End Sub	</pre>

	<p>Gördüğünüz gibi <strong>nesnelerin üyeleri(özellikleri ve metodları) nokta(.) ile 
	nesnelerinden ayrılmaktadır.</strong></p>
	<p>Bu arada kod bloğumuz içine bir döngü de yerleştirebiliriz, mesela evdeki 
	bütün pencereler için bu yukardaki sorgulamaları yapmak istediğimizi düşünün, 
	her pencere için tek tek kod yazmak çok zahmetli olacaktı, bunun yerine 
	çeşitli döngü yapılarını kullanabiliriz. Burada VBA kodlarından ziyade sadece 
	genel mantığı vereceğim, ilerleyen sayfalarda zaten kodlamasının nasıl 
	yapıldığını göreceksiniz.</p>
		<ul>
			<li>Birinci pencereden başla
</li>
			<li>pencerenin kulpu alüminyumsa açılmasına izin verme,</li>
			<li>Ahşapsa izin ver
		</li>
			<li>Diğer pencereye geç</li>
			<li>Eğer son penceredeysen programdan çık
</li>
			<li>Başa dön </li>
		</ul>
        </div>
        <h2 class="baslik">Detaylar</h2>
    <div class="konu">
	<h4><a name="donendeger"></a>Üyelerin(Property ve Metodlar) dönen değeri</h4>
	<p>Excelde hemen herşey bir nesnedir ve bunlar bir hiyerarşi içinde 
	bulunurlar. Hiyerarşinin en tepesinde <strong>Application</strong> nesnesi vardır, yani 
	Excelin kendisi. Onun altında workbook ve başka nesneler vardır. En sık kullanılacak nesneler 
	<a href="DortTemelNesne_Konular.aspx">bu sayfada</a> detaylıca ele alınacak olup bunların hepsi hiyerarşinin bir seviyesini gösterir.</p>
	<p>Alt nesneden kastımız aslında, bir propertydir, yani terminolojik olarak <strong>nesnenin 
	nesnesi</strong> diye bir kavram yok, ancak nesnenin propertysinin <strong>dönüş değeri</strong> bir 
	nesne tipinde(object type) olduğu için bundan nesnenin alt nesnesi gibi bahsederiz. 
	Ör:Worksheet nesnesinin Range propertysinin dönüş değeri Range nesnesi döndürür ve biz de bunu sanki nesnenin alt nesnesi gibi 
	yorumlarız.</p>
		<h4>Nesne Modeli Grafik Gösterimi</h4>
	<p>Excel Nesne Modelinin grafiksel gösterimi aşağı yukarı şöyledir. 
		Eski Excel versiyonlarında buna program içinden ulaşabiliyorduk ancak 
		şuan yok, ulaşabildiğim bu resmi de 
		<a href="https://msdn.microsoft.com/en-us/library/aa141044.aspx">buradan
		</a>aldım ancak orda bile Page2/3e tıkladığımda birşey göstermiyor.</p>
		<p>
		<img height="60%" src="https://i-msdn.sec.s-msft.com/dynimg/IC62897.gif" width="60%"></p>
	<p>
		Excelin Nesne modeli hakkında daha detaylı bilgi edinmek istiyorsanız
		<a href="https://msdn.microsoft.com/en-us/library/wss56bz7.aspx">MSDN'yi
		</a>ziyaret etmenizi tavsiye ederim.</p>
	<h3>Collection</h3>
	<p>O anda açık olan tüm nesne(obje)&nbsp;grubuna collection denir.&nbsp; Bir 
	nesnenin çoğul hali olarak ifade edilir. Ör:Workbook nesnesi, Workbooks 
	collectionının bir üyesidir. Mesela o anda sadece birinci Workbooku 
	kapayacaksanız Workbooks(1).Close derken, tüm Workbookları kapatmak için 
	Workbooks.Close dersiniz.</p>
	<p>Collection'ları döngüler içinde çok kullanacağız. Mesela aktif dosyanın tüm 
	sayfalarında işlem yapmak için aşağıdaki gibi bir kod yazacağız.</p>
	<pre class="brush:vb">Sub collectionlar()
   'Tanımlamalar
   For each ws in ActiveWorkbook.Sheets
	'kodlar buraya
   Next ws
End Sub</pre>
	<p>Collection konusunu burada bitirelim, döngülerde ve nesnelerde karşımıza 
	tekrar çıkacak, orada detaylarına değineceğiz.</p>
	<p><strong>NOT</strong>: VBA da bize Collection sınıfını sunar, böylece biz 
	de kendi collectionlarımızı yaratabiliyoruz. Bu konuya da yine ilerleyen
	<a href="DizilerveDizimsiYapilar_Collectionlar.aspx">sayfalarda</a> 
	değineceğiz.</p>
	<h3>Class</h3>
	<p>Felsefeyle ilgilendiyseniz platonun idealar dünyasını duymuşsunuzdur. Ona 
	göre dünyada gördüğümüz herşey idealar dünyasındaki bir ideanın dünyada 
	somutlaşmış halidir. Tıpkı bunun&nbsp;gibi VBA'daki her nesne de bir classın 
	somutlaşmış halidir. VBA ile gelen bi dolu class olmakla birlikte ileri 
	seviye bölümünde göreceğiniz üzere kendi class ve dolayısıyla&nbsp;nesnelerinizi 
	de yaratabilirsiniz. Biz mevcut classlar üzerinden bir örnek verip konuyu 
	burada bitirelim, çünkü gerçekten bu kadar detaya boğulmanıza şu aşamada hiç 
	gerek yok, sadece kavramları genel olarak bilin diye bu konuya değiniyorum.</p>
	<p>Mesela Workbooks koleksiyonunun bir üyesi olan Workbook nesnesi aslında 
	bir Workbook classının Workbook tipinde bir nesnesidir. Arkaplanda bu class 
	için tanımlanmış özellik ve metodları vardır. Nasıl idealar dünyasındaki bir 
	atın kulakları, uzun kuyruğu, 4 bacağı gibi özellikleri ve kişnemesi, 
	koşması v.s gibi eylemleri(metodları) varsa workbook classı için tanımlanmış name, path 
	gibi özellikler ve open, close, add gibi metodları vardır ve bunlar 
	bütün workbook nesneleri için geçerlidir.</p>
	<h3>Library(Kütüphane)</h3>
	<p>Bir veya daha çok classtan oluşan kümelere Library denir. (Teknik not:Bunlar aslında 
	bir dll dosyasından başka birşey değildir). Bunların bir kısmı default olarak VBA projelerimize 
	dahildir, bir kısmını ise ihtiyaca göre biz ekleriz, bir kısmını ise hiç 
	kullanmıyor olacağız. Default olarak gelen ve en sık kullanacağımız kütüphaneler 
	şunlardır.</p>
	<ul>
		<li>Excel</li>
		<li>VBA</li>
		<li>Office</li>
	</ul>
	<p>Bunun dışında Access ve Outlookla bilikte çalışmak için bunlara ait 
	kütüphaneleri de VBE içindeki <strong>Tools&gt;References</strong> 
	menüsünden ekleriz. Bir diğer önemli kütüphane de <strong>Scripting.Runtime</strong>'dır. Her ikisini de yeri geldiğinde detaylıca göreceğiz.</p>
	<h3>Object Browser ve nokta notasyonu</h3>
	<p>Tüm Excel Nesne Modeline ve fazlasına ulaşabileceğiniz yer VBE içinden 
	ulaşabileceğiniz <strong>Object Browser</strong>'dır. Burda sol üstte önce bir library seçip 
	sonra bu library içindeki classları ve classların hemen yanında da yani sağ 
	altta da bu classlara ait üyelere(metod, özellik ve olaylara) ulaşabilir, en 
	alt blokta da bunlar hakkında kısa bir bilgi alabilirsiniz.</p>
	<p><img src="/images/vbaobjectexplorer.jpg" width="60%" height="60%" class="zoomla"></p>
	<p>Nesnelerin üyeleri hakkında bilgiye ulaşmanın bir yolu da intellisense 
	teknolojisidir. Object tipli bir dönüş değeri olmayan tüm nesnelerde nesne 
	adını yazıp nokta koyduktan sonra tüm üyelerin gösterilmesine intellisense 
	teknolojisi denir. Bu şekildeki yazım tekniğine de nokta notasyonu denir. 
	Mesela aşağıda bir Range tipli nesnenin intellisense çıktısı görünmektedir.</p>
	<p><img src="/images/vbaobjeintelisense.jpg"></p>

	<p>Tabiki Object Browser toplu bir araştırma ve nesneler yazmak yerine seçme 
	imkanı sunduğu için daha makbuldür, intellisense ise araştırma yapmaktan 
	ziyade daha çok şu işe yaramaktadır. Eğer nesne adından sonra ortaya 
	çıkmıyorsa nesne ismini hatalı yazmışız demektir(Object tipli değil spesifik 
	bir dönüş tipli bir nesne olduğunu varsayıyorum). Ayrıca üye ismini uzun 
	uzun yazmak yerine bir iki harfi yazdıktan sonra Tab tuşuna basarak üye adı 
	otomatik tamamlanmakta ve bu da bize hız kazandırmaktadır.</p>

	<p>Yukarda bahsettiğimiz nesnelerin <strong>dönüş değeri</strong> konusunu <strong>Intellisense </strong>ile 
	bağdaştırmamızda fayda var. Şöyle ki; "Nesne", bildiğiniz gibi Objenin 
	Türkçesidir ama bi ayrım var. Nesne derken nesnenin kendisinden 
	bahsediyoruz, Object derken dönüş tipinden. Bu bağlamda Activesheet de 
	nesnedir ActiveCell de. Ancak ilkinin dönüş tipi Object iken ikincisininki 
	Range'tir. Aktif sayfanın dönüş tipinin WorkSheet olmasını istiyorsak bunu 
	bir değişkene atamalı ve bu değişkeni WorkSheet olarak tanımlamalıyız. Bu 
	arada Activesheet'in dönüş değeri neden Object? diye düşünebilirsiniz.
Bunun sebebi, bu nesnenin birden fazla şekle sahip olabilmesidir: Worksheet, Chart gibi. <strong>İşte Activesheet'te olduğu gibi, birden fazla anlama gelebilecek nesnelerin dönüş değeri hep 
	Object olmaktadır.</strong>
</p>
	<p>
	<img height="147" src="/images/vbaaktifobjedonendeger5.jpg" width="237"></p>
	<h4 id="global">Global sınıfı</h4>
	<p>Object Browserda Classes bölümünde ilk başta duran mavi renkli bir 
	&lt;<span style="color: blue"><strong>globals</strong></span>&gt; classı vardır. Bu class içinde bulunan üyeler global 
	tanımlanmışlardır ve<strong> bağlı oldukları nesnenin kullanımına ihtiyaç duymazlar</strong>. 
	Örneğin, <strong>Math</strong> classında bulunan ve mutlak değer almaya yarayan 
	<strong>ABS</strong> metodunu 
	kullanmak için bir Math nesnesine ihtiyaç duymayız, bu metodu doğrudan 
	kullanırız.</p>
	<p>İşte bu globals içinde bulunan tüm üyeler, farklı farklı classların 
	global üyelerini gösterirler. (Bunlar C# gibi dillerde static tanımlanan 
	üyelere benzemektedirler)</p>
	<p>&nbsp;<img alt="" src="/images/vbaglobalsclass.jpg"></p>


<h3 id="withend">With ... End With yapısı</h3>
<p>Makrolarınızı kaydederken sıklıkla göreceğiniz bir yapı olacak. <span class=" keywordler">With.. End With</span> yapısı. Bir nesnenin üyelerine arka arkaya sıklıkla başvurmanız gerektiği durumlarda bu yapıyı kullanırız. Zorunlu değil tabiki ancak, hem daha az kod yazmamızı sağlar hem de okunurluğu iyileştirir.</p>

<p>Bu yapıda, <strong>With... End With</strong> arasında bulunan ilgili nesnenin üyeleri, önlerinde nesnenin adı yazılmadan sadece . işaretini takip edecek şekilde yazılırlar.</p>

<p>Şimdi yine yukarıdaki pencere örneği üzerinden giderek bir örnek yapalım.</p>

<pre class="brush:vb">
With Pencere
   .kulp.malzeme="alüminyum"
   .kulp.kalınlık=10
   .kulp.çevir
   .aç
   evi_havalandır 'burada başka bir fonksiyon çağırıyoruz, o yüzden başında . yok
   If ev.hava="iyi" then
       .kapat
   End If
End With
</pre>
</div>
</asp:Content>
