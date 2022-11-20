<%@ Page Title='Fonksiyonlar UDFKullaniciTanimliFonksiyonlar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='5'></asp:Label></td></tr></table></div>

<h1>Kullanıcı Tanımlı Fonksiyonlar(UDF)</h1>

	<h2 class="baslik">Neden?</h2>
	<div class="konu">
	<p>Bazı anlar olur ki "Ya bunun formülü nasıl yazılıyordu" dersiniz, o formülü daha önce yazmışsınızdır ama o kadar uzun bir formüldür ki tekrar hatırlamak biraz zamana mal olur, hele o parantezler yok mu, "ya bu parantezlerden hangisi fazla" diye düşünür durursunuz. Burada bir alternatif, daha önceden yaptığınız ve bir yerlere(!) kaydettiğiniz çalışmayı bulmak ve formülü kopyalamak, bir diğeri de kendi fonksiyonunuzu 
	yani UDF(User Defined Functions) tanımlamak ve bunu her ihtiyacınız olduğunda çok kolay şekilde kullanmaktır.</p>

<p>UDF'in tek kullanım amacı bu değildir tabiki, bir diğer kulanım amacı ise, Excel'in mevcut 
fonksiyonlarıyla yapmanın zor olduğu hatta imkansız olduğu şeyleri bunlarla yapabilmektir. 
Evet, her ne kadar Excel'in çok geniş bir fonksiyon kütüphanesi olsa bile bunlar bazen 
yetersiz kalabilmektedir. Gerçi her yeni versiyonda bazı eksiklikler 
tamamlanmaktadır. Şahsen benim yazdığım en az bi düzine fonksiyon artık son 
durumda(bu sayfayı yazarken son versiyon 2016 idi) boşa çıkmıştır. Bunlara da yeri 
gedikçe bakacağız.</p>
</div>
	<h2 class="baslik">Giriş</h2>
		<div class="konu">
	<p>Şimdi UDF konusuna girerken olaylara bakışmızı biraz değiştirmekte, 
	kendimizi Micorosft çalışanıymış gibi düşünmekte fayda var. Kendimize şunu 
	soralım: Yerel fonksiyonları Excel 
	mühendisleri nasıl hazırlamış olabilir?</p>
	<p>Mesela <strong>SUM</strong> fonksiyonuna bakalım:</p>
	<p>
	<img src="/images/vbaudf1.jpg"></p>
	<p>Bu yukarıdaki gibi bir fonksiyonu biz nasıl hazırlardık? Bunu önce sözlü dile 
	getirelim. <strong>Seçili alandaki tüm hücrelerin değerini tek tek topla</strong>. 
	Algoritmik hali de şöyledir.</p>
	<ul>
		<li>İlk hücreden toplamaya başla</li>
		<li>Sonraki hücreye geç, bir önceki değerle topla</li>
		<li>Bunu son hücreye kadar devam ettir</li>
	</ul>
	<p>O zaman son minvalde kodumuz şöyle olacaktır:</p>
	<pre class="brush:vb">
Function Topla(alan As Range) As Double
Dim a As Range
Dim gecici As Double

For Each a In alan
    gecici = gecici + a.Value
Next a

Topla = gecici
End Function
</pre>
	<p>
	Sonuca bakalım.</p>
	<p>
	<img src="/images/vbaudf2.jpg"></p>
	<p>
	Ne yaptığmıza bir bakalım</p>
	<ul>
		<li>Fonksiyon ismini <span class="keywordler">Function</span> kelimesi ile belirttim, tipini de 
		Double.(Tip belirtmezsem Variant olur)</li>
		<li>Fonksiyonuma parametre olacak olan "alan"ı veri tipi Range olacak 
		şekilde belirttim.(Normalde yerel SUM fonksiyonu kendisine bir alan 
		değil de adedi belirsiz olan sayısal değerleri kabul eder, biz basit olsun diye 
		alan belirttik. Yerel SUM'ın yaptığı gibi de yapabilirdik, bunu 
		<a href="#param">ParamArray</a> 
		kısmında göreceğiz)</li>
		<li>Sonra geçici bir değişken tanımladım, hücreler içinde dolaşırken ara 
		toplamı hep bu geçici değişkende tuttum.(Geçici değişkenler Function 
		tanımlamalarında sıklıkla kullanılırlar)</li>
		<li>Sonra bir For Each döngüsü ile tüm hücrelerin arasında dolaşıp 
		toplamı hesapladım</li>
		<li>En son da fonksiyon ismi olarak belirttiğim ifadeye(Topla) geçicinin 
		değerini atadım</li>
	</ul>
	<p>Fonksiyon tanımlamanın jenerik yapısı aşağıdaki gibidir:</p>
	<pre class="brush:vb">
Function fonksiyonadı(parametre1 As veritipi,parametre2 As veritipi,...) As DönüşTipi
'Gerekliyse değişken tanımlamaları

'Kod bloğu
'varsa geçicideğer

fonksiyonadı=geçicideğer
'geçici değer yoksa hesabı direk fonksiyon adı üzerinde yaparız
'fonksiyonadı=hesaplama kodları
End Function</pre>
	<p>Mesela aynı mantıkla <strong>COUNT</strong> fonksiyonunu düşünün. Bu fonksiyon bildiğiniz 
	gibi sayı içeren hücreleri sayar. Hadi biz de aynı görevi gören bi UDF hazırlayalım. 
	VBA, boş hücreleri de 0 gibi düşündüğü için bunları numerik sayar, o yüzden 
	dolu hücrelere sayı içeriyor mu diye bakacağız.</p>
	<pre class="brush:vb">
Function NumerikSay(alan As Range)
Dim a As Range
Dim gecici As Double

For Each a In alan
   If IsNumeric(a) And Not IsEmpty(a) Then gecici = gecici + 1
Next a

NumerikSay = gecici
End Function</pre>
	<p>Aynı mantıkla COUNTA fonksiyonunu düşünün. Bu fonksiyon bildiğiniz gibi 
	içi dolu hücreleri sayar. Bunun da UDF versiyonunu hazırlayalım.</p>
	<pre class="brush:vb">
Function DolularıSay(alan As Range)
Dim a As Range
Dim gecici As Double

For Each a In alan
   If Not IsEmpty(a) Then gecici = gecici + 1
Next a

DolularıSay= gecici
End Function</pre>
	<h3>
	Performans</h3>
	<p>
	Hazırladığınız fonksiyonu binlerce satırdan oluşan listelerde 
	kullanacaksanız mutlaka hem fonskiyonun parametrelerini hem de kodda 
	kullanılacak tüm değişkenleri uygun veri tipinde tanımlayın. Buna 
	rağmen fonksiyonunuz çok komplike ise yeterince hızlı çalışmayabilir. Yerel 
	fonksiyonlarla yapmak çok zor değilse yerel fonksiyonları kullanmanız 
	gerekebilir. Yerel fonksiyonlar sonuçta en temel seviyde çalışırlar. Kendi 
	UDF'lerimiz ise bir seviye üstte çalşır, o yüzden temel seviyeye 
	çevrilmeleri gerekir, en son da makine diline çevrilirler. Bu performans 
	farkını test ettikten sonra görebilir, duruma göre karar verirsiniz.</p>
            <p>
	            &quot;Fonksiyonu UDF olarak hazırlamak istiyorum ama performansı da hızlı olsun&quot; diyorsanız XLL ile tanışma vaktiniz gelmiş demektir. Bunun için sizi <a href="../VSTO/ThirdPartyKutuphanler_ExcelDNA.aspx">şöyle</a> alayım.</p>
	</div>
	<h2 class="baslik">Yerleşim ve Erişim</h2>
		<div class="konu">
	<h3>Yerleşim</h3>
	<h4>Tekil kullanımlık UDF'ler</h4>
	<p>Yazdığımız fonksiyonları bazı durumlarda sadece ilgili dosya içinde 
	çalışmasını isteriz, çünkü o dosyaya özgü çözüm sunarlar. Bunları başka bir yerde 
	kullanmayacağımız için genel UDF'leri tutacağımız yerde bulundurmamıza gerek 
	yoktur.</p>
	<p>İlgili dosyada UDF yazmak için o dosyaya Modül eklemeli, modül sayfasına 
	yazmalıyız. ThisWorkbook ve Sheet modülleri UDF'ler için uygun değildir.</p>
	<p><img src="/images/vbaudf5.jpg"></p>
	<p><span class="dikkat">Dİkkat</span>:Tekil kullanımlık UDF'in kaydolduğu 
	dosyayı <strong>xlsm</strong> veya <strong>xlsb</strong> uzantısıyla kaydetmek gerekmektedir.</p>
	<h4>Genel kullanım UDF'lerinin yerleşimi</h4>
	<p>Eğer kodlarınızı genele yaygın bir şekilde kullanmak istiyorsanız bunun 
	için bir alternatif Personal.xlsb dosyasıdır ama bu genelde kötü bir alternatiftir. 
	Normal Sub prosedürler için Personal.xlsb güzel bir adrestir ancak 
	Function'lar için aynısını söyleyemeyiz. Zira bunlara ne <strong>Insert Function
	</strong>menüsündeki <strong>User Defined </strong>kategorisinden ulaşmak kolaydır, ne 
	de bir hücreye doğrudan ismini yazarak 
	ulaşabiliriz. </p>
	<p>User Defined kategorisinden 
	ulaşmak biraz karışıktır. Diyelimki fonksiyon adı DolularıSay olsun. Bu 
	fonksiyon listede <strong>Personal.xlsb!DolularıSay </strong>&nbsp;şeklinde 
	görünür ancak bunun için P harfine gidip önce 
	Personla.xlsb'yi bulmanız gerekmez. 
	UDF'ler, içinde bulundukları dosyadan bağımsız olarak kendi adlarına göre 
	alfabetik sıralıdır. Dolularısay için de D harfiyle başlayan fonksiyonlara 
	gitmeniz gerekir. Aşağıda görüldüğü üzere N harfiyle başlayan NumerikSay 
	fonksiyonu P harfiyle başlayan Personal.xlsb!DolularıSay 'dan sonra geliyor, 
	önce değil.</p>
	<p><img src="/images/vbaudf7.jpg"></p>
	<p>Bunu seçtiğimizde yukardaki ilk örnekte gördüğümüz gibi Excel içindeki 
	görünümü <strong>=Personal.xlsb!DolularıSay(B2:D7)</strong> şeklinde olup 
	biraz garip bir görünüme sahiptir. Normal bir dosya içindeki tekil kullanımlık UDF'lerde 
	ise bu sorun 
	yoktur.</p>
	<p>Peki soru şu:UDF'imizi hem tekil kullanımlık UDF'teki gibi önünde dosya ismi 
	olmadan kullanmak hem de genele yaygın kullanmak istiyorsak ne yapmalıyız? 
	<strong>Cevap</strong>:UDF'lerimizi <strong>Add-in</strong> içinde kaydetmek. </p>
	<h4>Add-in içinde yerleşim</h4>
	<p>
	Boş bir dosya açın, kaydet düğmesine basın, dosya tipini Excel Add-in olarak 
	değiştirin. </p>
	<p>
	<img src="/images/vbaudf6.jpg"></p>
	<p>
	Otomatikman adresin yukardaki gibi değişmesi lazım. Sizdeki 
	adres Excel versiyonunuza 
	göre değişebilir. Bende şöyle: <span>
	<strong>C:\Users\Volkan\AppData\Roaming\Microsoft\AddIns\</strong></span></p>
	<p>
	Bu add-in'i <strong>Developer menüsü&gt;Excel Add-in</strong> menüsünden aşağıdaki 
	gibi aktive edebilirsiniz. Bir kere aktive olduktan sonra Excel her 
	açıldığında bu dosya da açılacaktır. Dosyanın kendisi görüntülenemez. Sadece 
	VBE ortamında görünür. Yanlız bunlar 
	Personal.xlsb'den biraz farklıdır. Personal.xlsb'yi istersek unhide yaparak 
	görebiliriz ancak 
	add-in dosyaları asla normal worbooklar gibi görüntülenemezler. Aşağıdaki 
	görselde farkettiyseniz Excelle hazır gelen Anlaysis Toolpak ve Solver gibi 
	add-inler de var. Bunların kullanımıyla ilgili detaylara yandaki anamenüde 
	Excel altındaki Data menüsünden ulaşabilirsiniz.(NOT:Developer menüsündeki COM 
	Add-in butonu bu sitenin <a href="../VSTO/Giris_Konular.aspx">VSTO</a> 
	konusuyla ilgili olduğu için ona burada değinmeyeceğiz.)</p>
	<p>
	<img src="/images/vbaudf8.jpg"></p>
	<p>
	Bundan sonrasında yapılması Gereken bu dosyaya VBE'de modül ekleyip UDF kodlarımızı 
	oraya yazmaktır.</p>
	<p>
	<img src="/images/vbaudf9.jpg"></p>
	<p>
	Gördüğünüz gibi artık UDF'in önünde dosya adı yok ve de herhangi bir dosyada 
	kullanabiliyoruz.</p>
	<p>
	&nbsp;</p>
	<p>
	<img src="/images/vbaudf10.jpg"></p>
	<h3>
	Erişim</h3>
	<p>
	Add-in'lerinizi başka kişilerle de paylaşmaya karar verdiyseniz bunların 
	nerede olduğuna ulaşmak için VBE'de Immediate Window'a şunu yazın: <strong>
	?Application.UserLibraryPath. </strong>Add-inlerinizin aslında istediğiniz yere kaydedebilirsiniz, ancak daha sonra 
	aktive(veya pasifize) ederken&nbsp; kolaylık olması adına bu adreste olması 
	daha iyidir.</p>
	<p>
	Aktivasyon işlemini yukarda göstermiştik. Bir diğer alternatif de, 
	<strong>File&gt; 
	Options&gt;Add-ins</strong> menüsündendir.</p>
	<p>
	<img src="/images/vbaudf11.jpg"></p>
	<p>
	Son olarak,
	Add-inler sadece UDF depolamak için kullanılan araçlar değillerdir. Ayrıca menü yaratıp, bu menüye 
	çeşitli düğmeler ve altmenülere ekleyerek Sub prosedürlerinizi yani 
	makrolarınızı çalıştırmak için de bir arayüz sağlarlar. Bu konuya ayrı 
	<a href="Ileriseviyekonular_Add-InlerveCustomMenuler.aspx">sayfada</a> değineceğiz.</p>
	</div>
	<h2 class="baslik">UDF Açıklamaları ve Fonksiyon Kategorisi</h2>
		<div class="konu">
	<p>Fonksiyonlara açıklama eklemek özellikle başkalarının kullanımı için 
	faydalı olabilir. Bunun için aşağıdaki gibi bir ayarlama yapılabildiği gibi 
	toplu tanımlama yapma imkanı da varır, ki bunun için tabiki yine VBA kodu 
	kullanırız.</p>
	<p><img src="/images/vbaudf12.jpg"></p>
	<p>Çıkan kutuya direkt açıklama yazılır.</p>
	<p><img src="/images/vbaudf13.jpg"></p>
	<p>Toplu işlem için kullancağımız kodda ise fonksiyonun ait olacağı 
	kategoryi de belirtebiliriz. Mesela NumerikSay için İstatistiki kategorisine 
	koyabiliriz, ki yukardaki yöntemde bu yapılamamaktadır.</p>
	<p>Mevcut kategoriler aşağıdakiler olup bunların dışında yeni kategoriler 
	de yaratabiliriz.</p>
	<p>'0 Kategori yok, All içinde görünür<br>'1 Financial<br>'2 Date &amp; Time<br>'3 Math &amp; Trig<br>'4 Statistical<br>'5 Lookup 
	&amp; Reference<br>'6 Database<br>'7 Text<br>'8 Logical<br>'9 Information<br>
	....<br>'14 User Defined default<br>'15 Engineering (Analysis Toolpak 
	add-in'i kuruluysa kullanılabilir)</p>
	<p>
	Aşağıdaki kod, NumerikSay fonksiyonunu Sayısal Fonksiyonlar isimli kategori 
	içine atar, bu kategori mevut değilse bunu yaratır.</p>
	<pre class="brush:vb">
Sub UDFaçıklaması()
  Application.MacroOptions _
  Macro:="NumerikSay", _
  Description:="Belli bir alandaki sayısal değerlerin toplam adedini verir", _
  Category:="Sayısal Fonksiyonlar"
End Sub</pre>
	<p>Excel 2010 ile birlikte fonksiyonların argümanlarına da açıklama girme 
	imkanı gelmiştir. MacroOptions metoduna <strong>ArgumentDescriptions</strong> parametresini 
	ekleyerek bu işi hallederiz.(Normalde VBA ölü bir dildir, yani yeni 
	güncellemeler, özellikler eklenmez ancak ender de olsa böyle küçük 
	iyileştirmeler yapılmakta)</p>
	<p>Önemli bir nokta var ki, o da bu girdiğiniz açıklamalar sadece sizde geçerli 
	olur. UDF'erinizi bir add-in haline getirip gönderirseniz ve onlarda da 
	görünmesini isterseniz bu tanım ekleme kodlarını Workbook_Open içine yazmakta fayda var, 
	ki kişinin add-ini yüklenir yüklenmez açıklama kodu da çalışsın. UDF'lerin 
	açıklamaları ve parametrelerin açıklamalarıyla birlikte kodunuz çok uzayacak 
	gibi olursa bunları bir text dosyasından veya 
	Excel dosyasından okutabilirsiniz.(Biraz aşağıda örneği var)</p>
	<p>Şimda tek bir fonksiyon için nasıl yapıyoruz ona bakalım.</p>
	<pre class="brush:vb">Sub UDFaçıklaması()

Dim fonkAd As String
Dim fonkTanım As String
Dim fonkKategori As String
Dim argumanAdedi As Integer
Dim argumanlar() As String

argumanAdedi = 1
ReDim argumanlar(1 To argumanAdedi)

fonkAd = "NumerikSay" 
fonkTanım = "Belli bir alandaki sayısal değerlerin toplam adedini verir"
fonkKategori = 4

argumanlar(1) = "Toplanacak sayıların olduğu alanı seçin."

Application.MacroOptions _
Macro:=fonkAd, _
Description:=fonkTanım, _
Category:=fonkKategori , _
ArgumentDescriptions:=argumanlar

End Sub</pre>
			<p>Herhangi bir hücreye fonksiyonumuzu yazdığımızda aşağıdaki gibi 
			görünür:</p>
			<img src="/images/vbaudf14.jpg">
			<p>Aşağıda da Workbook_Open makrosuna yazılı olan kod var. Bu kod 
			ile, fonksiyon açıklamalarını bir excel dosyasından okuyoruz.</p>
			<pre class="brush:vb">
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    Workbooks.Open "C:\makrolar\kılavuz.xlsx", UpdateLinks:=0 'herkes dosyayı buraya koymalı
    Workbooks("kılavuz.xlsx").Sheets("UDFDesc").Select
    Call macrodesc
    Workbooks("kılavuz.xlsx").Close savechanges:=False
    Application.ScreenUpdating = True
End Sub

'bu da çaprılan macrodesc prosedürü
Sub macrodesc()

Dim fonkAd As String
Dim fonkTanım As String
Dim fonkKategori As String
Dim argumanAdedi As Integer
Dim argumanlar() As String
Dim i As Integer

Range("A2").Select
Do
    fonkAd = ActiveCell.Value2
    fonkTanım = ActiveCell.Offset(0, 1).Value2
    fonkKategori = ActiveCell.Offset(0, 2).Value2
    argumanAdedi = Range(ActiveCell, ActiveCell.End(xlToRight)).Cells.Count - 3
    
    ReDim argumanlar(1 To argumanAdedi)
    
    For i = 1 To argumanAdedi
        argumanlar(i) = ActiveCell.Offset(0, i + 2).Value2
    Next i
        
    Application.MacroOptions _
        Macro:=fonkAd, _
        Description:=fonkTanım, _
        Category:=fonkKategori, _
        ArgumentDescriptions:=argumanlar
Loop Until ActiveCell.Value = ""

End Sub
</pre>
</div>
	<h2 class="baslik">Diğer Hususlar</h2>
		<div class="konu">
	<h3><a name="param"></a>ParamArray ile belirsiz sayıda parametre temini</h3>
	<p>Yukarıda Excelin yerel SUM fonksiyonunu, kaynağı bir alan olacak 
	şekilde tasarlamış ve demiştik ki, aynı yerel SUM gibi sayısı belirsiz olan 
	elemanları içerecek şekilde de yapabiliriz. İşte bunun yolu <strong>ParamArray</strong> 
	ifadesini kullanmaktır. Önce örneği yapalım sonra açıklayalım.</p>
	<pre class="brush:vb" style="margin-top: 19px">Function ToplaParamarray(ParamArray sayılar()) As Double
Dim a As Variant
Dim gecici As Double

For Each a In sayılar
gecici = gecici + a
Next a

ToplaParamarray = gecici

End Function</pre>
	<p>ParamArray ifadesinden sonra bu sayısı beliri olmayan elemanları içerecek 
	bir dizi adını gireriz. Bu Variant tipli bir değişken olmak zorundadır, yani 
	başka bir veritipiyle tanımlayamayız. Zaten mantıklısı da budur. Mesela 
	Toplam örneğinde hem 0,25 gibi double tipindeki sayıları hem de 5 gibi 
	integerları input olarak alabiliriz, böyle bi durumda Varianttan başka çare 
	yoktur. </p>
	<p>Diğer kıstlamalar şöyle;</p>
	<ul>
		<li>Fonksiyonun parametre listesinde başka parametreler de varsa 
		Paramarray sonucu parametre olmalıdır. Ör:ToplamınXinciÜssü şeklindeki 
		fonksiyon şöyle tanımlanırdı. <strong>Function ToplamınXinciÜssü(üs As 
		Integer,ParamArray sayılar())</strong></li>
		<li>Bir fonksiyonda sadce 1 tane paramarray tanımlanabilir</li>
		<li>Bir fonksiyonda Optional ve Paramarray parametrelerden yanlız biri 
		kullanılabilir.</li>
		<li>İlgili dizinin tabanı 0'dır. İsterseniz Option Base ile genel dizi 
		tabanını 1 yapmış olun farketmez. ForEach yerine klasik for kulanılacaksa Lbound(dizi) to Ubound(dizi) 
		şeklinde kullanılabilir.</li>
	</ul>
			<p>NOT:Bu yukardaki örnekte bir alan seçimi yapılamayacağına dikkat 
			edin; tek tek sayı temin etmek zorundasınız. Excel'in SUM fonksiyonu 
			ise hem alan kabul edebiliyor hem de tek tek sayılar kabul 
			edebiliyor. SUM'ın bunu nası yaptığını düşünün ve her iki versiyonu 
			da kapsayacak bir fonksiyonu yazmayı deneyin.</p>
	<h3>Optional ile opsiyonel parametre temini</h3>
	<p>Bazen girdiğimiz parametrelerin sık kullanılan değerini biz baştan 
	gireriz ve son kullanıcıya bunu değiştirme imkanı veririz. Bu, yerel Excel 
	fonksiyonlarında da var olan default(varsayılan) değerlerle aynı şeydir. Mesela VLOOKUP 
	fonksiyonunun son parametresi opsiyonel olup default değeri TRUE(1)'dur, yani 
	girilmezse TRUE(1) algılanır.</p>
	<p>İşte biz de bu opsiyonel parametreyi Optional ifadesi ile sağlarız, istersek varsayılan değer de girebiliriz, girmezsek de bunun girilip 
	girilmediğini IsMissing fonksiyonu ile test ederiz. Ancak IsMissing 
	sorgulaması yapabilmek için ilgili parametrenin datatipi Variant olarak girilmelidir. (ParamArrayda datatipi Variant 
	olmak zorundayken Opsiyonellerde ise tavsiyedir, olur da IsMissing ile sorulgarız diye, 
	ama zorunlu değildir). Bu değer Variant değilse ve kullanıcı tarafından 
	girilmezse bunlara default değerleri atanır; String için "", sayısal tipler 
	için 0, boolean için false.</p>
	<p>Dikkat edilmesi gereken diğer hususlar şöyledir:</p>
	<ul>
		<li>Bunlar da ParamArray gibi son parametre olarak girilmelidir</li>
		<li>ParamArray'de belirttiğimiz gibi aynı anda hem ParamArray hem 
		Optional ifadesi kullnılamaz.</li>
		<li>Birden fazla Optional ifadesi kullanılabilir </li>
	</ul>
	<h4>Opsiyonel Örnek 1</h4>
	<p>Aşağıda varsayılan değerin girilmiş olduğu bir örnek bulunuyor. Bu 
	örnekte bir hücredeki metni kelimelere ayırıyoruz. Normal bir cümlede ayraç boşluk olacağı için ayracı girmeye gerek yok, 
	o yüzden default olarak " " atadım. Ama kullanıcı isterse kelimeleri farklı 
	bir ayraçla da ayırabilir. Mesela - ile ayrılmış kelimeler varsa ayraç 
	olarak - kullanılabilir. Fonksiyonda Split fonksiyonu kullanarak kelimeleri 
	parçalıyor ve bir diziye atıyorum. Kaçıncı parametresi ile de istediğim 
	kelimeyi elde ediyorum.</p>
	<pre class="brush:vb">
Function kelimesec(hucre As Range, kaçıncı As Byte, Optional ayrac As String = " ")
    Dim kelimeler As Variant
    kelimeler = Split(hucre.Value2, ayrac)
    kelimesec = kelimeler(kaçıncı - 1)
End Function</pre>
	<p>
	Örnek kullanım&nbsp; aşağıdaki gibidir. B7'deki formül şöyle:</p>
	<pre class="formul">=kelimesec(A7,2) //son parametreyi girmedim</pre>
	<p>
	<img src="/images/vbaudf15.jpg"></p>
	<h4>Opsiyonel Örnek 2&nbsp;</h4>
	<p>Şimdi üç tane opsiyonel değişkeni olan bir fonksiyonumuz var. Bunlardan birinin varsayılan 
	değeri girilmiş, 
	birinin girilmemiş, biri de Variant olarak tanımlanmış. 
	</p>
	<p>Bu örnek 
	bankacılık dünyasından bir örnek olacak. Mevduat/Kredi gibi hacimsel bir 
	büyüklük ile bu hacimden ne kadar kar elde ettiğimizi gösteren spread 
	bilgisini zorunlu olarak giriyoruz. Bu fonksiyonu iki şekilde kullanabiliyoruz. Eğer belirli bir alan seçilmezse 
	<span style="text-decoration: underline">şube bazında</span> 
NFG(Net Faiz Geliri) hesplayacağız, ancak ilgili alanı seçersek o alandaki 
	<span style="text-decoration: underline">MT(Müşteri Temsilcisi) adedi başına 
	</span>NFG hesaplanmış olacak. İşte bu alan bilgisini Variant olarak girdim, ki IsMissing ile alanın girilip girilmediğini sorgulayabileyim. 
	İkinci opsiyonel seçenek Kur bilgisi olup, ilgili hacim türünün TL mi yoksa döviz mi olduğuna göre değer alacak. Varsayılan 
	olarak döviz tipinin TL
	olduğunu düşünerek değeri 1 girdim, ancak kullanıcı isterse farklı bir kur değeri girebilir. Son değişken ise Para birimi olup, 
	varsayılan değer girilmemiştir. Kullanıcı isterse TL, USD gibi değerler girebilir, eğer girmezse 
	Stringlerin varsayılan değeri olan "" atanacak olup
	herhangi bir para birimi yazmayacaktır. 	</p>
	<p>Son olarak, fonksiyon için dönüş tipi belirtmedim, yani Variant olacak. 
	Zira Birim parametresi girilirse sonuç String, girilmezse double olacak,&nbsp; 
	yani ikisini de kapsayan&nbsp; bir tip olmalı, ki bu da Variant oluyor.</p>
	<pre class="brush:vb">
Function NFGHesapla(hacim As Double, spread As Double, Optional alan As Variant, Optional Kur As Double = 1, Optional Birim As String)
    Dim adet As Integer
    
    If IsMissing(alan) = True Then
        NFGHesapla = (hacim * spread * Kur / 1200) & Birim
    Else
        NFGHesapla = (hacim * spread * Kur / 1200) / WorksheetFunction.CountA(alan) & Birim
    End If
End Function	</pre>
	<p>
	Kullanım şekli aşağıdaki gibidir. G kolonunda şube başına ve para birimsiz 
	versiyonunu görürken, H kolonunda MT başına ve para birimli versiyonunu 
	görüyoruz.</p>
	<p>
		<img src="/images/vbaudf16.jpg"></p>
	<p>
		<strong>NOT</strong>: Bu örnekle, aslında modern programlama dillerinde 
		olan ancak VBA'de olmayan "method overloading" kavramını (tam olarak 
		olmasa da) bir nevi taklit etmiş olduk. Yani bir metodun ismi aynı olup 
		farklı parametre veya dönüş tipi alıyorsa bu işleme method overloading 
		denir. Biz de buna benzer birşey yapmış olduk. Hem alan tipini seçip 
		seçmemeye göre hem de sonuna parabirimi koyup koymamaya göre 4 farklı 
		kullanım şekli sunduk.</p>
	<h3 id="volatile">Volatile</h3>
	<p>
	Yazdığmız fonksiyon, eğer sayfada bir güncelleme olduğunda bundan 
	etkileniyorsa anında güncellenmez. Volatile ile bu güncellemeyi anlık 
	olarak yapmış oluruz. Ancak üstadlar der ki, UDF'inizi öyle bi hazırlayın 
	ki bu fonksiyonu kullanma ihtiyacınız hiç olmasın.</p>
	<p>
	<a href="https://stackoverflow.com/questions/24353506/non-volatile-udf-always-recalculating">Bu</a> 
	ve
	<a href="http://dailydoseofexcel.com/archives/2004/06/22/volatile-functions/">
	şu</a> sayfalarda bu konuyla ilgili çeştli tartışmalar da yapılmış, İngilizceniz varsa bakabilirsiniz.</p>
	<p>
	Özet&nbsp; tavsiyem: Araştırmalarınızda bu 
	ifadeyi görüseniz ne olduğunu bilin ama 
	bunu kullanmayın. UDF'inizde bunun kulanımına gerek olmayacak şekilde tüm 
	parametreleri dahil edin.(Örnek vererek kafanızı da karıştırmak 
	istemiyorum)</p>
			<h3>
			Default Değerler</h3>
			<p>
			Fonksiyonların da tıpkı değişkenler gibi default değerleri bulunur. 
			Örneğin, Boolean tipli bir fonksiyon için bir koşul bloğu içinde 
			sadece True ataması yapıyorsanız ve bu koşul sağlanmıyorsa, siz 
			açıkaça Else bloğu içinde False ataması yapmasanız bile fonksiyon 
			False döndürür. (Bu açıklama hem Excel UDF'leri hem VBA UDF'leri 
			için geçerlidir)</p>
			<p>
			Aşağıdaki örneği inceleyelim. </p>
			<pre class="brush:vb"> Function aysonumu(tarih As Date) As Boolean
    If Month(tarih) <> Month(tarih + 1) Then 
	aysonumu = True
    Else
	aysonumu=False
    End If
End Function</pre>
			<p> 
			Bu fonksiyonu daha ksıa bir şekilde aşağıdaki gibi yazabiliriz. Zira 
			fonksiyona ilk girildiği anda, fonksiyon booelan tipli bir fonksiyon 
			olduğu için giriş annda default değeriyle yani False olarak hayatına 
			başlar. Sonrasında kendisine bir atama olmazsa da bu değerini korur, 
			yani False döndürür.</p>
			<pre class="brush:vb"> Function aysonumu(tarih As Date) As Boolean
    If Month(tarih) <> Month(tarih + 1) Then aysonumu = True
End Function</pre>
			<p> 
			NOT:Bu fonksiyon hem VBA hem Excel UDF'i olarak kullanılabilen bir 
			fonksiyondur.</p>
	</div>

<h2 class="baslik">Çeşitli örnekler</h2>
	<div class="konu">
	<h4 class="baslik">Süpercombine ile hücreleri birleştirin</h4>
	<div class="konu">
	<p>Bu fonksiyon, ardışık bir hücre grubunu belirli bir işaret/ayraç ile 
	birleştirir. 2016 ile gelen TEXTJOIN ve CONCAT fonksiyonlarıyla gereksiz 
	hale gelmiştir ancak eski Excel versiyonlarını kullananlar için hala 
	geçerlidir.</p>
	<pre class="brush:vb">
Function Süpercombine(Hucre_grubu As Range, isaret As String)
x = ""
For Each k In Hucre_grubu
    If Not IsEmpty(k) Then
    x = x & k.Value & isaret
    End If
Next k

Süpercombine= Mid(x, 1, Len(x) - 1)
End Function</pre>
	<p>
	Örnek kullanım aşağıdaki gibidir. Müşteri numaraları ";" işareti ile 
	birleştirilmiştir.</p>
	<p>
	<img src="/images/vbaudfsupercombine.jpg"><br>
</p>
</div>

	<h4 class="baslik" id="superlookup">Süperlookup ile kayan lookup işlemi yapın</h4>
	<div class="konu">
	<p>Sheet2'de şöyle bir listemiz var,</p>
	<p>
	<img src="/images/vbaudf17.jpg"></p>
	<p>
	Sheet1'de bulunan aşağıdaki listeye(sadece A kolonu olduğunu düşünün) lookup çekmek 
	isteseydik şu formülü yazardık.(veya yardımcı bir index satırı kullanırdık, 
	ama şık olanı bu aşağıdakidir)</p>
	<pre class="formul">=VLOOKUP($A2;Sheet2!$A:$F;MATCH(B$1;Sheet2!$A$1:$F$1;0);0)</pre>
	<p>
	İşte bizim süperlookup fonksiyonumuz bizi bu uzun formülden kurtarmış 
	oluyor.</p>
	<pre class="brush:vb">
Function süperlookup(alan As Range, sütun As Range, aranan As Range)
On Error GoTo hata
    süperlookup = alan.Columns(1).Find(aranan, lookat:=xlWhole).Offset(0, sütun.Column - alan.Columns(1).Column).Value
    Exit Function
hata:
    süperlookup = "Bulunamadı"
End Function</pre>
	
	<p>
	<img src="/images/vbaudf18.jpg">	
	</p>
</div>

<h4 class="baslik">Çok eşleşmeli lookup</h4>
	<div class="konu">
	<p>Bu örneğin kodu aşağıda olup açıklamasını
	<a href="../Excel/FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx#cokeslesmelilookup">
	şu sayfada</a> bulabilirsiniz.</p>
    <pre class="brush:vb">
Function çok_eşleşmeli_vlookup(aranan As Range, alan As Range, kaçıncıkolon As Integer, Optional distinctmi As Boolean = True)
Dim a As Range
Dim dict As New Scripting.Dictionary 'Reference olarak eklenmiş olmalı, ekli değilse Late binding olarak yaratılabilir
  
If alan.Columns(1).Rows.Count = 1048576 Then
    Set alan = Range(alan(1, 1), alan(1, 1).End(xlDown))
End If
  
    For Each a In alan.Resize(, 1)
        If Not dict.Exists(a.Value) Then
            dict.Add a.Value, a.Offset(0, kaçıncıkolon - 1).Value
        Else
            geçici = dict(a.Value)
            dict.Remove (a.Value)
            x = a.Offset(0, kaçıncıkolon - 1).Value
            If distinctmi = True Then
                dict.Add a.Value, geçici & IIf(InStr(1, geçici, x, vbTextCompare) > 0, "", ";" & x)
            Else
                dict.Add a.Value, geçici & ";" & a.Offset(0, kaçıncıkolon - 1).Value
            End If
        End If
    Next a
  
çok_eşleşmeli_vlookup = dict(aranan.Value)
End Function </pre>
	            <%--<p id="link1"> --->vazgeçtim, çünkü, bu kısım için de yükleme yaptığı için masterdaki adres bilgiinin içini 
	            <script> değiştiriyor, prev/next'in çalışmasını bozuyor
		            //getContent("#link1","../Excel/FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx #cokeslesmelilookup_pre");
		            d("#link1").load("../Excel/FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx #cokeslesmelilookup_pre");
	            </script>
	            </p>--%>
        
</div>

	<h4 class="baslik">Gün nosundan(1-7) gün adını veren fonksiyon</h4>
	<div class="konu">
	<p>Normade bunu aşağıdaki fomrülle yazabiliriz ama TEXT fonksiyonu bazen 
	karışık olabilmekte ve bazı kişelerin yıldızı bu fomrülle barışık 
	olmayabiliyor. </p>
	<pre class="formul">=TEXT(A1+1;"gggg") //A1'de 2 yazıyorsa  salı döndürür. 1 ekleme sebebi günlerin Ameirkan fomratına göre Pazardan başlamasıdır</pre>
	<p>Gün UDF'i çok daha basit gibi duruyor.</p>
	<pre class="brush:vb">
Function gün(hucre As Range)
	Select Case hucre.Value
	    Case 1
	        gün = "Pazartesi"
	    Case 2
	        gün = "Salı"
	    Case 3
	        gün = "Çarşamba"
	    Case 4
	        gün = "Perşembe"
	    Case 5
	        gün = "Cuma"
	    Case 6
	        gün = "Cumartesi"
	    Case 7
	        gün = "Pazar"
	    Case Else
	        gün = "*****Hata*****,Ben sadece 1 ve 7 arasındaki günler için çalışırım"
	End Select	
End function</pre>    	
	<p>Hücredeki formül</p>
	<pre class="formul">=gün(A1)</pre>
	</div>
	
	<h4 class="baslik">Uçhariçortalma ile&nbsp;bir veri kümesindeki uç 
	değerleri elimine edin</h4>
	<div class="konu">
	<p>Diyelim ki hedefleme veya tahminleme yapıyorsunuz. Hedef verdiğiniz 
	birimin (bölge/şube/bayi/mağaza/v.s) son 1 yıllık satışlarına bakacak ve 
	ortalama alacaksınız, ama bir seferlik aşırı düşük/büyük rakamların da 
	ortalamayı etkilemesini istemiyorsunuz. O yüzdeb en küçük ve en büyük 
	rakamları çıkartıp kalan 10 ayın ortalamasını almak istiyorsunuz. Bunun
	için normalde şu Excel fonksiyonunu yazardınız.
	<pre class="formul">
=(SUM(B2:B13)-SMALL(B2:B13;1)-LARGE(B2:B13;1))/(COUNT(B2:B13)-2)	</pre>
<p>Ama biz şu UDF ile işimizi kısaca hallederiz.</p>
	<pre class="brush:vb">
Function uçhariçort(alan As Range, Uç As Variant)
Dim aratoplam As Double
Dim enbüyükler As Double
Dim enküçükler As Double

For i = 1 To Uç
    enbüyükler = enbüyükler + WorksheetFunction.Large(alan, i)
Next i
    
For i = 1 To Uç
    enküçkler = enküçkler + WorksheetFunction.Small(alan, i)
Next i
     
aratoplam = WorksheetFunction.Sum(alan) - enbüyükler - enküçkler
uçhariçort = aratoplam / (alan.Count - Uç * 2)

End Function</pre>
<p>UDF'imizi <strong>=uçhariçort(B2:B13;1)</strong> şeklinde kullanınca sonuç aşağıdaki gibidir.</p>
<p><img src="/images/vbaudfuçhariç.jpg"></p>
		<p>Sona girdiğimiz 1 parametresi en küçük/büyük 1 değeri hariç tutmamızı 
		sağlar. Eğer 2şer küçük/büyük hariç tutmak istersek buraya 2 gireriz, 
		bu sefer kalan 8 ayın ortalaması alınmış olur.</p>

</div>
	<h4 class="baslik" id="supertrim">Süpertrim	ile TRIM yapılamayan karakterleri de silin</h4>
		<div class="konu">
	<p>Bu fonksiyon Outlook gibi farklı bir ortamdan kopyalanarak alınan 
		rakamların önündeki ve sonundaki boşlukları yoketmek için kullanılır. 
		Genelde herkes Excel'in TRIM fonksiyonunu bilir ve bununla bu boşlukları 
		kaldırabileceğini düşünür ancak aslında bunlar her zaman bildiğimiz boşluk olmayabiliyor, 
		o yüzden TRIM yetersiz kalıyor.
	<a href="../Excel/FormulasMenusuFonksiyonlar_MetinselFonksiyonlar.aspx#supertrim">Metinsel fonksiyonlarda</a> bu
	 durumu açıklamıştık. Detayı oradan okuyabilirsiniz. İşte bu linkte verilen fomrülü yazmak 
	 yerine aşağıdaki UDF oldukça iş görecektir</p>
	
	<pre class="brush:vb">
Function süpertrim(hucre As Range)
    süpertrim = Val(WorksheetFunction.Trim(WorksheetFunction.Substitute(hucre, Chr(160), "")))
End Function
</pre>
</div>
	

	<h4 class="baslik">Sapmaoranı ile veri kümesindeki sapmayı bulun</h4>
	<div class="konu">
	<p>Bu fonksiyon ile incelediğiniz küme içinde bi dengesizlik var mı, rakamlar birbirinden çok mu farklı
	yoksa makul sınırlarda mı sapma var, bunu görmümş olursunuz. Yapılan işlem aslında kümenin standart sapmasının
	ortalamaya bölümüdür. Bu oran ne kadar küçükse o kadar dengelidir, yani rakamlar birbirine yakındır, ne kadar 
	büyükse sapma miktarı o kadar yüksektir. Tek elemanlı bir kümede hata alınır, siz eleman sayısının tek 
	olup olmamasına bakarak fonksiyonu iyileştirin.</p>
	<pre class="brush:vb">
Function sapmaoranı(alan As Range)
    sapmaoranı = WorksheetFunction.StDev_S(alan) / WorksheetFunction.Average(alan)
End Function</pre>
		<p>

		Aşağıdaki örnekte şubelerdeki müşteri temsilcilerine bağlanmış müşteri 
		adetleri görülmekte. 2 nolu şubede portföyler arasında dengesizlik 
		sözkonusu ve müşterilerin yeniden dağıtılmasında fayda var diyebiliriz.</p>
		<p>

		<img src="/images/vbaudfsapma.jpg"></p>
</div>

</div>
</asp:Content>
