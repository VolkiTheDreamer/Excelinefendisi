<%@ Page Title='DortTemelNesne Worksheet' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr>
<td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' 
runat='server' Text='Dört Temel Nesne'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'>
</asp:Label></td></tr></table></div>
<h1>Worksheet</h1>
    <h2 class="baslik">Giriş</h2>
    <div class="konu">
<p> Excel Nesne Modelinde Workbooktan sonraki seviyede sayfalar gelir. Excelde 
4 tip sayfa vardır. Bunlar;</p>
	<ul>
		<li>Çalışma sayfası(<strong>Worksheet</strong>:En çok bunları kullanacağız)</li>
		<li>Chart Sheet(Başlı başına bir sayfa olan grafik sayfaları)</li>
		<li>Macro Sheet(Bildiğimiz makro değil, eskiyle uyumluluk adına duruyor, biz 
		bunlara hiç girmeyeceğiz)</li>
		<li>Dialog Sheet(Eskiyle uyum adına duruyor, bunlara da girmeyeceğiz)</li>
	</ul>
	<p> Biz en çok <span class="keywordler">Sheet </span>ve
	<span class="keywordler">Worksheet </span>nesnelerini kullanacağız. 
	Yukardaki 4 maddeden anlaşıldığı üzere <strong>Worksheet</strong> nesnesi,
	<strong>Sheet</strong> nesnesinin bir alt türü oluyor. Sheet nesnesi, Sheets 
	koleksiyonunun, Worksheet nesnesi de hem Worksheets koleksiyonunun hem de 
	Sheets koleksiyonunun bir 
	üyesidir. Bu koleksiyonlar bir dosyadaki tüm sayfaları ifade eder ve genelde 
	döngülerde (Döngü konusunu hiç bilmiyorsanız en azından bir <strong>For Next</strong> için
	<a href="KosulifadeleriveDonguler_ForDonguleri.aspx">buraya</a> bakıp tekrar 
	gelin), veya toplam sayfa adedini saydırma gibi özelliklerle birlikte 
	kullanılır. Diyelim ki dosyamızda 3 çalışma sayfası bir de grafik sayfası 
	varsa <strong>Sheets.Count</strong> 4 sonucunu döndürürken <strong>Worksheets.Count 
	</strong>3 sonucunu 
	döndürür.</p>
	<p> Bunu kendiniz de deneyebilrisinz. Yukardaki özelliklerde bir dosya 
	hazırlayın ve VBE'de şu kodları çalışıtırın</p>

	<pre class="brush:vb">
Debug.Print ActiveWorkbook.Sheets.Count
Debug.Print ActiveWorkbook.Worksheets.Count
</pre>

	<p> <strong>Not</strong>:Biz birçok yerde Sheets ve Worksheets'i birbiri yerine kullanacağız.</p>
        </div>


	<h2 class="baslik">Temel işlemler</h2>
	<div class="konu">
	<h3> Sayfalara erişim ve referans</h3>
	<p> Sayfalar bir koleksiyon üyesi oldukları için onlara koleksiyonun <strong>item</strong> 
	özelliği ve bu özelliğin index numarası ile ulaşabiliriz. Item özelliği 
	koleksiyonların default özelliği olduğu için tüm diğer koleksiyonlarda 
	olduğu gibi bunda da yazılmadan es geçilebilir.</p>
	<p> Yani <strong>Worksheets.Item(1)</strong> ile <strong>Worksheets(1)</strong> tamamen özdeştir. Keza Workbooks.Item(1).Sheets.Item(2) 
	ile Workbooks(1).Sheets(2) de özdeştir. Gördüğünüz üzere başka bir 
	Workbooktaki bir sayfaya da Workbook ile sayfa arasına nokta koyarak 
	ulaşıyoruz. Koleksiyonlarda indexler 1'den başlar(En soldaki sayfanın index değeri 1'dir.)</p>
	<p>Sayfalara index numarasıyla olduğu gibi sayfa adı ile de, Worksheets("ilksayfa") 
	gibi, ulaşabiliriz.</p>
	<p>Mevcutta aktif olan sayfaya da <span class="keywordler">ActiveSheet
	</span>ifadesi ile ulaşırız. Activesheetle ilgili detaylara
	<a href="#acitvehseet">aşağıda</a> değineceğiz.</p>
	<h3>Seçme ve Aktive etme</h3>
	<p>Range nesnesinde nasıl bir veya birden çok hücreyi 
	seçmek için <span class="keywordler">Select</span> metodu kullanılıyorsa sayfalar için de aynı metod 
	kullanılır. Yine Range'te olduğu gibi tek bir hücreyi aktive etmek için
	<span class="keywordler">Activate</span> metodu kullanılıyordu, burda da yine aynı metod tek bir sayfayı aktive 
	etmek için kullanılır.</p>
	<p>Aşağıdaki kod ile tüm sayfalarda dolaşıyor ve her sayfanın ilk hücresine 
	artan bir şekilde sıra numarası yazıyorum.</p>

<pre class="brush:vb">
For Each ws In ActiveWorkbook.Worksheets
	ws.Select
	Range("A1").Value=i
	i=i+1
Next
</pre>

	<p><span class="keywordler">Activate</span> gizli olan sayfaların seçimi için 
	de kullanılabilirken, <span class="keywordler">Select
	</span> ile gizli sayfalar seçilemez. Mesela şimdi bi dosya açın, 5 sayfası olsun, 
	2.sini gizleyin. Sonra aşağıdaki kodu çalıştırın.</p>
	<pre class="=brush:vb">
Sub gizlisec_activateli()
Dim ws As Worksheet ' bu sefer ws'yi Worksheet olarak tanımladım
For Each ws In ActiveWorkbook.Sheets
    ws.Activate
    Range("a1") = ws.Index
Next
End Sub
</pre>
	<p>Bu kod ile tüm sayfalarda A1 hücresinin dolu olduğunu görürsünüz, 
	2.sayfayı unhide edip kontrol edebilirsiniz.</p>
<pre class="brush:vb">
Sub gizliyisec1()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
	ws.Select 'gizli olduğu için hata verir
	Range("a1") = ws.Index
Next
End Sub</pre>
	<p>Bu yukardakini "sayfa visible mı?" diye kontrol ederek de yapalım, bu sefer hata 
	almayız.</p>
<pre class="brush:vb">
Sub gizliyisec2()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
	If ws.Visible then 
   	  Range("a1") = ws.Index
   	   ws.Select 
	End If
Next
End Sub
</pre>
	<h3>İsimlendirme</h3>
	<p><span class="keywordler">Name</span> özelliği ile sayfanın ismini elde 
	eder veya onu değiştiririz. 
	Yani hem okunan hem yazılan(sets and returns) bir özelliktir. Index ise değiştirilemez 
	bir özellik olup sadece okunurdur.</p>
	<pre class="brush:vb">ws.Name="Yeni sayfa"</pre>
	<h3>Yeni sayfa açma(ekleme) ve silme</h3>
	<p>Yeni sayfa açmak/eklemek için <span class="keywordler">Add </span>metodu 
	kullanılır.</p>
	<p><span class="keywordler">Syntax:Worksheets.Add(Konum)</span> Konum 
	belirtmezsek aktif sayfanın soluna ekler. Konum belirtmek için 
	parametreleriyle birlikte kullanırız. <strong>Worksheets.Add 
	Before:=Worksheets("Sheet3")</strong></p>
	<p>Sayfayı yaratırken aynı anda ismini de verebiliriz.</p>
	<pre class="brush:vb">Worksheets.Add(Before:=Worksheets("Sheet2")).Name =&nbsp;"Yeni Sayfa"</pre>
	<p>Sayfaları silmek için <span class="keywordler">Delete </span>metodunu 
	kullanırız.</p>
		<pre class="brush:vb">Worksheets(1).Delete</pre>
	<p>Silme işlemlerinde Excel bize uyarı çıkarır, silmek istediğimizden emin 
	miyiz diye. Bu tür uyarılar bazen uzun makrolarda sorun çıkarabilir, hele 
	bir de bilgisayarımızın başında değilkenki bir saate ayarlanmış bir kod ise 
	biz gelip müdahale edene kadar ekranın takılı kalmasına neden olur, ve varsa 
	sonrasında ayarlanmış kodların da çalışmasını engellemiş our. 
	Bu tür durumlarla karşılaşmamak için kodların başına <strong>
	Application.DisplayAlerts=False</strong> cümleciği yazılır, ilgili kodlar 
	bittikten sonra da True'ya çevrilir. Bunun detaylarına Application konusunda 
	tekrar geleceğiz.</p>
	<h3 id="sayfagizleme">Sayfaları Gizleme ve Gizleneni tekrar gösterme</h3>
	<p>
	Sayfaları gizlemek veya tekrar göstermek için beklenenin aksine bir eylem değil
	<span class="keywordler">Visible</span> adında bir özellik 
	kullanıyoruz. Bu özellik booelan tipinde değer döndürür.</p>
	<pre class="brush:vb">Worksheets("Sheet2").Visible = False/True</pre>
	<p>Bir diğer yöntem de bu özelliğe True/False atamak yerine
	<span class="keywordler">XlSheetVisibility</span> enumerationlarını kullanmak 
	olabilir. Alacağı değerle şöyledir.</p>
	<ul>
		<li><strong>xlSheetVisible</strong>:Sayfayı gösterir</li>
		<li><strong>xlSheetHidden</strong>:Sayfayı gizler</li>
		<li><strong>xlSheetVeryHidden</strong>:Sayfayı öyle bir gizler ki, 
		kullanıcılar sayfalara sağ tıklayıp Unhide tuşuna bastıklarında bile gizli 
		sayfa listesinde görünmez. Bunu nerede kullanmak isteyebilirsiniz? 
		Mesela workbook/worksheet koruma şifrenizi veya bir veritabanı bağlantısı için 
		yazdığnız Connection String içindeki bağlantı parolasını VBA kodu içine 
		doğrudan yazmak istemiyorsunuzdur. Bunu bir xlVeryHidden nitelikli sayafanın 
		bir 
		hücresine yazıp buradan okutturabilirsiniz.</li>
	</ul>
	<p>Bu konuyla ilgili örnekler için <a href="#sifregizleme">tıklayınız</a>.</p>
	<h3>Taşıma ve Kopyalama</h3>
	<p><span class="keywordler">Move</span> ve <span class="keywordler">Copy
</span> metodularını kullanırız. İkisi de Before(Önce) ve 
	After(Sonra) olmak üzere iki parametre alır. </p>

<pre class="brush:vb">Worksheets("Sheet3").Move After:=Worksheets("Sheet1")
Worksheets("Sheet3").Copy Before:=Worksheets("Sheet1")
Worksheets("Sheet1").Copy Before:=Workbooks("ExcelVBA.xlsm").Sheets("Sheet3")
Worksheets("Sheet1").Copy 'yeni bir dosya açıp ve direkt bu dosyaya kopayalar
</pre>

<p>Bu metodları tek başına kullandığınızda yeni bir dosya açar ve oraya 
	taşır/kopyalar. Aşağıdaki <a href="#ornek1">şu örnekte </a>bu özelliği 
	kullanıyoruz.</p>
		<p>Birkaç kopyalama ve taşıma örneğini macro recorder ile kendiniz de 
	yapabilirsiniz. Basit bir konu olduğu için daha fazla detaya gerek görmedim.</p>
	<h3>Sayfalarda koruma (Protection &amp; Unprotection)</h3>
	<p>Yazdığımız kodlar, Sayfa korumalı bir dosyada işlem yapmaya kalkarsa şöyle 
	bir hata alırız<span>: <strong>"Run-time error 
	'1004': Application-defined or object-defined error"</strong>.</span></p>
	<p>
	Bu sorunun üstesinden gelmek için, öncesinde sayfanın korumalı olup 
	olmadığını kontrol edebilir, varsa korumayı kaldırabiliriz. Ama bunun da 
	bir sakıncası olabilir, o da şu ki, kodlarda ilerlerken başka bir hata 
	çıkarsa program durur ve sayfamız korumasız kalır. Bunun için bir de Hata 
	Yakalama kodu yazmamız ve ilgili yerde tekrar koruma koymamız gerekir.</p>
	<p>
	Korumaya almak <span class="keywordler">Protect</span>, korumayı kaldırmak
	<span class="keywordler">Unprotect</span> metodu ile sağlanır.</p>
	<p>
	Şimdi bir örnek yapalım. Yeni bi dosya açın ve ilk sayfasına koruma uygulayın, 
	şifresi de 1234 olsun. Sonra da şu kodu ekleyip çalıştırın.</p>
	<pre class="brush:vb">Sub sayfakoruma1()
Dim ws As Worksheet
Set ws = ActiveSheet

If ws.ProtectContents = True Then 'koruma var mı diye sorguluyoruz
   ws.Unprotect Password:="1234"
End If

'On Error GoTo hatayakala 'herhangi bir hatada sayfayı tekrar korumaya almak için ilgili yönlendirmeyi yapıyoruz

Range("A1") = Environ("USERNAME")
'çeşitli işlemler
'0'a bölme olacak ve hata alacak
a = InputBox("Bir sayı girin")
b = InputBox("Bu sayıyı kaça bölelim")
MsgBox "Sonuç: " & a / b

ws.Protect Password:="1234"
Exit Sub

'hatayakala:
'ws.Protect Password:="1234"

End Sub</pre>
	<p>
	İlk olarak, a için abc değerini girin, b için bir sayı girin. İkinci 
	denemede de b için 0 değerini girin. İlki sayısal bir değer girmediğmiz için, 
	ikincisi de 0'a bölmeye çalıştığı için hataya neden olacaktır. Hata yakalama kodları commentli olduğu için de dosyamız protectionsız kalacaktır.</p>
	<p>
	Şimdi yukardaki commentli kısımları commentsiz hale getirip tekrar 
	çalıştıralım. Bu sefer hata yakalanacak ve çıkmadan önce tekrar şifreleme 
	yapılacaktır.</p>
		<p>
		<strong>NOT</strong>:Filtrelemeye izin verme gibi seçeneklerden 
		faydalanmak icin makro kaydetme aracından faydalanabilirsiniz. Ancak bu 
		şekilde protection koymaya çalıştığınızda recorder şifreyi kaydetmez, 
		bunu manuel eklemeniz gerekir.</p>
	</div>
	<h2 class="baslik">İleri Seviye İşlemler</h2>
	<div class="konu">
	<h3>Protectionlı sayfalara devam</h3>
	<p>
	Yukarda korumalı sayfalarda çalışma şeklini görmüştük. Korumalı sayfalarla 
	çalışmanın daha şık bir yolu var aslında. <span class="keywordler">
	UserInterFaceOnly</span> parametresi.</p>
	<p>
	Bu parametreye True değeri atanırsa, sayfa sadece kullanıcı işlemlerine karşı 
	korumalı olur, VBA kodları için koruma geçersiz olur. Gördüğünüz gibi 
	Microsoft'taki abiler bunu da düşünmüşler. Bu parametrenin default değeri 
	False olup değer atanmaması durumunda tahmin edeceğiniz üzere full koruma 
	sağlanır.</p>
	<p>
	Yanlız unutulmamalıdır ki, bu argüman sadece bir kereliğine koruma sağlar. Dosyayı kapatıp tekrar 
	girdiğinizde, full korumalı açılır. Bu özelliğin daimi olması için Workbook_Open eventi içine uygun kod yazılabilir.</p>
	<pre class="brush:vb">Private Sub Workbook_Open()
Dim ws As Worksheet
   For Each ws in Activeworkbook.Worksheets
	ws.Protect Password:="1234",UserInterFaceOnly:=True
   Next ws
End Sub</pre>
	<p>
	Bu arada olur da sayfanın korumasının ne şekilde yapıldığını görmek 
	isterseniz <span class=" keywordler">ProtectionMode</span> özelliğine bakmanız gerekir. True ise 
	UserInterFaceOnly moduyla korunmuştur, false ise full koruma vardır.</p>
	<p>Akabinde kodla sayfa içinde bir yerlere birşeyler yazdırabilirsiniz. 
	Mesela bi veritabanından bir veriyi refresh edebilir, refresh işleminin 
	tarih ve saatini bir hücreye yazdırabilirsiniz.</p>
		<p>Bu özelliğin ne için kullanıldığını hala gözünde canlandıramamış olanlar için şöyle 
		belirtelim: Bunun daha zahmetli alternatifi şudur: Korumayı geçici 
	süre kaldırmak, gerekli hücrelere verileri yazdırmak, sonra korumayı geri 
	koymak olurdu. Zahmetli olduğu kadar risklidir de, zira korumayı geri aktive 
		etmeyi hatırlayacağınızın bir garantisi yok.</p>
	<h3>Ş<a name="sifregizleme"></a>ifrenin VBE'de görünmemesi</h3>
	<p>Korumalı bir sayfada korumayı koyma/kaldırma işlemlerini VBA kodları 
	şeklinde yapıyorsanız, meraklı bir kullanıcı, birazcık VBE ortamını da 
	biliyorsa, bir şeklide VBE'ye girip kodları karıştırırsa şifreleri görebilir. 
	Zira Protect ifadesinden sonraki Password:=""" parametresi ayan beyan görünmektedir. Bunun için VBE üzerinden kodlarınızın görünmemesi için bir şifre koymanız gerekebilir.</p>
	<p><img src="/images/vbaprotectvbe.jpg"></p>
	<p>Ancak diyelim ki bir nedenle VBA protection koymak istemiyorsunuz, 
	Workbook Protection da koymak istemiyorsunuz, bu durumda diğer 
	alternatif de şu olabilir: Şifreyi bir sayfanın A1(veya istediğiniz herhangi bir) hücresine koyup, gereken 
	yerlerde şifreyi buradan okumaktır. Sonra bu sayfayı da <strong>xlVeryHidden</strong> olarak 
	saklamak, böylece kodları kurcalayan kişi "ya burda Sheets(1)'nin A1 
	hücresinden şifreyi alınıyor görünüyor, ama A1 hücresi boş, bu ne biçim iş anlamadım ben" diye 
	yakınıp durur. Zira deneyimli bir makro kullanıcısı olmadığını varsayarsak kendisinin 
	xlVeryHidden sayfa türünden haberi olmayacaktır. "Gizli bir sayfa mı acaba?" 
	diyip sayfa sekmelerine sağ tıklayıp Unhide dediklerinde de bu sayfayı orada 
	göremeyecekler ve çıldırıp duracaklardır.</p>
	<p>Bunun için ya şu kodu yazacağız, </p>
	<pre class="brush:vb">Sub şifreleme()
   Sheets(1).Visible = xlVeryHidden
End Sub</pre>
	<p>Ya da VBE'de Project penceresinden ilgili sayfayı seçip Properties 
	penceresinden de şu ayarı yapacağız:</p>
	<p><img src="/images/vbaprotectveryhidden.jpg"></p>
	<p>Kod yazarak yapma yolunu seçersek bunu hemen silelim ki bir başkası böyle bir sayfanın 
	varlığından haberdar olmasın. Tabi umarız ki kullanıcı, bu Project 
	penceresindeki şifre sayfasını görüp de Excel önyüzünde göremeyince 
	işkillenmez. </p>
		<p>Şimdi bir örnekte gösterelim. </p>
		<ul>
			<li>Yeni bir dosya açın</li>
			<li>2.sayfa yoksa yaratın ve bunu <strong>xlVeryHidden</strong> olarak ayarlayın</li>
			<li>2.sayfann A1 hücresine bi şifre yazın</li>
			<li>Şifre yazdığınız hücreye Name olarak şifre adını verin, 
			formatını beyaz renk yapın görünmez olsun, ayrıca <strong>Format 
			Cells&gt;Protection</strong> menüsünden Hidden diye de işaretliyin</li>
			<li>İlk sayfaya Developer menüsünde bi düğme ekleyin</li>
			<li>Düğmenin click eventine aşağıdaki kodu yazalım</li>
			<li>Sayfaların <a href="#codename">CodeName</a>'lerini tersine 
			çevirin ki iyice kafa karıştırsın. (1.sayfanın adını Sheet2, 
			2.sayfanın adını Sheet1 yapıyoruz)</li>
		</ul>
		<p><img src="/images/vbaworksheetprotect.jpg"></p>
		<pre class="brush:vb">
Sub Button1_Click()
    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect [şifre] 'Name'e az kullanılan bu yöntemle ulaşarak daha fazla kafa karıştırıyorum
        ActiveSheet.Shapes("Button 1").OLEFormat.Object.Text = "Koruma kaldırıldı"
        Sheet1.Unprotect [şifre] 'şifre sayfasına code adıyla ulaşıp daha da kafa karıştıryoruz ki kuracalayan kişi artk vageçzsin
    Else
        ActiveSheet.Shapes("Button 1").OLEFormat.Object.Text = "Korumaya alındı"
        ActiveSheet.Protect [şifre]
        Sheet1.Protect [şifre]
    End If
End Sub
</pre>
	<p>Bu yöntemi, kırması zor olsun diye çok uzun bir şifre belirlediğinizde de 
	kullanabilirsiniz. Uzun şifreyi xlVeryHidden sayfaya yazarsınız. Ana sayfada 
	da bir hücre belirlersiniz, oraya çift tıklandığında dosya Unprotect olur, 
	böylece uzun şifreyi manuel girmek zorunda kalmazsınız.</p>
		<p>Bir diğer alternatif de Workbook koruması koymaktır. Şifreyi de yine herhangi bir 
	gizli sayfaya koyabilirsiniz, bu sefer VeryHidden olmasına gerek yok, zaten 
	sayfaları hide/unhide özelliği tamamen pasif olacak.</p>
	<p>Bu yöntemler ayrıca, DB connectionların şifrelerini almak için de 
	kullanılabilir, ancak ribbonun da pasifize edilmesi lazım, yoksa ribbondaki 
	Properties'ten de girip şifreyi görebilirler. Veya datayı çektikten sora 
	ilgili Table'ı Range'e döndümeniz gerekir ki şifre görünmesin. Bunları,
	<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">yeri 
	geldiğinde</a> detaylıca göreceğiz.</p>
	<h3>
	<strong id="acitvehseet">ActiveSheet</strong></h3>
	<p>
	Bir sayfanın özelliğiyle ilgili bir ifade yazıyorsak ve bu özelliğin önünde sayfa ismi yoksa, Excel bunun o anda aktif olan sayfanın bir özelliği 
	olduğunu anlar. Bu bağlamda ActiveSheet.Range("A1") ile Range("A1") tamamen aynı şey olarak algılanır. </p>
	<p>
	Ancak böyle Worksheet ifadesini belirtmeden kod yazmanın bir istisnası var. 
	Eğer, makronuz süreç içinde Sayfa modüllerinden birine girerse, veya 
	direkt oraya yazdığınız bir makroyu çalıştırıyorsanız ve o anda başka bir 
	sayfadaysanız, kodlar işlemi o aktif sayfada değil, ilgili sayfa modülünün 
	sayfasında yapar. Örneğin Sheet1, Sheet2, Sheet3 şeklinde 3 sayfanız olsun. 
	Sheet1 modül sayfasında şu kodu yazalım;</p>
<pre class="brush:vb">
Sub deneme()
  Range("a1") = 2
End Sub
</pre>
	<p>Şimdi 2.sayfayı aktif hale getirelim ve VBE'ye geçip Sheet1 modülündeki 
	kodu çalıştıralım. </p>
	<p>Gördüğünüz gibi o an aktif sayfa 2.sayfa olmasına rağmen 1.sayfanın A1 
	hücresine 2 değeri yazıldı.</p>
	<p>O yüzden sadece Workbook modülü ve standart modüllerdeki kodlarda sayfa 
	ismi belirtilmediyse Activesheet olarak algılanma sözkonusudur diyebilirz.</p>
	<h4>Activesheet ve Dönen Değer</h4>
	<p>Aktif nesnelerle ilgili önemli bir konu da, nesnenin dönen değerine göre 
	intellisense'in çalışıp çalışmayacağıdır. Eğer nesnenin dönen değeri bir Obje 
	tipinde ise intellisense çalışmaz, obje tipinde değil de spesifik bir nesne 
	tipinde ise intellisense çalışır. Mesela Activeworkbookun dönen değeri bir 
	Workbook olup, hemen arkasından bir nokta(.) yazınca intellisense çalışır 
	ve bu nesneye ait özellik ve metodlar listelenir.</p>
	<p><img src="/images/vbaaktifobjedonendeger1.jpg"></p>
	<p>Bu da ActiveWorkbookun intellisense görüntüsü:</p>
	<p>
	<img src="/images/vbaaktifobjedonendeger3.jpg"></p>
	<p>Activesheet'in ise dönen değeri Object'dir, bunu daha önce de görmüştük. Zira 
	activesheet normal bir çalışma sayfası(worksheet) da olabilir, bir grafik sayfası da. 
	Doğal olarak bunların özellik ve metodları da farklıdır. Birden fazla dönüş değeri 
	olan nesneler de Object tipli olmaktadırlar. O yüzden bunlarda intellisense 
	çalışmaz.</p>
	<p><img src="/images/vbaaktifobjedonendeger2.jpg"></p>
	<p >Activesheetin intellisense'i de bu nedenle çıkmamaktadır:&nbsp;</p>
	<p>	<img src="/images/vbaaktifobjedonendeger4.jpg"></p>
	<p>	Peki Activesheet'e ait intellisensi çıkarmanın hiç mi yolu yok? Tabi ki var: 
	Onu bir Worksheet nesnesine atamak.</p>
	<p>	<img src="/images/vbaaktifobjedonendeger5.jpg"></p>
	<h3>
	<strong>Sayfal<a name="codename"></a>arın kod ismi </strong>ile gerçek ismi</h3>
	<p>
	Sayfalar yaratıldıkları anda "Sheet 1", "Sheet 2" gibi isimlere sahiptirler. 
	Kullanıcı isterse bunların isimlerini ve sırasını sonradan değiştirebilir.
	</p>
	<p>
	Sayfaların bir de, yaratıldıkları anda sahip oldukları Sheet1(bu sefer bitişik 
	yazılan), Sheet2 gibi kod adları vardır, bunlar da VBE'deki properties 
	penceresinden yeniden adlandırılabilir, ama pratikte daha çok sayfaların Excel arayüzündeki 
	gerçek isimlerinin yeniden adlandırıldığı görülür. Sayfaların Exceldeki 
	sırası değişse bile kod isimlerinin ismi değişmediği için sırası aynı 
	kalır. Aşağıdaki görüntülerde bunu görebiliyoruz. İlk sayfa gizlenmiş, ikinci 
	ve üçüncü sayfa ise yer değiştirmiş.</p>
	<p>
	<img src="/images/vbasheetnamecodename1.jpg"></p>
	<p>
	Ancak gördüğünüz gibi kod isimleri aynı duruyor.(Tabiki kod isimleri de değişirse bunlar da 
	alfabetik olarak sıralanacaktır)</p>
	<p>
	<img src="/images/vbasheetnamecodename2.gif"></p>
	<p>
	İşte sayfalara ulaşmanın bir şeklini daha görmüş olduk, kod isimlerle. Ancak 
	dikkat edilmesi gereken bir nokta var, o da kod isimlerin sadece&nbsp; 
	ilgili dosya içinden(normal modül veya thisworkbook ile sheet modülleri) 
	ulaşılabilir olduğudur. Şimdi 
	tüm yöntemleri birarada görelim;</p>
	<pre class="brush:vb">Sub sayfalara_erişim()
	Sheet1.Select 'kod isimle. (sadece bunda intelisense çıkar)
	Sheets(1).Select 'Sheets koleksiyonu ve index
	Worksheets(1).Select 'Worksheets koleksiyonu ve index
	Sheets("bireysel").Select 'Sheets koleksiyonu ve sayfa adı
	Worksheets("bireysel").Select 'Worksheets koleksiyonu ve sayfa adı
End Sub</pre>
	<p>Bu konuyla ilgili güzel bir makale, üstatlardan C.Pearson'a ait bu
	<a href="http://www.cpearson.com/Excel/RenameProblems.aspx">linkte</a> var.</p>

</div>
	<h2 class="baslik">Filtre ve Sıralama işlemleri</h2>
<div class="konu">
	<p>Sıralama ve Filtreleme işlemleri her ne kadar hücreler üzerinde yapılıyor 
	olsa da Worksheet nesnesinin özellikleridirler. Öyle çok kompleks bir 
	tarafları olmadığı için Macro Recorder ile kaydedilmiş bir kod üzerinden 
	inceleyebiliriz.</p>
		<pre class="brush:vb">
ActiveWorkbook.Worksheets("calculatedlar").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("calculatedlar").Sort.SortFields.Add Key:=Range( _
        "F2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
With ActiveWorkbook.Worksheets("calculatedlar").Sort
        .SetRange Range("A2:F175")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With</pre>
    
    <p>Recorder'ın ürettiği kodda bir iki yeri değiştirmek gerekir, böylece makromuz daha dinamik hale gelir.</p>
		<ul>
			<li>Worksheets("calculatedlar")&gt;ActiveSheet yapalım</li>
			<li>Key:=Range("F2")&gt;Key:=ActiveCell(veya duruma göre kalabilir)</li>
			<li>Range("A2:F175")&gt;Bunu nasıl değiştireceğimizi
			<a href="DortTemelNesne_Range.aspx#Ornekler">Range</a> konusunda 
			görmüştük</li>
		</ul>
		<p>Keza filtreleme işleminin kodu da basit olup, dinamik hale 
		getirilecek kısımları değiştirmek yeterlidir.</p>
		<pre class="brush:vb">Range("B1").Select
ActiveSheet.Range("$A$1:$F$175").AutoFilter Field:=2, Criteria1:="Şube"</pre>
		<p>Bu örnekte Range("$A$1:$F$175") yerine Range("A1").CurrentRegion 
		denilebilir. Görüğünüz gibi Sort işleminde başlık hariç tutulurken 
		burada başlık satırı dahildir. O yüzden Resize ve Offset kullanmadan 
		işimizi halledebiliriz. </p>
		<h3>Filtre modları ve filtreyi kaldırma</h3>
		<p>Filtrelerle ilgili olarak kafa karışıklığına neden olabilecek iki 
		konur var. Filtrenin açık/kapalı olması ile kriterin 
		uygulanıp/uygulanmamış olması. Bunlardan ilki <span class="keywordler">
		AutoFilterMode </span>özelliği ile elde edilirken ikincisi
		<span class="keywordler">FilterMode&nbsp; </span>özelliği ile elde 
		edilir.</p>
		<p>Filtreyi uygulamak/kaldırmak Range nesnesinin
		<span class="keywordler">AutoFilter </span>metodu ile olur. </p>
<ul>

<li>Filtre henüz yokken kullanılırsa filre uygulanır</li>
<li>Yanında parametre yoksa sadece filtre okları görünür</li>
<li>Parametreyle uygulanırsa direkt ilgili filtreleme yapılmış olur</li>
<li>Filtre varken kullanılırsa kriterler silinir ve filtre okları kalkar</li>
</ul>

<p>Kriter uygulanmış bir data kümesinde kriterleri silmek ama filtre 
		oklarını açık bırakmak için Worksheet nesnesinin
		<span class="keywordler">ShowAllData</span> metodu kullanılır.</p>
		<p>Aşağıda bütün bu durumları anlatan güzel bir örnek var. Bu örnekte 
		sayfadaki data kümesinin rasgele olarak 4 durumundan birine girmesini 
		sağlıyoruz, sonra hangi durumdaysa onunla ilgili bir mesaj veriyorum.</p>
		<pre class="brush:vb">
Sub filtresub()

a = WorksheetFunction.RandBetween(1, 4)
MsgBox "Case " & a & " gerçekleşecek"

Select Case a
    Case 1 'o an kapalıysa kapalı kalmaya devam, açıksa kapanır
        ActiveSheet.AutoFilterMode = False 'buna sadece false atanıyor, true atanamaz
    Case 2 'açıkken filtre konursa kapanır
        If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
    Case 3 'kapalıyken filtre konursa açılır
        If ActiveSheet.AutoFilterMode = False Then Selection.AutoFilter
    Case 4 'filtre okları aktif olsa da olmasa da burası çalışır
        ActiveSheet.Range("$A$1").CurrentRegion.AutoFilter Field:=2, Criteria1:=Range("b2")
End Select

If ActiveSheet.AutoFilterMode = True Then 'filtre okları açıksa. Case 3 veya 4
    If ActiveSheet.FilterMode = False Then
        MsgBox "Case 3:Filtre açık ama kriter yok"
    Else
        MsgBox "Case 4:filtre açık ve kriter var, şimdi kriter kaldırılacak, ama filtre açık kalacak"
        ActiveSheet.ShowAllData
    End If
Else 'filtre okları yoksa, yani henüz bir autofilter düğmesine basılmamışsa
    MsgBox "Case 1 veya 2:filtre yok"
End If		
End Sub</pre>
    
	</div>
	<h2 class="baslik">Diğer Önemli İşlemler</h2>
	<div class="konu">
		<h3>Sayfa ayarları</h3>
		<p>Page Setup, Print v.b işlemleri için macro recorderdan faydalanmanızı 
		öneriyorum.</p>
		<h3>Calculation</h3>
		<p>Calculation konusuna Application nesnesine detaylıca <span>
		değiniyoruz</span>. Önemli 
		bir metod olup Application konusunda mutlaka incelemenizi tavsiye 
		ediyorum.</p>
		<h3>Sayfalarda gezinme</h3>
		<p>Sayfalarda dolaşma işini yukarıdaki birçok örnekte gördüğümüz gibi 
		bir döngü aracılığı ile yapabilmekle birlikte, sadece bir sayfa geri 
		veya ileri gitme işini, sayfa adı v.s yazmadan kolayca yapmamızı 
		sağlayan iki özellik var.</p>

<pre class="brush:vb">
Activesheet.Next.Select ' sonraki sayfa
ActiveSheet.Previous.Select ' önceki sayfa
</pre>
		<h3 id="paste">Hafızadan birşey yapıştırma</h3>
		<p>Hafızadaki(Clipboard) bilgiyi aktif sayfadaki aktif hücreye 
		yapıştırmak için iki metod var. Birincisi Worksheet sınıfının normal <span class="keywordler">
		Paste </span>metdou, ikincisi Range nesnesinin <span class="keywordler">PasteSpecial
		</span>metodu. <strong>Worksheet.Paste</strong> metodu ile Excel arayüzünde yaptığımız gibi 
		içerikte ne varsa yapıştırılır:Veri, formül, format v.s. <strong>
		<a href="DortTemelNesne_Range.aspx#pastespecial">Range.PasteSpecial</a></strong> 
		ile ise yine Excel arayüzde yaptığımız özel yapıştırma türlerini 
		yapabiliyoruz. Sadece değerleri, sadece formatı v.s. </p>
		<p>Aşağıdaki kod ile 1.sayfada aktif hücrenin CurrentRegion'ındaki 
		hücreleri 2.sayfadaki aktif hücreye yapıştırıyoruz.</p>

<pre class="brush:vb">
 Sheets(1).Select
ActiveCell.CurrentRegion.Select
Selection.Copy
Sheets(2).Select
ActiveSheet.Paste
</pre>


		<p>Bu arada Worksheet sınıfının da PasteSpecial metodu var ama biz onu 
		çok kullanmayacağız, zira bununla Grafik gibi nesneleri veya Access gibi diğer uygulamalardan birşeyler yapıştırabiliyorsunuz.</p>
		<pre class="brush:vb"> ActiveChart.ChartArea.Copy
Range("M2").Select
ActiveSheet.PasteSpecial Format:="Picture (PNG)", Link:=False, _
DisplayAsIcon:=False</pre>

		<p>Aşağıdaki örneği de <a href="https://msdn.microsoft.com/en-us/library/office/ff835858.aspx">MSDN</a>'den aldım, burada diğer format seçeneklerine de bakabilirsiniz. Başka programlarla çalışmayacaksanız çok sık ihtiyacınız olacağını sanmam.</p>


<pre class="brush:vb">
Worksheets("Sheet1").Range("F5").PasteSpecial _ 
 Format:="Picture (Enhanced Metafile)", Link:=False,
 DisplayAsIcon:=False</pre>
		<h3>
		Parent özelliği</h3>
		<p>
		Bazen bir sayfanın hangi workbookta olduğunu elde etmek isteriz. Bunun 
		için hiyerarşide bir üst basamağa çıkmamızı sağlayan Parent özelliğini 
		kullanırız.</p>
		<pre class="brush:vb">Debug.Print TypeName(Activesheet.Parent) 'Workbook
Debug.Print Activesheet.Parent.Name 'ilgili Workbook'un adı</pre>


</div>



<h2 class="baslik">Çeşitli Örnekler</h2>
<div class="konu">
<h4 class="baslik"><a name="ornek1"></a>Tüm sayfaları ayrı workbooklar olarak kaydetme</h4>
<div>
<p>Bu örnekte açık olan bir dosyadaki tüm sayfaları ayrı 
ayrı dosyalar olarak seçilen bir klasöre kaydediyoruz.</p>

<pre class="brush:vb">
Sub sayfaları_wb_olarak_kaydet()

Dim ws As Worksheet
Dim wb As Workbook
Dim fd As FileDialog
Dim klasör As String
Dim dosya As String

On Error GoTo hata

Application.ScreenUpdating = False
Set wb = Application.ThisWorkbook


Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = True Then
   klasör = fd.SelectedItems(1)
Else
   Exit Sub
End If


For Each ws In wb.Worksheets
   ws.Copy
   'Burada bi versiyon kontrolü yapılmasında fayda var, ama örneği basitleştirmek adına onu es geçiyorum ve default formatta kaydetmesine izin veriyorum

   dosya = klasör &amp; "\" &amp; Application.ActiveWorkbook.Sheets(1).Name &amp; ".xlsx"
   ActiveWorkbook.SaveAs dosya
   ActiveWorkbook.Close False
Next


Call Shell("explorer.exe " &amp; klasör, vbNormalFocus)
Application.ScreenUpdating = True
Exit Sub

hata:
Application.ScreenUpdating = True
MsgBox Err.Description
End Sub
</pre>
</div>

<h4 class="baslik">Tüm sayfaları Unhide etme</h4>
<div>
<p>Bu örnekte tüm gizli sayfaları açıyoruz. Buna QuickAccessBardan erişmek isteyebilirsiniz, ben öyle yapıyorum açıkçası.</p>
<pre class="brush:vb">
Sub tümsayfalarunhide()
  For i = 1 To Sheets.Count
     Sheets(i).visible = True
  Next i
  Sheets(1).Select
End Sub</pre>
</div>


<h4 class="baslik">İlk sayfa hariç tümünü hide etme</h4>
<div>
<p>Diyelim ki bir sayfasında Karne, diğer sayfalarında bu karneyi besleyen sayfaların olduğu bir dosyanız var. Sık sık sayfaları Unhide ve tekrar Hide etme ihtiyacınız oluyor. Bir üstteki kod ile tüm sayfaları Unhide etmiştik. Datayla oynadıktan sonra şimdi tekrar gizleyeceğiz. Bunu da QuickAccessBara eklerseniz müthiş pratiklik sağlar. Bu kodu Karne sayfasındayken çalıştırmanız gerekiyor.</p>
<pre class="brush:vb">
Sub aktifhariç_hideall()
Dim ws As Worksheet
Set ws = ActiveSheet
For i = 1 To Sheets.Count
    If i <> ws.Index Then
        Sheets(i).visible = False
    End If
Next i
End Sub
</pre>
</div>


<h4 class="baslik">Tüm sayfalarda ilk kolonda A'dan Z'ye sıralama</h4>
<div>
<p>Diyelim ki size birçok sayfası olan bir dosya geldi. Ama sayfaların hiçbiri sıralı değil. Bu örnekte tüm sayfalarda A'dan Z'ye sıralama işlemini tek seferde yapıyoruz</p>
<pre class="brush:vb">
Sub tümsayfalarda_sırala()

For i = 1 To Sheets.Count
    Sheets(i).Select
    ActiveWorkbook.Sheets(i).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(i).Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Sheets(i).Sort
        .SetRange Range(Range("A2"), Range("a2").End(xlDown).End(xlToRight))
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Next i

End Sub
</pre>
</div>
	
</div>
</asp:Content>
