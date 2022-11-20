<%@ Page Title='Giris MakroKAydetme' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'>
</asp:Label></td></tr></table></div>
<h1>Makro Kaydetme ve VB Editörü</h1>
<h2 class="baslik">Makro Kaydetme</h2>
<div class="konu">
<p>İlk sayfada, bir makro yazmaya başlamadan önce bir makro kaydedip oluşan kodu inceleyerek de makro öğrenmeye başlayabileceğinizden bahsetmiştim. Şimdi gelin birlikte, bir makro nasıl kaydedilir ve yazılan kod üzerinden nasıl oynama yapılır ona bakalım.</p>
	<p>İlk olarak Developer menüsünde Record Macro diyoruz.</p>
	<p><span>Makrolarımızı genelde <strong>Personal.xlsb</strong> dosyasına kaydedececeğimiz için 
	depolama yeri olarak burayı seçiyoruz.</span></p>
	<p>
	<img src="/images/vba_giris_record.jpg"></p>
	<p>Şimdi basit bir dizi işlem yapalım.</p>
	<ul>
		<li>A sütununu seçelim</li>
		<li>Tüm sütunu bold ve rengini de kırmızı yapalım</li>
		<li>A1'e "Merhaba" yazalım</li>
		<li>Son olarak, Excelin versiyonuna göre pencerenin çeşitli yerlerinde 
		bulunabilecek olan Stop tuşuna basalım.( Benimki pencerenin sol alt 
		köşesinde)
		<img src="/images/Vbarecorderstop.jpg"></li>
	</ul>
	<p>Şimdi <strong>Alt+F11</strong> tuşlarına basarak veya Develpoer menüsünden(veya 
	QuickAccessToolbardan) VB editörünü açalım.</p>
	<p>Burda son yazdığımız makro hep en büyük numaralı ModuleX içine gider. 
	Ör:Module1, Module2,,,,Module5 varsa biz Module 5 içine gidip bakalım ve 
	kodumuzu görelim.</p>
	<pre class="brush:vb">
Sub Macro2()
'
' Macro2 Macro
'

'
    Columns("A:A").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("A1").Select
    Selection.FormulaR1C1 = "Merhaba"
End Sub
	</pre>
	<p>Gördüğünüz gibi makromuz<span>&nbsp;'Sub'&nbsp;</span>ifadesi ile 
	başladı, sonra makromuzun adı, sonra da () işaretleri geliyor. Arada kod 
	parçaları var, en son da '<span>End Sub</span>' ifadesi gelir.</p>
	<p>' işareti açıklama cümleleri içindir, kendinize makronun o bölgesiyle 
	ilgili hatırlatmalarda bulunmak isteyebilirsiniz. Bu açıklamalar oldukça 
	faydalı olabilmektedir, özellikle bir makroyu seyrek kullanıyorsanız, neyi 
	niçin yaptığınızı hatırlamak adına faydalı olmaktadır. Burada da, Excel'in 
	kendisi otomatik bir açıklama eklemiş olduğunu görüyorsunuz. Yorumlar yeşil 
	renkli görünürler. Kaydedilen 
	makrolardaki bu açıklamaları silebilirsiniz.</p>
	<p>İsterseniz bundan sonra kendiniz çeşitli satırlara müdahale edebilir, 
	bazılarını çıkarabilir, yeni satırlar ekleyebilirsiniz. Mesela bu makro her 
	çalıştığında A sütununu seçecektir, bunun yerine hep o anda seçili olan 
	sütunla ilgili işlem yapan bir makro yazmak isteyebilirsiniz, ki bu daha 
	işlevseldir. Bunun için, yukardaki gibi herhangi bir sütunu seçen bir makro 
	kaydedin, sonra kod kısmına giderek, Columns("A:A") yazan satırı çıkarın, 
	böylece direkt seçili olan şeyle (burada sütun seçilidir) işlem yapılır. 
	Sütun seçili değilse, ama bulunduğunuz hücrenin bağlı olduğu sütunun 
	seçilmesini isterseniz sütunlu ifadeyi şu şekilde değiştirmeniz yeterli 
	olacaktır.</p>
	<p><strong>Columns("A:A").Select=>ActiveCell.EntireColumn.Select</strong></p>
	<p>Bu konularda detaylı bilgi ilerleyen sayfalarda verileceği için şuanda 
	daha fazla ayrıntıya girilmeyecektir.</p>
	<p>Oluşturduğunuz makrolara kısayol tuşu da atayabilir, bunları sonradan 
	amacınıza uygun olarak değiştirebilir veya iptal edebilirsiniz. Şimdi bi de 
	hızlıca kısayol işlemleri nasıl yapılır ona bakalım.</p>
	<h2>Kısayol işlemleri</h2>
	<p>Bir makroya kısayol tuşunu makro kaydederken atamak oldukça kolaydır: 
	Aşağıdaki kutuya istediğiniz harfi koyabilirsiniz.(Harf dışında bir karakteri 
	kabul etmez.)</p>
	<p>
	<img src="/images/vbashortcut.jpg"></p>
	<p>Peki makro kaydediciyle oluşturmadığımız bir koda nasıl kısayol tuşu atarız. Basit:Developer menüsünden 
	Macros butonuna basalım, istediğimiz makroyu seçelim ve Edit diyelim. Eğer 
	Personal.xlsb gibi gizli bir dosyadaki makroyu editleyeceksek Excel buna izin 
	vermez. Dosyayı önce Unhide etmeli, arkasından makro ayarını yapmalıyız, o 
	da bittikten sonra tekrar gizlemeliyiz. </p>
	<p>Aşağıda çıkan kutuda aynı yere istediğimiz harfi yazalım. Bu arada farkettiyseniz bu ekranda sadece Ctrl tuşuyla 
	çalıştırılan kısayollar oluşturabiliyoruz. Farklı tuş kombinasyonlarına kısayol tuşu atamayı öğrenmek için biraz daha(nesnelerin efendisi olan Application nesnesini tanıttığımız sayfaya gelmeyi) beklemelesiniz.</p>
	<p>
	<img src="/images/vbashortcut2.jpg"></p>
	<p>
	Son olarak, kısayol atadığımız bir makrodan bu kısayol tuşunu nasıl 
	kadırırız? Basit:Yine Macros düğmesine basarız, yukardaki çıkan diyalog 
	kutusunda çıkan harfi sileriz.</p>
</div>


<h2 class="baslik">VBE(Visual Basic Editörü)</h2>
<div class="konu">

	<h2>Üç pencere</h2>
	<p>Henüz başlangıç aşamasındayken bir VB penceresi genel olarak aşağıdaki gibi 
	görünür. 1 numaralı bölmede makrolara ismen ulaşacağınız project penceresi bulunur, onun altında 
	2 numaralı bölmede Workbook/Worksheet veya modüllerin özelliklerini gösteren properties penceresi bulunur, 
	burayı genel olarak modüllerin adını değiştirmek için kullanacağız, onun 
	dışında kapalı tutabilirsiniz. Bizi ilgilendirecek kısım daha çok 
	3 nolu geniş alan olacaktır. Kodlarımızı burada yazacağız, bütün düzeltme ve çalıştırma işlemlerini 
	de bu alan üzerinden yapacağız.</p>
	<p>
	<img src="/images/vba_giris_editor.jpg" width="60%" heigth="60%" class="zoomla"></p>
	<p>Projects penceresiyle ilgili olarak söyleyebileceklerim şimdilik şu kadar 
	olacaktır. 'Bir makro nereye yazılır?'ın ilk cevabını burada belirledikten 
	sonra 3 nolu panele geçip yazmaya başlıyoruz. O yüzden kodumuzun ilk etapta neyin içine 
	yazılacak sorusunun cevabıdır. Seçenekler şöyle:</p>
	<ul>
		<li>Excel Objects içinde bir sayfa</li>
		<li>Excel Objects içinde ThisWorkbook</li>
		<li>(Genel) Modül</li>
		<li>Class Modül</li>
		<li>UserForm</li>
	</ul>
	<p>Microsoft Excel Objects altında Sheets ve ThisWorkbook nesneleri bulunur. 
	Sheets'lerden birinde yazılan makrolar genellikle Worksheet_SelectionChange 
	and Worksheet_Calculate gibi sayfa olaylarını yönetmek için kullanılırken, 
	ThisWorkbook içine yazılan makrolar dosya seviyesindeki olayları 
	ele almak için kullanılır, Open/Close gibi. Bu nesneler hakkında daha 
	detaylı bilgi ilerleyen <a href="Olaylar_Konular.aspx">bölümlerde</a> verilecektir.</p>

	<p>(Genel) Modüller, herhangi bir sayfa üzerinde çalışabilecek genel 
	prosedürleri depolamak için kullanılır. Vereceğimiz örneklerin çoğu bu 
	Modüller içine yazılacak makrolar şeklinde olacaktır.</p>

	<p>UserForm ve Class modüllere de ilerleyen bölümlerde değinilecektir.</p>

	<p><strong>Önemli Not:</strong><a href="DebuggingveHataYonetimi_BreakpointlerveIzlemePencereleri.aspx">İzleme 
	Pencreleri</a> bölümünde detaylı göreceğiz ancak, başlangıçta öğrenmeniz 
	gereken bir pencere daha var. Ctrl+G tuşlarına basarsanız en aşağıda
	<strong>Immediate Window</strong> çıkacaktır. Kodlarımız 
	içinde sık sık <span class="keywordler">Debug.Print</span> ifadesini 
	kullanacağız. Bu, kendisinden sonra gelen ifadenin değerini Immediate 
	Window'a yazan bir metoddur. O yüzden bu penecereniz de sürekli açık kalsın 
	derim.</p>
	<h2>Menüler</h2>
	<p>Excel'in kendi içindeyken Ribbondaki menülere sık sık başvururuz ancak açık söylemek 
	gerekirse VBE penceresindeki menülerle çok fazla işiniz olmayacak. Digital 
	İmza ayarlamasını yapmak ve projemize referans(dll) eklemek için Tools 
	menüsünü arada bir kullanacağız.</p>
	<p>Bir diğer kullanacağımız menü de View menüsü olacak. Burada Immediate, Local 
	gibi yardımcı pencereleri açacağız. Bunlara da sonrada değinilecektir.</p>
	<p>Bir de, özel Add-inler vardır ki, eğer bunlardan birini satın aldıysanız 
	veya ileri seviyelere geldiğinizde kendiniz yaptıysanız(bunun için VBA'in 
	dışında şeyler bilmek gerekir) buna ait menüleri de kullanabilirsiniz. Eğer maddi 
	durumunuz uygunsa, şu <a href="http://codevba.com/"> sayfadaki</a> gibi bir 
	Add-in kurmanızı tavsiye ederim. Mesela bu linkteki Add-in hem 
	muhteşem bir pratiklik sağlıyor hem de hatasız kod yazılmasını sağlıyor. 
	Bence bu linki bir yere not edin ve kendinizi biraz geliştirdikten sonra ilk 
	etapta deneme sürümünü bi indirin, kendiniz görün neler yapılabildiğini.</p>

	</div>
</asp:Content>
