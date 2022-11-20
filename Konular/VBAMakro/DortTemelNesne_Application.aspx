<%@ Page Title='DortTemelNesne Application' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Dört Temel Nesne'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>

<h1>Nesnelerin Efendisi - Application</h1>
<h2 class="baslik">Giriş</h2>
<div class="konu">
	<p>Önceki bölümlerden hatırlayacağınız üzere, Excel, nesneler hiyerarşisi 
	üzerine kurulmuş bir <strong>Nesne Modeline</strong> sahiptir ve işte şimdi göreceğimiz
	<span class="keywordler">Application</span> nesnesi 
	de hiyerarşinin en tepesinde bulunur.</p>
	
	<p><strong>Application</strong> nesnesi Excelin ta kendisidir ve bu yüzden de default 
	nesnedir. Bu şu demek, bazı durumlarda bu ifadeyi 
	yazmanıza gerek olmadan buna ait özellik ve metodları kullanabiliriz. Yani 
	<strong>Application.ActiveWorkbook</strong> yazmak ile <strong>ActiveWorkbook</strong> yazmak arasında hiçbir 
	fark yoktur. Ancak bazı durumlarda da Application'ı açıkça yazmak gerekir. 
	Aslında bu konu <strong>Global</strong> sınıfı ile ilgili bir konu olup önbilgi adına 
	<a href="Giris_ExcelNesneModeli.aspx#global">buraya</a> bakabilirsiniz. 
	Eğer Excelin kendisiyle ilgili bir işlem olacaksa o zaman Application'ı açıkça 
	yazmamız gerekir. Mesela Excelden çıkış için <strong>Quit</strong> metodunu kullanmak, 
	Excelin ekrandaki boyutlarını ayarlamak gibi. Bununla beraber benim tavsiyem, 
	Application'ı her durumda yazmanızdır, ancak olur da internette araştırma 
	yaparken Application'ın yazılmadığını görürseniz de şaşırmayın. Aşağıda 
	özellik ve metodların tanıtımında Application'ın yazılması gereken durumlar 
	için bunu açıkça yazdım, diğer durumlar için yazmadım, ancak kod 
	örneklerinde size verdiğim tavsiyeyi tuttm ve Applicationu hep yazmaya 
	çalıştım.</p>

        <p>Bu arada herhangi bir nesnenin <span class=" keywordler">Application</span> özelliğini kullanarak da bu 
	Application nesnesini elde edebiliriz. Daha teknik bir ifadeyle, bazı 
	nesnelerde bulunan Application özelliği(propertysi) Application tipinde 
	değer döndürür, yani Application nesnesi elde edersiniz. Tabi bazen kullandığımız nesne Excel 
	Nesne Modeline ait bir nesne değil de mesela bir Outlook nesnesi olabilir, 
	böyle bir durumda Application propertysi kullandığınızda dönen değer 
	tabiki Excel değil, Outlook olacaktır. Bunları
	<a href="DigerUygulamalarlailetisim_Konular.aspx">Diğer Ofis uygulamlarıyla 
	çalışmak </a>bölümünde göreceğiz.</p>
	</div>

	<h2 class="baslik">Genel Görünüm ve Uygulama Seviyesi İşlemleri</h2>
	<div class="konu">
	
	<p>Excel'in genelini ilgilendiren birçok üye mevcuttur. Bunların birçoğu 
	<strong>File&gt;Options</strong>'tan ulaşabileceğiniz ayarların VBA karşılıklarıdır. 
	Önemlilerine, daha doğrusu kısa ve orta vadede kullanma ihtimaliniz olan 
	üyelere bir bakalım.</p>
	
		<p><span class="keywordler">Application.DisplayFullScreen özelliği</span>: 
		Boolean tipinde değer döndürür.True atandığında Ribbon olsun durum 
		çubuğu olsun hiçbirşey göstermez, sadece hücreler ve formül çubuğu 
		görünür.</p>
	<p><span class="keywordler">Application.DisplayFormulaBar özelliği:</span> Bu da Boolean 
	tiplidir. False atanırsa formül çubuğu gösterilmez. Bunu bazen formülleri 
	göstermemek için(protection yaparak da sağlanır) bazen de Displayfullscreen 
	özelliği ile birlikte, ekranda maksimum alanda yer açmak kullanılır.</p>
		<p><span class="keywordler">Application.DisplayScrollbars özelliği</span>:Bu 
		da Boolean tiplidir, scrollbarları 
		gösterir veya gizler.</p>
		<p><span class="keywordler">Application.Interactive özelliği</span>: 
		Boolean döndürür. Diyelim ki çok fazla copy-paste yapan bir makronuz 
		var, kodunuz da uzun sürüyor, beklerken o sırada Word veya Outlook'ta 
		vakit geçireyim dediniz. Outlookta yazdığınız bir metni kesip başka bir 
		yere kopyalamaya karar verdiniz, ancak tam az önce de Excel VBA kodunuz da bi 
		kesme işlemi yapmıştı, siz şimdi clipboarda Outlook metnini almış 
		oldunuz ve kodunuz hızlıca akıp geçti, ve paste işlemini yaparken 
		Excelden aldığı parçayı değil, Outlooktaki metni yapıştırdı. İşte böyle 
		bir durum olmasın diye bu tür işlemlerinizin olduğu kodlarınızın başına 
		bu özelliği yazıp False değerini atayabilir, kodun sonunda bunu yine 
		True'ya döndürebilirsiniz.</p>
		<pre class="brush:vb">Application.Interactive = False
'kodlar
Application.Interactive = True</pre>
		<p><span class="keywordler">Application.Quit metodu:</span>Excelden çıkış için 
	kullanılır. </p>
	<p>Bu arada Excelden çıkış yapılmasını yakalayacak bir event yok malesef. Bunu farklı 
        yöntemlerle tespit eden bazı makaleler gördüm ama oldukça karışık olduğu için buraya almak istemedim. 
        Ben şahsen, bunun yerine Personal.xlsb dosyasının kapanıp 
	kapanmadığını bu dosya içindeki Workbook_Beforeclose olayı ile yakalıyorum, bunun kapanması demek zaten 
	birçok durumda Excelin kapanması demek oluyor, ki bu da işimi görüyor. 
	Bunların detayını Olaylar(Events) bölümünde göreceğiz.</p>
		<p><span class="keywordler">Application.StatusBar</span>: Görev çubuğuna 
		mesaj yazmak için kullanılır. Özellikle kullanıcıları hem bilgilendirmek 
		hem de bilinçli bir şekilde mesaj kutusu çıkarmak istemediğinizde 
		faydalıdır. (Bazen schedule edilmiş işlerde arkadan gelen kodların 
		takılmasını engellemek için, bazen de kullancılardan gelen "Bu kadar 
		mesajbox çok can sıkıcı" itirazlarını ele almak için). Ancak 
		unutulmamalıdır ki, MsgBox kadar da dikkat çekici değildir, hatta bazı 
		durumlarda verdiğiniz mesaj gözden kaçabilir de. O yüzden kritik 
		mesajları MsgBox ile vermenizi tavsiye ederim. <a href="#Doevents">Aşağıda</a> bu özelliği ProgressBar olarak nasıl 
		kullanıyoruz, onu da göreceğiz.</p>
		<h4>Son Söz</h4>
		<p><strong>Object Browser</strong> veya MSDN üzerinden Uygulama seviyesinde yapılabilecek 
		daha bir çok ayarlama olduğunu görebilirsiniz, buraya önemli olduğunu 
		düşündüklerimi aldım, diğerlerini siz de araştırabilrsiniz. Bunların 
		çoğu Excel Options üzerinden yapacağınız ayarlamalara denk gelir. Ör: 
		Autorecover ayarı, autocorrect ayarı, dosyalar açıldığında linkleri 
		update etsin mi ayarı gibi.&nbsp;</p>
		</div>
		
		<h2 class="baslik">Kod hızlandırıcılar</h2>
		<div class="konu">
		
		<p>Aşağıdaki 5 özellik kodlarınızın başında False, sonunda True olarak 
		ayarlandığında performans kazanımı sağlar.</p>
		<p><span class="keywordler">Application.ScreenUpdating özelliği:</span>Uzunca bir 
		makro çalışırken ekranın bi gidip geldiğini, titrediğini görmüşsünüzdür(veya göreceksinizdir), 
		hele hele farklı workbooklar arasında gidip gelme sözkonusu ise bu durum 
		çok daha göze çarpar. Aslında tüm bu ekran hareketleri, genel süreci uzatan bir 
		rol oynar<span>, zira işlemciniz o sırada ekranı güncellemekle de 
		ilgilenmektedir. O</span> yüzden bu ekran hareketini kapatarak kodunuzu 
		hızlandırabilirsiniz. Bunu da bu özelliğe False değeri atayarak 
		yapıyoruz. Kod bitmeden hemen önce açmayı unutmayın tabi. </p>
		<p>Kod çok uzun sürüyorsa ScreenUpdating=False durumunda kullanıcılar 
		Excelin kitlendiğini düşünebilir, o yüzden arada bir hareket göstermek 
		iyi olabilir. Bunu da <a href="#Doevents">Doevents</a> metodu 
		ile yapabiliriz.</p>
		<pre class="brush:vb">'Genellikle bir döngü içinde mantıklıdır
Application.ScreenUpdating=False
Do Until oldumu = True
'ara kodlar
DoEvents 'burda ekran tazelenir
Loop</pre>
		<p><span class="keywordler">Application.DisplayAlerts özelliği:</span> Kodunuz 
	çalışırken Excel bize bazı uyarılar çıkarabilir, bunlar da genelde can sıkıntısı 
	yaratabilir. Özellikle schedule edilmiş makrolarınız varsa ve bunların 
	birinde bir sayfa silme, veya varolan dosya üzerine yazma gibi size uyarı 
	çıkaran kodlar varsa, bu özelliğe False atamazsanız ekran ilk uyarıda 
	takılı kalır ve sizin bir cevap vermenizi bekler. Eğer bir seri schedule 
	edilmiş kodunuz varsa, böyle bir durum kabul edilemez. O yüzden en faydalı 
	bulduğum özelliklerden biri budur. Bu özelliğe false atandığında uyarılara 
	varsayılan cevap verilir ve kod devam eder. <strong>Kod bittiğinde de bu özelliğe 
	otomatik True değeri atanır.</strong></p>
		<p>Ancak bu özellik iki durumda işe yaramaz. </p>
		<ul>
			<li>Mesaj kutularında. Özellikle schedule programınız varsa içinde 
			mesaj kutusu kullanmamaya çalışın. Hata yönetimi işlemlerinde bile 
			kullanmayın, onun yerine kendinize mail gönderebilir, 
			StatusBarı kullanabilir veya
			<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx#logger">Log kaydı 
			tutabilirsiniz.</a></li>
			<li>İçinde başka dosyalara link olan dosyalar açıldığında Linkleri 
			update edeyim mi sorusu. Burda DisplayAlerts yerine <strong>Workbook.Open
			</strong>metodunun <strong>UpdateLinks </strong>parametresi kullanılır.
			<a href="DortTemelNesne_Workbook.aspx">Buradan</a> bakabilirsinz. 
			Gerçi bunu da aşmanın bir yolu var, ama dikkatli kullanılmasında 
			fayda var.<span class="keywordler">Application.AskToUpdateLinks</span> özelliğine False atanırsa 
			linkler otomatik güncellenir, ve bu soru karşımıza çıkmaz, böylece her 
			Workbook.Open metodunda tek tek Updatelinks özelliğine değer girmek zorunda 
			kalmayız. Ama bazı durumlarda otomatik update olmasın isterseniz bunu kullanmak yerine
			Workbook.open metodunun <strong>Updatelinks</strong> parametresini kullanın.</li>
		</ul>
			<p>NOT: Kodunuzda bir uyarı mesajı çıkma durumu yoksa bu özellik 
			ekstra bir hızlanma sağlamayacaktır.</p>
	<p><span class="keywordler">Application.EnableEvents özelliği:</span> DisplayAlerts 
	ile birlikte en çok değer verdiğim bir diğer özellik de budur. Hatta 
	QuickAccess barda bu özelliği True ise False, False ise True yapan bir 
	düğmem bile var. Bunu&nbsp; biraz sonra açıklayacağım. Öncelikle ne işe 
	yarar ona bakalım. </p>
		<p>Bu özellik, herhangi bir event(olay) tetiklenmesin diye kullanılır. 
		İki tür kullanım şekli olabilir. </p>
		<ul>
			<li>Bir programın en başına false, en sonunua tekrar true olacak şekilde. 
			Böylece tüm kod boyunca hiçbir olay tetkilenmez.</li>
			<li>Bir döngü içinde satır silme, hücre değeri değiştirme gibi bir işlem vardır, ve 
			sayfa modüllerin birinde Worksheet_change eventiniz de vardır, sadece bu event 
			tetiklenmesin diye ilgili döngünün başına ve sonuna konur. Böylece döngüden 
		çıkıldığında diğer eventlerin tetiklenmesine imkan verilmiş olur.</li>
		</ul>
<p>
<span><span class="keywordler">Application.DisplayStatusBar özelliği</span>:Bu 
da Boolean tiplidir, En alttaki durum çubuğunu gösterir veya gizler. 
(Application.StatusBar özelliği ile karıştırılmamalıdır, bu ikincisinde durum 
çubuğunda yazan metni alırız veya metin yazarız.). Excel, her işlem sırasında 
StatusBar'ı güncellemekle uğraşmayacağı için performansa katkısı olacaktır.</span></p>
			<p>
			<span class="keywordler">ActiveSheet.DisplayPageBreaks 
			özelliği</span>:Bu özellik herne kadar Application nesnesine ait 
			olmasa da bağlam olarak buraya daha uygun olduğu için buraya aldım. 
			Kodunuz çalışırken Excel page berakleri tekrar tekrar hesaplamak 
			durumunda kalabilir. Bunu kapatarak(False atayarak) performansı 
			iyileştirebilirsiniz. </p>
			<p>
			<span class="keywordler">Application.Calculation özelliği:</span>Bu özelliğe aşağıda detaylı değineceğim.</p>
			<hr>
		<p>Ben çok faydalı bulduğum bu 5 özelliği bir prosedüre bağladım, ve birçok 
		makroya girerken bunlara(bazen sadece ikisine) false değerini atıyorum, koddan çıkarken de 
		tekrar true değerine döndürüyorum. Fonksiyon ve kullanım şekli 
		aşağıdaki gibidir:</p>
<pre class="brush:vb">
'Ana prosedürü
Public Sub AlertUpdatingEvent(a As Boolean, u As Boolean, p As Boolean, c As Boolean, Optional e As Boolean = True)
With Application
  .DisplayAlerts = a
  .ScreenUpdating = u
  If c = False Then
    .Calculation = xlCalculationManual
  Else
    .Calculation = xlCalculationAutomatic
  End If
  .EnableEvents = e
End With
ActiveSheet.DisplayPageBreaks = p
End Sub

'Prosedürü çağırma şeklim
Sub hızlandırıcılar()

On Error GoTo hata
AlertUpdatingEvent False, False, False, False 'Son parametreyi eklemediğimi için default değeri olan True atanır. Böylece eventler çalışmaya devam et sin ama diğer özellikler kapansın demiş oldum
'çeşitli kodlar
'....

AlertUpdatingEvent True, True, True, True 'eski hallerien getirdim
Exit Sub

hata:
AlertUpdatingEvent True, True, True, True 'eski hallerien getirdim

End Sub</pre>
			<p>
			Calculation özelliği True/False değeri almadığı için IF kontrolü ile 
			bunu çözdük. Bir diğer husus da henüz görmediğimiz hata yakalama 
			bloklarını kullanmış olduk. Olur da bir şekilde kodumuz tam bitmeden 
			ortada bi yerde patlarsa bu False atadığımız tüm özellikleri hata 
			bloğunda tekrar True'ya döndürmüş oluyoruz.</p>

</div>

	<h2 class="baslik">Hesaplama, Zamanlama(Scheduling) ve Bekle(t)me</h2>
	<div class="konu">

	<p>Bu başlıktaki konular her zaman olmamakla birlikte genelde birarada 
	kullanılmaktadırlar, en azından bir kullanım yakınlığı vardır diyebiliriz.</p>
		<h3>Hesaplama işleri</h3>
		<p><span class="keywordler">Application.Calculation özelliği:</span> Excelin 
	formüller için hesaplama yöntemini seçmenizi sağlar. Excelin bu özelliğini 
	bildiğinizi varsayıyorum, bilmiyorsanız öncesinde mutlaka 
	<a href="../Excel/FormulasMenusuDiger_CalculationSecenekleri.aspx">buraya</a>,
	<a href="https://msdn.microsoft.com/en-us/library/office/bb687891.aspx">
	buraya</a> ve
	<a href="http://www.decisionmodels.com/calcsecrets.htm">buraya</a> bakın. 
	</p>
		<p>Bu özelliğin alabileceği 3 enumaration değeri var. <br><br><strong>xlCalculationAutomatic</strong> 
	:Varsayılan değer budur. Herhangi bir hücrede değişiklik olduğunda tüm 
	workbooklarda formüller yeniden hesaplanır. <br><strong>xlCalculationSemiautomatic</strong> 
	:Table'lar dışında herşey otomatik hesaplanır. <br><strong>xlCalculationManual</strong> 
	:Hesaplama işlemi kapalıdır. Kullanıcı hesaplama yapana kadar da öyle kalır.<br>
	<br>Özellikle büyük formüllü dosyalarda bir makro çalıştıracaksanız ve herhangi 
	olumsuz bir etkisi olmayacaksa öncesinde hesaplama kapatılıp makro bitmeden 
	hemen önce de tekrar açılabilir.</p>
<pre class="brush:vb">
Application.Calculation = xlCalculationManual
'burada diğer işler yapılır
Application.Calculation = xlCalculationAutomatic
</pre>
		<p><span class="keywordler">Calculate metodu</span>: Tüm workbooklardaki 
		yeni, değişmiş ve volatil formüllerin hesaplanmasını sağlar(o anda Manuel hesaplama seçimi 
		yapıldıysa anlamlıdır, aksi halde&nbsp;zaten formüller hesaplanmıştır ve 
		gerek yoktur). 
		Calculation property'sinin aksine bunda Application denmesine gerek 
		yoktur.&nbsp;Örnek biraz aşağıda bulunmaktadır.</p>
		<p><span class="keywordler">Application.CalculationState özelliği</span>: Hesaplamanın ne 
		durumda olduğunu gösterir. Alabileceği değerleri bir kodla görelim.</p>
<pre class="brush:vb">
Sub CalcState()
If Application.CalculationState = xlDone Then 'enumeration değeri 0
	MsgBox "Hesaplama Bitti"
ElseIf Application.CalculationState = xlPending Then 'Görev çubuğunda "Calculate" yazar 'enumeration değeri 2
	MsgBox "Tetiklendi ama henüz hesaplama başlamadı"
Else 'xlCalculating 'enumeration değeri 1
	MsgBox "Hesaplama devam ediyor 'Görev çubuğunda %sel bir oran görünür
End If
End Sub	 </pre>
		<p><strong>Done</strong> ve <strong>Calculating</strong> gayet aşikar fakat 
		<strong>Pendingi</strong> tam olarak anlamamış 
		olabilirsinz. Hani bazen Excel manuel hesaplama modundayken, 
		formüllerden birine baz teşkil eden bir hücreyi değiştirdiğinizde en 
		alttaki durum çubuğunda <strong>Calculate</strong> yazar, bazen de dosyanızda çok sayıda 
		formül varsa Excel bu kadar 
		formülle başa çıkamaz ve en altta yine Calculate yazar. İşte bu durum 
		xlPending durumudur. Böyle durumlarda Excelde hesaplama yapmak için Formulas 
		menüsünden Calculate demek veya F9'a basmak gerekir. <br><strong>NOT</strong>:Bir de kısırdöngülü 
		formüllerde Calculate yazdığını görürsünüz, bu da bir xlPending 
		durumudur ancak onun çözümü aşağıdakiler değil, kısırdöngüye neden olan 
		formülü düzeltmektir.</p>
		<p>İnternette birçok forumda yaygın bir kullanım örneği olarak aşağıdaki 
		kod parçası verilir. Deniyor ki, "kodunuz çalışmaya başlamıştır, 
		büyük bir hesaplama yapıyordur, ancak daha hesaplama bitmeden bir 
		sonraki satıra geçer, bu da hatalı sonuçlar neden olabilir, o yüzden 
		aşağıdaki kod ile kodununzun aşağı satıra geçmesini engellersiniz". </p>
<pre class="brush:vb">
Application.Calculate 'hesaplamaya başladınız
Do While Application.CalculationState <> xlDone
     DoEvents
Loop
'kodun kalan kısmı</pre>
		<p>
		Halbuki Calculate metodu asynchoronus değildir, yani hesaplama 
		bitmeden zaten bir sonraki satıra geçmez. O yüzden yukarıdaki tavsiye 
		bence anlamsızdır. Ancak bir şekilde(forumlarda yardım isteyen diğer 
		kişilerin başına gelen çok özel durumlarda, artık neyse o özel durumlar bilemiyorum) 
		böyle birşey olduğunu farkederseniz bu kodu kullanabilirsiniz.</p>
		<p>
		Bu arada şu farkı iyi anlamanız gerekiyor; Calculate metodunu 
		uyguladığınızda sanki Excelde F9'a basmış veya Calculation menüsünden 
		Calculate butonuna basmış gibi olursunuz ve formüller yeniden hesaplanır ancak sayfanız o an hala Manuel modda kalmaya devam eder ve 
                sonraki aşağı/sağa formül kaydırma işlemleri sonucunda formüller hesaplanmaz. Halbuki Calculation özelliğine xlAutomatic atayarak hem 
                hesaplamayı 
		açmış olurusunuz hem de statüyü kalıcı olarak Otomatiğe çevirmiş 
		olursunuz ve sonraki formül kaydırmalarda formüller hemen hesaplanır. 
		Hangisi ihtiyacınıza uygunsa onu kullanmalısınız. Eğer ki geçici bir hesaplama yapmak istiyorsanız Calculate metodunu, kalıcı hesaplama için ise Calculation özelliğini kullanabilirsiniz.</p>
		<p>
		Önemli bir husus da şudur; Application, Worksheet ve Range nesnleri için varolan Calculate metodu 
		Workbook için bulunmamaktadır. Ancak aşağıda gibi bir kod ile sadece Activeworkbook'un 
		Calculation işlemini yapabilrsinz.</p>
		<pre class="brush:vb">
Sub CalcBook()
    Dim ws As Worksheet
    Application.Calculation = xlManual
    For Each ws In ActiveWorkbook.Worksheets
        ws.Calculate
    Next
    Set ws = Nothing
End Sub</pre>
		<p><span class="keywordler">Application.CalculateFull</span>: Otomatik veya Manuel 
		modda olun farketmez, tüm formüllü hücreleri yeniden hesaplar. Calculate 
		metodundan farklı olarak, sadece yeni, değişmiş ve volatil formülleri 
		değil, tüm formül içeren hücreleri tekrar hesaplar. Bu yüzden 
		genelde(her zaman değil) normal Calculate metoduna göre daha yavaştır. Durum 
		çubuğunda ısrarla Calculate yazıyorsa yani xlPending durumundan bir 
		türlü çıkamıyorsanız bunu kullanabilrsiniz. 
		Klaveye kısayolu <strong>Ctrl+Alt+F9</strong>'dur.</p>
		<p><span><span class="keywordler">Application.CalculateFullRebuild:</span>Bu metod 
		CalculateFull ile aynı işi yapıyor gibi görünüyor, Excel 2007 ve sonrası 
		kulanıcıların çok kullanacağı bir metod değildir. Özetle şunu diyebilirim 
		ki, 2007 öncesi versiyonlarda aşırı formülden dolayı hesaplama zinciri 
		bozulduysa ve F9 yaptığınız halde Excel hesaplama yapmıyorsa bu metod işe 
		yarayacaktır. Ancak sanki tüm hücrelere formülleri 
		tekrar girmek gibi iş yaptığı için CalculateFull'e göre biraz daha yavaştır.</span></p>
		<p><span><span class="keywordler">Application.CalculationInterruptKey</span>:Hesaplamanın 
		hangi tuşla iptal edileceğini söyler. Bunun pratik kullanımı, 
		Personal.xlsb'nin Workbook_Open makrosu içine yazma şeklindedir. Ben 
		şahsen sadece ESC tuşuna(xlEscKey) basıldığında hesaplamanın iptal 
		edilmesini istiyorum, size de bunu öneririm. Zira eliniz yanlışlıkla bi 
		ok tuşuna değse bile o anda %90larda olan calculation tekrar %0dan 
		başlayacaktır.</span></p>
<h4>Genel öneriler</h4>
		<ul>
			<li>Önce normal Calculation yapın. Sonra state kontrol edin, hala 
			xlPendingse CalculateFull uygulayın.</li>
			<li>Kod hızlandırıcılar bölümündeki 3 özelliğe bazen bu Calculation'ı 
			da ekleyerek daha hızlı kod çalıştırabilirsiniz. Ancak kullanımı konusunda 
			dikkatli olmak gerekir, zira arada bir yerlerde formül çekme/uzatma ve sonra 
			Copy-Paste işlemi varsa Calculation sonucunda hatalı durumlar 
			oluşabilir.</li>
		</ul>
		<pre class="brush:vb">Sub Calculationlar()

AlertUpdatingEvent False, False, False
Application.Calculation = xlCalculationManual
'kodlar buraya gelir

'arada bir açmak gerekebilir
ActiveSheet.Calculate 'duruma göre Application.Calculate veya Range("...").Calculate
If Application.CalculationState=xlPending then Application.CalculateFull
'tekrar kapatalım
Application.Calculation = xlCalculationManual
'diğer kodlar

'çıkışta tekrar eski haline getiriyoruz
AlertUpdatingEvent True, True, True
Application.Calculation = xlCalculationAutomatic

End Sub
</pre>
		<p>
		Son olarak Calculate işlemlerinin VBA ve Excel ilişkilerini tekrar şöyle 
		bir özetleyelim:&nbsp;</p>
		<p>
		<table class="alterantelitable">
				
			<th>İşlem</th>
			<th>Excel</th>
			<th>VBA</th>
		
			<tr>
				<td>Tüm workbookları hesaplatmak </td>
				<td>Calculation&gt;Calculate(veya F9)</td>
				<td>Application.Calculate</td>
			</tr>
			<tr>
				<td>Aktif sayfayı hesaplatmak</td>
				<td>Calculation&gt;Calculate Sheet(Shift+F9)</td>
				<td>ActiveSheet.Calculate</td>
			</tr>
			<tr>
				<td>Aktiveworkbook hesaplatmak</td>
				<td>-</td>
				<td>Döngü içinde Sheet.Calculate</td>
			</tr>
			<tr>
				<td>Belli bir range'i hesaplatmak</td>
				<td>-</td>
				<td>Range.Calculate</td>
			</tr>
			<tr>
				<td>Full hesaplama yapmak</td>
				<td>Ctrl+Alt+F9</td>
				<td>Application.CalculateFull</td>
			</tr>
		</table>
</p>
		<h4>Calculation için örnek bir senaryo</h4>
		<p>Şimdi diyelim ki departmanınızdaki kişilerin kullanması için çok 
		sayfalı ve çok formüllü bir excel dosya hazırladınız. İlk sayfada tek 
		sayfalık bir karne/skorkart tarzı birşey var, diğer sayfalarda ise toplu 
		listeler. Hepsi de datayı gizli bir sayfadan alıyor.</p>
		<p>Liste sayfalarında çok fazla satır ve sütun ve hep SUMIFS tarzı 
		formüller olduğu için bunlarda sort veya filter işlemleri çok ağır 
		olmaktadır, zira bu iki işlem de calculation tetikleycisidir. Çözüm 
		şöyle olabilir:</p>
		<p>Dosyanın Workbook_Activate(Neden Workbook_Open olmadığını az sonra 
		belirteceğim) eventine dosya açılır açılmaz Calculation'ı xlManual yapan 
		kodu ekledim, ve bi mesajbox ile bunu kullanıcıya bildiriyorum.(Mesajbox 
		sinir bozucu gelirse statusbara da yazdırabilirsiniz). Karne sayfasına 
		gelince ise Worksheet_Change eventine yazdığm kod ile 
		sadece belli hücreler değiştiğinde hesaplama yapmasını sağlıyorum. Kullanıcı olur da o sırada başka 
		dosyalarda işlem yapmak isterse Workbook_Deactivate eventine 
		calculationı tekrar otomatik yapan bir kod yazdım ki, kullanıcı o sırada 
		Calculationın kapatıldığını unutup diğer dosyalarda formül uzatma gibi işler yaparsa hep aynı 
		sonucun yazdığını görüp 
		şaşırmasın. Hatta ortalama bir kullanıcı Excelin Calculation özelliğinden 
		bihaberdardır bile diyebiliriz. Kodlar şöyle:</p>
		<pre class="brush:vb">
Private Sub Workbook_Activate()
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Dosya aktive olduğu için Calculation yine geçici olarak Manuel yapıldı"
End Sub
 
Private Sub Workbook_Deactivate()
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = "Başka dosyayı açtığınız için Calculation tekrar otomatik yapıldı"
End Sub
		</pre>
		<p><span class="keywordler">Application.Volatile metodu</span>:Bu metoda 
		UDF bölümünde değindiğimiz için burada ayrıca bahsetmiyorum.</p>

		<h3>Erteleme ve bekleme</h3>
		<p><span class="keywordler">Application.OnTime metodu</span>: Site boyunca 
		zaman zaman programlanmış işlerden veya İngilizce tabiri ile işleri schedule 
		etmekten bahsediyorum, 
		mutlaka dikkatinizi çekmiştir. İşte bu işi bu harika metod ile 
		başarıyorum. Genel kullanım şeklini aşağıda veriyorum ama bu konuyla 
		ilgili uzunca bir örneğe 
		<a href="DigerUygulamalarlailetisim_VeriTabani-ConnectionListObjectveQueryTable.aspx#hayaletprotokol">
		şu sayfada </a>ele aldım.</p>
		<p>Şimdi metodun genel syntaxına bakmadan önce görevini açıkça 
		belirtelim: Bir prosedürün belirli bir anda çalışmasını sağlar, ki bu 
		tanım bize onu asıl amacı dışında(ama faydamıza olacka şekilde) kullanacağımızı da söylemektedir, yani 		<span class="keywordler">Wait </span>ve <span class="keywordler">Sleep</span> 
		metodları yerine. Bunlara da hemen bu metoddan sonra değineceğiz.</p>
		<p><strong>Syntax</strong>: <strong>ApplicationObject.OnTime(EarliestTime, ProcedureName, LatestTime, 
		Schedule</strong></p>
		<p>Tam açıklaması şöyle oluyor. ProcedureName ismindeki makro 
		EarlistTimeda başlasın, o sırada başka bir makro çalışıyorsa veya Exceli 
		meşgul eden başka birşey varsa da LatestTime'a kadar çalıştırmayı 
		denesin. Eğer LatestTime belirtilmezse, excelin meşguliyeti bitene kadar 
		bekler ve sonra çalıştırır. Yani eğer, "Kod,programladığım saatten en geç 
		1 saat içinde çalışsın, yoksa çalışmasının bi anlamı yok, çünkü o rapor 
		artık işe yaramaz olur" dediğiniz bir durum varsa bu parametreyi 
		"EarliestTime + 1 saat" olarak belirtebilirsiniz, aksi durumda boş bırakın. 
		Schedule parametresi default değeri True'dur ve genelde yazılmaz, schedule 
		ettiğiniz bir prosedürü iptal etmek için bu değere <strong>False</strong> atarsınız.</p>
		<p><strong>Önemli Not</strong>:Excelden her çıkış yaptığınızda, tüm schedule programı 
		sonlanır. Eğer, recursive(tekrarlı) yani bittikten sonra yeniden 
		schedule edilen programınız varsa Exceli hep açık bırakmanız gerekir, ki 
		benim bilgisayarımda olan budur. Ve bence bu metoddan verim almanın en 
		güzel yolu onu recursive bir şekilde kullanmaktır. Şimdi küçük bir örnek 
		bakalım, siz sonra yukarda linkini verdiğim yerden daha detaylı örneği 
		incelersiniz.</p>
<pre class="brush:vb">Sub ontimeornek()
    Application.OnTime Now + TimeSerial(0, 0, 3), "mesajver" 'Asynhronous metoddur
    MsgBox "beklemeden çalıştım"
End Sub

Sub mesajver()
    MsgBox "selam"
End Sub</pre>
		<p>Örnekte gördüğünüz üzere mesajver makrosunun çalıştırılacağı zamanı 
		Run tuşuna bastıktan 3 sn sonra çalıştıracak şekilde parametrik verdim. 
		Yani burada spesifik bir saat belirtmek yerine, şimdiye(Now) referansla 
		bir saat de verebiliyoruz. Örnek gösterimler şöyle olabilir.</p>
<pre class="brush:vb">
Application.OnTime "22:30:00"
Application.OnTime Now + TimeValue("00:10:00") 'Şimdiden 10 dk sonra
Application.OnTime Now + TimeSerial(0,10,0) 'Bu da aynı. TimeSerial'de saat, dakika ve saniye virgülle ayrılır</pre>
		<p><span class="keywordler">Application.Wait metodu</span>: Programın belirli 
		bir süre durmasını(beklemesini) sağlar. Peki neden? Neden programınızın 
		bir süre durmasını bekleyesiniz ki? İşte örnek senaryolar olmayınca 
		malesef makro öğrenimi çok zor olmaktadır. Önce gelin nasıl 
		kullanılacağına , sonra nedenine bakalım.</p>
<pre class="brush:vb">
Sub bekle()
   Application.Wait (Now + TimeValue("0:00:10")) 'Synhronous metoddur
   MsgBox "bekleyip çalıştım"
   Call mesajver
End Sub

Sub mesajver()
    MsgBox "selam"
End Sub</pre>
		<p>Bu metod Boolean döndürdüğü için belirli bir zaman geçip geçmediğini 
		kontrol etmek için de kullanılır.</p>
		<pre class="brush:vb">If Application.Wait(Now + TimeValue("0:00:10")) Then '10 sn geçtiyse. =True demeye gerek görmeyebiliyoruz, önceki konuları hatırlayacak olursanız
 Application.Speech.Speak "Zaman doldu" 'Evet, Excel 2013ten itibaren artık konuşuyor 
End If</pre>
		<p><span class="dikkat">Dikkat:</span> Bu metod kullanılırken çok dikkat 
		etmek gerekir, zira ilgili süre geçene kadar Excel kitlenir.</p>
		<p>Şimdi de örnek bir senaryo düşünelim. Diyelim ki bir makronuzu 
		sabah/gece 5'e 
		schedule ettiniz: Kodunuz bir veritabanını güncel veri gelmiş mi diye 10 dk'da bir tarıyor, ve sonunda 5:40ta yani 4. seferinde güncel datayı gördü ve 
		hemen çekti, bikaç işlem yaptı, 5.43te işi bitti ama kod devam ediyor, 
		ettiği yerde başka bir veritabanı bağlantısı yapacak, ama siz biliyorsunuz 
		ki o veritabanı 6:30da doluyor, 5.43te buraya bağlanmaya çalışırsa 
		güncel olmayan veriyi alabilir, işte böyle bir durumda kodu 6:30a kadar 
		bekletmek gerekebilir.</p>
		<pre class="brush:vb">Application.Wait "06:30:00"</pre>
		<p>Hatta bunu bir de saat 6:30dan önce&nbsp; mi diye kontrol etmek 
		lazım, eğer ilk veritabanını sorgulanması 5:40 değil de 6:30dan sonraya 
		kaldıysa ikinci veritabanı için beklemelik bir durum olmayacaktır, o 
		yüzden Wait kullan<span style="text-decoration: underline"><strong>mama</strong></span>k gerekir, aksi halde ertesi sabah 6:30'a kadar 
		Exceliniz bloke olur.</p>
		<p>Bir diğer örnek durum da şu olabilir. Veritabanı işlemlerinde 
		göreceğiz gerçi ama <strong>RefreshAll</strong> gibi <strong>asyncrohnous</strong>(sonraki 
		satıra geçmek için beklemeyen) bir metod çalıştığında 
		kod okuma devam eder. Eğer tüm refresh işleminin bitmesini beklesin 
		istiyor ve tahminen refreshin ne kadar süreceğini biliyorsanız 
		ilgili süre kadar bekletebilirsiniz.</p>
		<pre class="brush:vb">'Önceki kodlar
Me.RefreshAll
Application.Wait Now + "00:30:00" '30 dakika yeterlidir
'diğer kodlar</pre>
		<p>Yine bu
		<a href="https://www.experts-exchange.com/questions/25861576/How-do-I-pause-wait-1-second-in-Excel-VBA.html">
		linkte</a> sorulan soruya da güzel bir yanıt verilmiş. Soru şu: "Bir 
		veritabanından 3000 farklı kayıt okumaya çalşıyorum, ancak bazen kod o 
		kadar hızlı akıyor ki, bazı kayıtları okuyup hedef dosyaya yazdıramıyorum. 
		Excelin her kayıt için yeterince beklemesini nasıl sağlarım." Buna verilen 
		cevap oldukça güzel, gerçi bunda Wait metodu yerine hemen bir alttaki 
		Sleep fonksiyonu kullanılmış. Ayrıca DoEvents bilgisi de gerekiyor, ki o da 
		Sleepten 
		hemen sonraki konu.</p>
		<p><span class="keywordler">Sleep Metodu:</span> Bu metod VBA metodu 
		olmayıp Windows fonksiyonudur, bu yüzden bunu kodumuzun başına import 
		etmek gerekir. Wait ile aynı işi yapar, tek farkı, milisaniye cinsinden 
		paremetre almasıdır. İmport işlemi dahil tam bir kod örneği aşağıda 
		bulunmaktadır. (import işlemleri hakkında detaylı bilgi
		<a href="https://msdn.microsoft.com/en-us/library/office/gg278581.aspx">
		burada</a> mevcuttur)</p>
		<pre class="brush:vb">'bu kısım import kısmıdır, kod sayfasının en tepesine yazılır
#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) '64 Bit Sistemler için
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '32 Bit Sistemler için 
#End If

'bu kısım esas kod kısmı
Sub sleepmetodu()
    MsgBox "Başlıyorum"
Sleep 10000 '10bin mili saniye yani 10 sn bekliyoruz
    MsgBox "Süre bitti"
End Sub</pre>
		<p>Yukarda belirttiğimiz gibi, <span class="keywordler">OnTime</span> metodunu da amacı dışında bir nevi 
		bekleme amacı olarak kullanabiliriz, ancak Wait ve Sleepte kod akışı o 
		satırda duruken OnTime'da durmaz devam eder, o yüzden OnTime eğer bu 
		amaçla kullanılacaksa ilgili prosedürün son satırı olmasında fayda var.</p>
		<pre class="brush:vb">Sub tetikleyici()
'çeşitli kodlar
Application.OnTime Now + TimeSerial(1,0,0), "makrom"
End Sub

'şimdi bu durumda 1 saat sonra makrom makrosu başlayacak. Ekran 1 saat boyunca serbest.
Sub makrom()
'çeşitli kodlar
End Sub</pre>
		<p>Bunların dışında bir de farklı bir yaklaşım var ki, bunda da 
		OnTime'da olduğu gibi bekleme süresince Excele erişimimiz açık 
		durmaktadır, zira bunda bir döngü içinde <span class="keywordler">DoEvents</span> 
 
		metodu kullanılmaktadır.</p>
		<pre class="brush:vb">Dim newTime As Date
newTime = Now + TimeValue("00:00:10")
Do While Not Now &gt;= newTime
    DoEvents
Loop
MsgBox "selam</pre>
		<p><span class="keywordler"><a name="Doevents"></a>DoEvents:</span> MSDN açıklaması 
		çok yüzeysel malesef, ben yine de bu açıklamayı vereceğim, sonra nerelerde kullanıldığını 
		söyler ve birkaç örnek gösterirsem daha anlaşılır olur.</p>
		<p><strong>MSDN açıklaması: </strong>Program akışını işletim sistemine verir.</p>
		<p>Şimdi MSDN'nin açıklamasını biraz 
		daha açalım. Bir Windows uygulamasında aynı anda onlarca program 
		çalışabilmektedir. Sizin Excel VBA kodunuz işlemciyi çok fazla bloke 
		ederse Windows buna sinirlenebilir, hatta bakar ki Excelden ses seda 
		yok, bunun kitlendiğini düşünüp kapatmaya çalışabilir, çünkü belleksiz 
		kalan diğer gariban programlar çalışsın diye 
		düşünür. İşte böyle durumlarda araya bi <strong>DoEvents</strong> sokmak gerekebilir, 
		ki Windows Excel'in yaşadığını düşünsün. Peki genel olarak nerelerde kullanırız ona bakalım:</p>
		<ul>
			<li>Büyük döngülerde araya girebilmek için(özellikle kitlenme 
			durumları yaşıyorsanız)</li>
			<li>ScreenUpdating=False olduğunda kullanıcı Excelin kilitlendiğini 
			sanmasın diye ekran tazelemede</li>
			<li>Bir şartın olmasını beklerkenen<ul><br>
<pre class="brush:vb">
Do Until beklenen=true
   DoEvents
Loop</pre>
			</ul>
			</li>
		</ul>
		<p>Bu yukardaki ilk iki maddeyi içeren bir örnek verelim. Uzun bir 
		döngüsel işleminiz var diyelim, hızlı çalışsın diye ScreenUpdating=False 
		yaptınız. Kullanıcı program kitlendi sanmasın diye, <strong>DoEvents</strong> metodu 
		Statusbarı bir nevi <strong>Progressbar</strong> gibi kullanmamızı sağlayacak. Tabi bu aşağıdaki 
		örnekte %1den %100e giden bir progressbar yapmış olduk ancak siz 
		isterseniz başka ölçüler kullanabilrisiniz. Mesela 20 bölgelli bir 
		bankada her bölgenin işi 1 dk sürüyorsa, her döngü sonuna "20 bölgede " 
		&amp; i &amp; " adedinde işlem tamam" gibi bir metin yazdırabilirsiniz.&nbsp; </p>
		<pre class="brush:vb">
Sub doeventprogressbar()
	Dim i As Long
	Dim bas As Double
	bas = Timer
	Application.ScreenUpdating = False
	For i = 1 To 100000 '100.000in 100lük dilimlere böldüğümüzde her bir bölümün büyüklüğü
		Cells(1, 1) = i
		If i Mod 1000 = 0 Then
			DoEvents
			Application.StatusBar = "%" &amp; i * 100 / 100000
		End If
	Next i
	Application.ScreenUpdating = True
	MsgBox Round(Timer - bas) &amp; " sn sürdü"
End Sub
</pre>
		<p>Bu arada bu koddan screenupdatingi çıkarın, programın ne kadar 
		yavaşladığını göreceksiniz.</p>
	</div>
	
	
		<h2 class="baslik"><a name="filefolder"></a>Dosya ve Klasör işlemleri</h2>
<div class="konu">	
	<p>Bazen kullanıcıdan, üzerinde işlem yapılacak bir dosya veya klasör 
		seçmesini isteriz. Bazı durumlarda seçilen dosya ile sadece işlem 
		yapılırken bazen dosyanın açılması sağlanır. </p>
		<p>Bu bölümde anlatılan konular genel olarak, seçilen dosyayı açma veya 
		bir şekilde dosya/klasör ismi elde etme amacıyla kullanılan işlemlerle 
		ilgilidir. Daha genel olarak tüm dosya işlemlerini
		<a href="Dosyaislemleri_Konular.aspx">şurada</a> ele alıyor olacağız.</p>
		<p>Kullanıcıdan dosya/klasör bilgisi istemenin en ilkel yolu bunu bir 
		inputboxla sormak olacaktır, ancak şükür ki VBA'de bunu yapmamızı 
		sağlayan daha iyi yöntemler var. Şimdi bunlara bakalım:</p>
		<p><span class="keywordler">Application.FileDialog özelliği</span>: 2002 
		yılında gelen bu özellik bundan daha önce varolan <strong>GetOpenFilename</strong> ve 
		<strong>GetSaveAsFileName</strong> özelliklerinin gelişmiş halidir. O yüzden bu ikisinin 
		artık çok kullanmaya gerek yok, ama başka kodlarda görmeniz durumunda 
		ne olduğunu bilmeniz için onlara da kısaca değineceğiz. Bu 
		özelliğin FileDialogType şeklinde tek bir parametresi vardır ve o da <strong>MsoFileDialogType </strong> 
		sabitlerinden biri olabilir. Bunlar:</p>
		<ul>
			<li>
				<strong>msoFileDialogFilePicker: </strong>Dosya seçtirir, path 
				dahil tam 
				ismini döndürür(Ör:"C:\Hedefler\Satış\2016.xlsx"</li>
			<li><strong>msoFileDialogFolderPicker: </strong>Klasör seçtirir, 
			path dahil tam 
			ismini döndürür(Ör:"C:\Hedefler\Satış"</li>
			<li><strong>msoFileDialogOpen: </strong>Açılacak dosyayı seçtirir(onu 
			açmaz, sadece seçtirir)</li>
			<li><strong>msoFileDialogSaveAs:</strong>Farklı Kaydet dialog 
			kutusunu açar, dosyayı kayderken ismin ne olacağını girmenizi sağlar(dosyayı 
			kaydetmez,sadece isim ve adres belirlersiniz) </li>
		</ul>
		<p>
		Bu property ile <strong>FileDialog</strong> nesnesi elde edilir. Bu nesnenin de kendi 
		metod ve özellilikler vardır. Genel olarak önce bu tipte bir değişken 
		yaratmak ve ona atama yapmak intellisense açısından uygun olacaktır.</p>
		<p>
		Bu nesnenin iki metodu var, pratikte en çok kullanacağınız metodu <span class="keywordler">
		Show</span> metodudur. Bir seçim yapıldıysa -1(True) döndürür, seçim 
		yapılmadan işlem iptal edilirse 0(False) döndürür.
		<a href="Temeller_Terminoloji.aspx#fonksiyonmetod">Bu sayfada </a>
		gördüğümüz gibi, bu metod arka planda bir function prosedür olarak 
		hazırlanmıştır, çünkü bize bir değer döndürüyor. </p>
		<p>
		Seçim sonucu True ise <span class="keywordler">Execute </span>metodunu 
		yazarak duruma göre uygun işlemi de yaptırabilrisiniz. Ancak bunun 
		yerine Workbook.Open veya Workbook.SaveAs gibi metodlar da 
		kullanılabilir. Execute yazmak basit görünse de arka planda seçim tipi 
		ve sonucunu karşılaştıran bir koşullu yapı barındırdığı için perfomans 
		sorununa neden olabilir, özellikle büyük kodlarda. O yüzden doğrudan 
		Workbook metodlarını kullanmanızı tavsiye ederim.</p>
		<p>
		Önemli propertyler ise şöyledir.</p>
		<ul>
			<li><strong>AllowMultiSelect</strong>: Çoklu seçim yaptırma imkanı verir, 
			sadece dosyalar için geçerlidir, klasörlerde çoklu seçim yapılamaz.</li>
			<li><strong>Title</strong>:Dialog kutusunun başlığını değiştirebilirsiniz.</li>
			<li><strong>InitialFileName</strong>:Seçim yaptırırken sık kullanılan bir dosya/klasör 
			varsa default olarak bunu seçtirebilirsiniz, SaveAs yaparken de yine 
			aynı mantıkla default bir adres ve isim belirleyebilirsiniz.</li>
			<li><strong>Filters</strong>:Hangi tür dosyaların gösterileceğini 
			belirlersiniz.</li>
			<li><strong>SelectedItems</strong>:Seçilen dosyaların/klasörlerin 
			tam adresini verir. Tek seçim yapıldıysa SelectedItems(1) şeklinde 
			kullanılır.</li>
		</ul>
		<p>Şimdi ilk olarak dosya seçme örnekleriyle başlayalım, sonra da 
		diğerlerine geçelim. </p>
		<p>Aşağıdaki örnekte dosya seçtiriyorum ve seçilen dosyaları siliyorum. Bu arada <strong>With .. End With</strong> yapısını nasıl kullandığımıa dikkat edin. Bilgi almak için <a href="Giris_ExcelNesneModeli.aspx#withend">buraya</a> tıklayın.</p>
<pre class="brush:vb">
Sub Dosyaislemlerim()

Dim i As Byte
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
  .AllowMultiSelect = True
  .Title = "Silinecek dosyaları seçin"
  .Show
  
   For i = 1 To .SelectedItems.Count
      Kill .SelectedItems(i) 'buradaki kill metodunu sonra göreceğiz, şimdilik sadece sdosya silmeye yaradığını bilin
   Next
End With

End Sub
</pre>
	<p>Dialog kutusunu seçim yapmadan kapadığımızda bunu anlayan bir kontrol noktası koyalım. Bu 
	kontrol noktasını doğrudan metodun sonuç değeri ile yapıyoruz. </p>
		<pre class="brush:vb">
Sub Dosyaislemlerim()

Dim i As Byte
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .AllowMultiSelect = True
    .Title = "Silinecek dosyaları seçin"
    If .Show = 0 Then
        MsgBox "Seçim yapmadan çıkış yaptınız"
        Exit Sub
    End If
               
    'iptal edilirse buraya gelmeden program sonlanır, çünkü Exit Sub denildi
    For i = 1 To .SelectedItems.Count
        Kill .SelectedItems(i)
    Next
End With
 
End Sub		</pre>
		<p>Şimdi de diğer 3 tip için de örnekler yapalım. Önce dosya aç:</p>
		<pre class="brush:vb">Sub Dosyaislemlerim()

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogOpen)

With fd
    '.AllowMultiSelect = True
    .Title = "Açılacak dosyayı seçin"
    .InitialFileName = "C:\deneme.xlsx"
     If .Show = True Then
        .Execute 'veya Workbooks.Open (.SelectedItems(1))
     End If
End With

End Sub</pre>
		<p>Şimdi default dosya adı belirtmeyelim, kullanıcı seçsin ama çok 
		kalabalık dosya türünü barındıran bir görüntü de olmasın, sadece excel 
		dosyaları olsun.</p>
		<pre class="brush:vb">Sub Dosyaislemlerim()

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogOpen)

With fd
   '.AllowMultiSelect = True
   .Title = "Açılacack dosyayı seçin"
   '.InitialFileName = "C:\inetpub\wwwroot\ee\yuklemeler\pivot - data.xlsx"
   .Filters.Clear 'varsayılan olakra 20 tane dosya tipi var, bunları temizlzeyelim
   .Filters.Add "Excel dosyaları", "*.xls*"
   .Filters.Add "Tüm dosyalar", "*.*" 
   If .Show = True Then
     .Execute 'veya Workbooks.Open (.SelectedItems(1))
    End If
End With

End Sub
</pre>

<p>Eğer farkettyiseniz msoFileDialogFilePicker ve msoFileDialogOpen ifadelerinin her ikisini de dosya açmada kullanabiliriz. Hatta dosya silme işlemi için bile msoFileDialogOpen kulanılabilir. Yukarıdaki msoFileDialogFilePicker ile yapılmış dosya silme örneğinde kendiniz deneyip görebilirsiniz. O halde neden iki ayrı ifadeye gerek var diye düşünüyor olabilirsiniz. Buna malesef benim de cevabım yok. O kadar araştırdım ancak bir açıklama bulamadım. Bulduğum zaman bu paragrafı 
güncellerim.</p>

<p>Şimdi de folder işlemlerine bakalım</p>
<pre class="brush:vb">Sub folderislemleri()
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)

With fd
   .AllowMultiSelect = True 'buna True dense bile çoklu seçim yaptırmaz
   .Title = "klasör sçein"
    If .Show = True Then
       MsgBox .SelectedItems(1) &amp; " klasörünü seçtniz"
    End If
End With
End Sub</pre>
<p>Son olarak da SaveAs işlemi yapalım.</p>
<pre class="brush:vb">
Sub Dosyaislemlerim()

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogSaveAs)

With fd
   .InitialFileName = "C:\deneme.xlsx"
   If .Show = True Then
     .Execute 'veya Workbooks.Saveas ile
   End If
End With

End Sub</pre>
		<p><span class="keywordler">Application.GetOpenFilename ve Application.GetSaveAsFilename</span>: 
		Yukarıda bahsettiğim gibi yeni kodlarınızda bu iki metod yerine 
		FileDialog yöntemini kullanmanızı tavsyie ederim. Bunları sadece geriye 
		dönük uyumluluk adına ve karşınıza bunları içeren bir kod 
		geldiğinide anlamanız için bilmeniz gerektiğini düşünüyorum.</p>

		<p><strong>Syntax: GetOpenFilename( GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)</strong></p>

		<p>Bunda FileDialogda olduğu gibi bir nesne yaratmaya gerek yok, 
		doğrudan kullanılabilir. Bundan dönen değer üç şey olabilr. Seçim 
		yapılmadıysa False(Boolean) veya tek bir dosya seçim yapıldıysa 
		dosyanın tam adını(path dahil) veren bir string ya da çoklu seçim 
		yapıldıysa bir dizi. O yüzden dönüş tipi varianttır ve tanımlanırken de böyle 
		tanımlanmalıdır.</p>
		<p>Yine FileDialogda olduğu gibi burda da doğrudan dosya açma veya 
		kaydetme yok sadece dosya ismi elde edilir, sonrasında ayrı bir satırda 
		dosya açma işlemi yapılır.</p>
		<pre class="brush:vb">
Sub getopenfilenameornek()
 
Dim filtreler As String
Dim başlık As String
Dim dosya As Variant
 
    filtreler = "Excel dosysaları(*.xls*),*.xls*, Tüm dosylar (*.*),*.*"
    başlık = "Açılacak dosyayı seçin"
    dosya = Application.GetOpenFilename(filtreler, 5, başlık, , True)
 
    If IsArray(dosya) Then 'multi parametresini true belirlediğimiz içn öncelikle dizi olup olmadıını kontrol etmemiz lazım
        For Each a In dosya
            Workbooks.Open (a)
        Next a
    Else
        If dosya = False Then
            MsgBox "seçim yapılmadı"
        Else
            Workbooks.Open dosya
        End If
    End If

End Sub

</pre>
		<p><strong>SaveAsFileName</strong> de bunun aynısı olup sadece save etmeye yarıyor. </p>
		<p><span class="keywordler">FindFile</span>: Bu da,&nbsp; 
		GetopenFileName ve FileDialog gibi yine Open File dialog kutusunu getirir 
		ancak geriye bir dosya adı döndürmez. Eğer seçim yapıldıysa dosyayı açar 
		ve True değerini döndürür, seçim olmadıysa False döndürür. Bunu 
		diğerleriyle bir anlam bütünlüğü var diye buraya aldım ancak açıkçası 
		pratikte nasıl bir kullanımı olur bilmiyorum, şahsen ben şimdiye kadar 
		hiç kullanmadım.</p>
	
	</div>


	<h2 class="baslik">Diğer üyeler</h2>
	<div class='konu'>
	
	<p><span class="keywordler">Application.ActivateMicrosoftApp metodu:</span>  
	Bu metod, başka bir MS Office uygulamasını açar. Eğer halihazırda ilgili 
	uygulama açıksa onu aktive eder yoksa yenisini yaratır ve açar. Ancak bunu 
	bu şekilde doğrudan kullanmak yerine ilgili Office uygulamasını obje olarak yaratıp onun 
	Nesne modeline ulaşmak istemeniz durumunda ise(ki daha çok bu yöntemi 
	kullanacaksınız) farklı bir yöntem kullanılır. Bunu da bu
	<a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx">buradan</a> 
	görebilirsiniz. Bu metoda dönecek olursak aşağıdaki örnek kodda Word 
	uygulaması açılmakta.</p>
		<pre class="brush:vb"> Application.ActivateMicrosoftApp xlMicrosoftWord</pre>

	<span class="keywordler">Application.Inputbox metodu</span>: Bunu <a href="Temeller_Interaktivite.aspx">burada</a> 
		inceledik.
		<p><span class="keywordler">Application.OnKey metodu</span>: Excel, bir makroyu kaydederken 
		bize bunu bir kısayol tuşuna atayıp atamayacağımız konusunda bir imkan 
		sunar, ancak bunun bir sınırı vardır ki o da sadece <strong>Ctrl</strong> tuşunu 
		kullanmak zorunda olmamızdır ve bu da bir noktadan sonra yetersiz kalmaya başlıyor. İşte bu noktada OnKey 
		metodu yardıma koşuyor, bununla istediğiniz tuş kombinasyonlarına atama 
		yapabiliyorsunuz ve bu tuşlara bastığınızda bir olay tirgger olmuş gibi 
		istediğiniz makro çalışmaya başlıyor.</p>
		<p>Bu tuş kombinasyonu, sadece mevcut Excel oturumunda geçerli 
		olmaktadır. Excel, kapatılıp tekrar açıldıktan sonra kullanılamazlar. 
		Süreklilik kazandırmak için bunları Personal.xlsb dosyanızın 
		Workbook_Open makrosuna yazabilirsiniz.</p>
		<p>Ör:</p>
		<pre class="brush:vb">
Private Sub Workbook_Open()
  Application.CalculationInterruptKey = xlEscKey
  Application.OnKey "+^{F}", "Calculationlar"
End Sub	</pre>
		<p>OnKey metodunu, Çeşitli Windows veya Office kısayol tuşlarını kontrol 
		etmek için de kullanabilirsiniz. Mesela Cut/Copy işlemlerini engellemek 
		için kullanabilirsiniz. Bunu da yine Workbook eventleri ile birlikte 
		kullanmak gerekiyor. Bununla ilgili örnek biraz daha uzun olduğu için 
		onu <a href="Olaylar_WorkbookOlaylari.aspx#cutcopyengel">Workbook 
		Eventleri </a>sayfasına aldım.</p>

<span  id="Versiyon" class="keywordler">Application.Version:</span>Bu özellikle Excelin versiyonunu öğreniyoruz. Böylece kullanıcının Excel versiyonuna göre davranışımızı değiştirebiliyoruz. Mesela 2010 versiyonu ile birlikte gelen Slicer'larla ilgili bir işlemi 2007 ve öncesi kişilerde yapmaya çalışırsak hata alırız. Keza, Slicerlar 2010'da geldi ama sadece Özet tablolarda kullanılmak üzere gelmişti. Table'lar üzerinde kullanımı 2013 versiyonuyla geldi. Bu yüzden bir Table üzerinde Slicer kullanımı olacaksa yine hata alınır.</p>

<p>Şimdi Excelin konuşması özelliğini kullanan başka bir örnek düşünelim. Bu özellik 2013 versiyonuyla geldiği için versiyon numarası 15.0'dır. Versiyon numaralarına <a href="/Konular/Excel/Giris_ExcelinTarihselGelisimi.aspx">buradan </a>ulaşabilirsiniz. Diyelim ki, ortak kullanım için bir Add-in yaptınız ve bu addinde Kokpit adında tüm raporlara ulaşmayı sağlayan bir UserForm var. Kullanıcıların raporlara Kokpit üzerinden ulaşmasını istiyorsunuz, çünkü buradan ulaştıklarında raporlar Readonly açılıyor. Böylece siz raporlarda bir düzenleme yapmak istediğinizde kimsede açık bulunmamış oluyor ve düzenlemelerinizi rahatlıkla yapabiliyorsunuz. Ancak bazı yaramaz arkadaşlar dosyalara ortak alandan ulaşmaya çalışabilir. İşte onlar için ThisWorkbook modülünün Workbook_Open makrosuna aşağıdaki kodu yazabilirsiniz.</p>

<pre class="brush: vb">
Private Sub Workbook_Open()
    'diğer kodlar(Kullanıcının siz olması durumunda aşağıdaki kodun çalışmamasını sağlayacak kodlar dahil, şimdilik kafa karıştırmasın diye bunları atladım)
    If Not Me.ReadOnly Then
        If Val(Application.Version) >= 15 Then
            Application.Speech.Speak ("Hey. Bana ortaka alandan değil, Kokpit formu üzerinden gir") 'Konuşarak iletişim
        Else
            MsgBox "Hey. Bana ortaka alandan değil, Kokpit formu üzerinden gir" 'MsgBox ile iletişim
        End If
        Logger "Bilgi", 0, "Ortak alandan girmeye calisma"
        Me.Close savechanges:=False
    End If

End Sub
</pre>
</div>

</asp:Content>
