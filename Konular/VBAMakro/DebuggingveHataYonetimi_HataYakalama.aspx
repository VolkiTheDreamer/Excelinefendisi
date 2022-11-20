<%@ Page Title='DebuggingveHataYonetimi HataYakalama' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik">
	<div id='gizliforkonu'>
<table>
<tr>
<td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td>
<td><asp:Label ID='Label2' runat='server' Text='Debugging ve Hata Yönetimi'></asp:Label></td>
<td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td>
</tr>
</table>
</div>

<h1>Hata Yakalama(Error Handling)</h1>
	<p> Derler ki Wright kardeşler uçağı uçurmadan önce düşmeyi öğrenmişler. "Nasıl 
iyi bir şekilde düşelim ki ölmeyelim" diye düşünmüşler. Uçağın icadına giden yol 
böyleymiş.</p>
	<p> Biz de makro yazarken çok sık hata yapacağız, ama bu hataları öyle zarif 
	bir şekilde ele almalıyız ki, bi terslik durumunda yumuşak bir iniş olsun.</p>
	<p> İşte bu bölümde program akışı içinde 
	oluşan hataları nasıl ele alacağımızı ve yumuşak iniş yapmayı öğreneceğiz. 
	Zira iyi bir makro içinde mutlaka olası hataları ele alan bir bölüm olmalıdır. </p>
	<h2 class="baslik">Giriş</h2>
		<div class="konu">
	<p> Kodumuzda, oluşabilecek hataları ele aldığımız satırlara Error 
	Handler(Hata Yakalama) satırları denmektedir. Eğer kodumuzda hata yakalama satırı yoksa VBA'in 
	aşağıdaki klasik hata mesajını 
	görürüz. Bu da özellikle programınız kullanacak başka kişiler varsa çok 
	şık olmaz.</p>
	<p> <img src="/images/vbadebuggiris.jpg"></p>
	<p> <strong>Amacımız</strong>, kodumuzun uygun yerlerine hata kontrol noktaları 
	koymak, bunların bazısında program akışını uygun yere yönlendirmek, bazısında 
	hatayı düzeltip tekrar başa(veya başka bir yere) döndürmek veya 
	anlaşılır bir mesaj vererek programı sonlandırmak olmalıdır.</p>
	<p> <strong>Bununla birlikte</strong>, geliştirme sırasında hiç bir hata yakalama kullanmamak uygun olabilir, hatta böylesi daha iyidir, böylece çıkan hata 
	mesajlarında "Debug" 
	düğmesine tıklayarak nerede hata aldığımızı görebiliriz. Ayrıca önceki 
	bölümde gördüğümüz gibi <strong>Debug.Assert</strong> ile de tasarım 
	sırasında da test etmiş oluruz.</p>
	<p> <img src="/images/vbadebugdebbug1.jpg"></p>
	<p> Debug'a basınca hatanın olduğu yer sarılı gösteriliyor.</p>
	<p> <img src="/images/vbadebugdebbug2.jpg"></p>
	<p> <strong>NOT</strong>:Kodunuz korumalıysa Debug düğmesi pasif olur.</p>
	<h3> Hata çeşitleri</h3>
	<p> Farklı kaynaklarda farklı sayılar verilmekle birlikte ben hataları genel olarak 
	2 
	kategoriye ayıryorum:</p>
	<ul>
		<li><strong>Derleme hataları(Compile error)</strong>:Programın derlenmesini engelleyen 
		hatalardır. <strong>Örnekler</strong>:<ul>
			<li>Option Explict açıkken tanımlanmamış değişken kullanımı </li>
			<li>If bloğunda Then yazmadan bir alt satıra geçmek </li>
			<li>Next yazılmamış bir For döngüsü</li>
			<li>Tüm yazım hataları(syntax error):Range yerine Rage yazmak</li>
			<li>Olmayan bir prosedürü çağırmak</li>
			<li>Bir prosedüre modülle aynı ismi vermek</li>
			<li>v.s</li>
		</ul>
		</li>
		<li><strong>ÇalışmaZamanı hataları(Runtime error):</strong>Kodunuz 
		sorunsuz derlenmiştir, ancak 
		çalışma sırasında bir yerde hata verebilir. Örnekler:
		<ul>
			<li>0'a bölme, </li>
			<li>Bir değişkene taşıyabileceği değerden daha fazla değer 
			atanması(integer'a 50.000 atanması gibi)</li>
			<li>Olmayan bir sayfaya/dosyaya erişmek</li>
			<li>v.s</li>
		</ul>
		</li>
	</ul>
	<p> <strong>NOT</strong>:Yazımsal kaynaklı(syntax) compile hataları sonucunda çıkan dialog kutuları bazen can 
	sıkıcı olabilmektedir. Bunların çıkmasını engelleyebilirsiniz, tabi yine de 
	hatalı olan kod kırmızı olarak gösterilecektir. Böylece hatanızı hala anında görmeye 
	devam edersiniz. Bununla sadece dialog kutusunda kurtulmuş oluyoruz. 
	Baktınız ki bi türlü hatanın nedenini çözemiyorusunuz, bu işlemi geriye alıp 
	tekrar mesaj çıkmasını sağlayabilirsiniz.</p>
	<p> <img src="/images/vbahatayakala1.jpg"></p>
	<p>Derleme hatası olup olmadığını <strong>Debug&gt;Compile</strong> menüsünden test 
	edebilirsiniz. Varsa tekrar yapın, taa ki bu düğme gri, yani pasif olana kadar.</p>
	<p>Runtime hatalarını <span class="keywordler">On Error</span> ile başlayan 
	cümleciklerle ele alacağız. Bunlara geçmeden önce bir de beklenti türüne 
	göre hataları inceleyelim.</p>
	<h4>Beklentiye göre Hata çeşitlieri</h4>
	<p>Hatalar beklentiye göre iki gruba ayrılır.</p>
	<ul>
		<li>Beklenen hatalar</li>
		<li>Beklenmeye hatalar</li>
	</ul>
	<p>Program akışı sırasında bir yerde <strong>hata olma olasılığı varsa</strong>, 
	yani hata beklenen bir hataysa bunları koşullu yapılarla yakalayabiliriz.</p>
	<p>Örneğin bir metin dosyasından okuma yapmadan önce(veya bir excel 
	dosyasını açmadan önce) onun varolup 
	olmadığını öğrenmeye çalışmak akıllıca olur. Dosya varsa okuma yaparız, yoksa yapmayız 
	ve kullanıcıya bi mesaj çıkarırız. Böyle <strong>spesifik</strong> bir beklentiye karşı 
	spesifik bir mesaj vermiş oluruz. </p>
<pre class="brush:vb">
Sub OpenFile()
    Dim dsy As String
    dsy = "C:\deneme.txt"
    
    ' DosyaVarmı diye bir fonksiyon yazmış olalım, bu fonksiyon ile onun varlığını sorguluyoru
    If Not DosyaVarmı(sFile) Then 'DosyaVarmı(dsy)=False da yazabilrdik
        MsgBox "Dosya mevcut değil, dosya adının doğru olduğundn emin olup tekrar deneyin"
			Exit Sub
    Else
        'yapılacak işler
    End If
       
End Sub</pre>
	<p> Keza yine iki sayının toplamını aldıracağız diyelim. Burada kullanıcı 
	sayılardan birini veya ikisini de nümerik olmayan bir değer girebilir. Bu 
	beklenen bir durum olup bunu da If blokları ile yakalarız.</p>
	<p> Aşağıda göreceğimiz hata kodlarını ise genelde beklenmeyen hata durumlarında 
	kullanırız. Yani programa, <strong>genel</strong> bir hata durumunda ne 
	yapacağını söylemiş oluruz.</p>
	</div>
	
	
	<h2 class="baslik">On Error ifadesi</h2>
	<div class="konu">
	<p> On Error kalıbı hata yakalamanın kalbidir ve bu kalıbı takip eden bir dizi ifade vardır. Bunlar, 
	"hata olması 
	durumunda şunu yap" şeklinde özetlenebilecek ifadelerdir. </p>
	<ul>
		<li>On Error Resume Next</li>
		<li>On Error GoTo <em>etiket</em></li>
		<li>On Error GoTo 0</li>
		<li>On Error GoTo -1</li>
	</ul>
	<p> Şimdi bunlara tek tek bakalım</p>
	<h3> On Error Resume Next</h3>
	<p> Bu kalıp, VBA'e hatayı görmezden gelmesini ve bir sonraki satıra 
	geçmesini söyler. Yani hatayı düzeltmek yerine yola devam eder. O yüzden çok 
	dikkat edilmesi gereken bir kalıptır. Birçok durumda bunu kullanmaktan sakınmanız 
	gerekir, bütün üstatlar da bunu söyler zaten, ancak belli bazı durumlar vardır ki 
	bu kalıbı kullanmak akıllıca olmaktadır. </p>
	<p> Mesela 
	bankada çalışıyoruz diyelim ve bir otomatik mail makromuz olsun. Bankadaki belli bir roldeki tüm 
	personele, kişiye özel olacak şekilde mail gönderiyor olalım. Ancak 
	elimizdeki kişi listesi güncel olmayabilir, yani bu kişilerden bazıları 
	artık bankadan ayrılmış olabilir. Böyle bir durumda programa şunu 
	söyleyebiliriz: "Sırayla herkese mail at, maili bulunmayanlar için durma, 
	bir sonrakine devam et". Bunu demezsek, program, ilk ayrılmış kişide durur. 
	Bunu ele almanın başka yöntemleri de var tabi; Resolve metodunu kullanmak, On 
	Error GoTo etiket diyip, ilgili kaydı formatlamak ve bir sonrakine devam 
	etmek gibi ama konuyu anlatmak adına basit düşünüyoruz.</p>
	<p> İkinci bir örnek, her gün belirli bir saate schedule edilmiş ve o günün sonucunu tarih adıyla kaydeden bir rutininiz olsun. 
	Bir süre sonra 
	diskinizde şişkinlik olmasın diye her gün t-30 tarihli raporları silmeye yarayan 
	da başka bir rutininiz olsun. Yani bize sadece son 30 günün raporu yeter 
	diyoruz. Ama bazı günler raporlarınız oluşmamış olabilir, o yüzden 
	varolmayan bir t-30 raporunu silmeye çalışabilirsiniz. Böyle günlük schedue 
	edilmiş 40 raporunuzun ikisi için rapor oluşmadığını düşünelim. Diğer 38ini tehlikeye 
	atmamak 
	için yine bu kalıbı kullanabiliriz. Tabi bunu bir alternatifi de dosya mevcut mu 
	diye kontrol etmek de olabilir&nbsp;ama bu kontrol de kodun süresini uzatacaktır. 
	Böyle bir durumda bu kalıbı kullanmak daha pratik bir çözüm olacaktır.</p>
	<p> Aşağıda ilk örneğimize ait kod bloğunu görüyoruz.</p>
	<pre class="brush:vb">Sub onerrorresumeornek()
'Dim ile değişken deklerasyonları

'1000 diye sabit bir rakam olmaz tabi, bunu son satır gibi bir yöntemle yakalamanız lazım _
örnek basit olsun diye 1000 yazdım
For i=1 to 1000 

On Error Resume Next 'sadece döngü içinde hata olursa bir sonraki satıra geçsin

'mail gönderim kodları

Next i

'varsa diğer kodlar

End Sub</pre>
	<h3> On Error GoTo Etiket</h3>
	<p> Bu kalıp, tüm hata türlerini yakalar ve programı etiketin olduğu yere 
	yönlendirir. Bu kısımda genelde şunlar yapılır:</p>
	<ul>
		<li>Hata mesajları veririz</li>
		<li>Olası hatayı düzeltebiliyorsak düzeltir ve hata aldığımız yere veya 
		hemen sonraki satıra yönlendiririz</li>
		<li>Programın başında False atadığımız 
	bazı özellikleri burada tekrar True'ya döndürürüz</li>
		<li>Sayfa veya Workbook seviyesinde protection'ı geçici kaldırdıysak 
		tekrar protection uygularız</li>
		<li>Log tutarız</li>
	</ul>
	<p> Etiketin hemen öncesinde bir <strong>Exit Sub </strong>ifadesinin olmasında fayda vardır, yoksa hatasız 
	ilerleyen kod son olarak 
	buraya gelir ve gereksiz yere burdaki kodları da çalıştırır, ki bunu 
	istemeyiz.</p>
	<p> Hata mesajında genelde <strong>Err.Description </strong>ve <strong>Err.Number 
	</strong>özelliklerini kullanırız. 
	Err.Number için bir koşullu yapı kullanılarak, "Hata numarası şuysa şu mesajı, buysa bu 
	mesajı, diğer durumlar için Err.Description &amp; "Volkanla görüşün" " gibi bir 
	mesaj verdirebiliriz.</p>
	<pre class="brush:vb">
Sub Hataornek()
On Error Goto hata

With Application
    .DisplayAlerts=False
    .ScreenUpdating=False
End With

   'çeşitli kodlar

'program bitişi öncesinde bu ayarları eski haline getriyoruz
With Application
    .DisplayAlerts=True
    .ScreenUpdating=True
End With

'hata etiketinden hemen önce programdan çıkıyoruz
Exit Sub
hata:
'hata durumunda bu ayarları eski haline getriyoruz
With Application
    .DisplayAlerts=True
    .ScreenUpdating=True
End With
MsgBox Err.Description &amp; vbcrlf &amp; "Bu bilgi yeterli değilse Vlkanla görüşün lütfen"
End Sub
</pre>
	<p><span class="dikkat">Dikkat:</span>Bu kalıbı kullanırken akılda bulundurulması gereken önemli bir husus var, 
	o da bir anda sadece tek bi hata yakalama bloğunun aktif olabileceğidir. Yani 
	bi hata için <strong>On Error GoTo hata1</strong> dediniz diyelim, hata1 bloğunda da bir hata meydana 
	gelirse oraya da <strong>On Error GoTo hata2 </strong>yönlendirmesi 
	yaparsanız bu etkisiz 
	kalacatır, çünkü ilk hata yakalayıcı hala aktiftir. Bu durumu aşmanın bir yolu var 
	ve biraz 
	sonra bunu göreceğiz. Siz şimdi sadece, bunun gerçekten işe yaramadığını görmek için şu 
	kodu çalıştırmayı deneyin.</p>
	<p>Bu kodda, önce 0'a bölme hatası yapıyoruz. Bu bizi hata1 etiketine 
	yönlendiriyor. Hata1 Error Handlerında da, bu hatayı erişimimiz olmayan bir 
	dosyaya yazmaya ve kaydını tutmaya çalışıyoruz. Yine hata alınca bu sefer 
	hata mesajı görüntüleniyor. &nbsp;</p>
	<pre class="brush:vb">Sub cokluhata()
On Error GoTo hata1

x = 0
MsgBox 1 / x
MsgBox "İşlem sonucu başarıyla gösterildi"

Exit Sub
hata1:
On Error GoTo hata2

dosya = "C:\deneme.txt" 'bu dosyayı yönetici izni sağlayarak oluşturdum, yani dosya var,
'ancak izin isteyen bir konum olduğu içn dosyaya yazma sırasında hata alacağım
Open dosya For Append As #1
Write #1, Now, Err.Number, Err.Description, Environ("username")
Close #1

Exit Sub
hata2:
MsgBox Err.Description, , "Hata"

End Sub</pre>
	<h3> On Error GoTo 0</h3>
	<p> 
	Hiçbir hata yakalama kodu yazmazsak, VBA ilk hatada bize 
	hata mesajı gösterir 
	demiştik. İşte <strong>On Error GoTo 0</strong> da, VBA'in default modu olan "Hata yakalama 
	bloğu yok, tüm hatalar için uyarı göster"e 
	dönüş yapmamızı sağlar. Böyle birşeyi neden isteyelim ki? Birazdan 
	göreceğiz. Ama şunu anlamışsınızdır ki, bu kalıbı direkt programın başında kullanmak 
	anlamsızdır, zaten default ayar öyledir.</p>
	<p>
	Bu kalıbı daha çok, varolan bir hata yakalama rutini iptal etmek için kullanırız. 
	Neden böyle birşeye ihtiyaç duyacağımızı bir düşünün! Sonra devam edin.</p>
	<p>Mesela 
	yukarıda bahsettiğimiz otomatik mail gönderim 
	makrosunu düşünün. Yaklaşık 1000 kişiye mail gidecek, ama gönderim yapacağınız 
	kişilerden bazılarına mail gönderilemiyorsa bir sonrakine geçsin demiştik. Bunun için de 
	<strong>On Error Resume Next</strong> demiştik. Ancak mail gönderememe dışında başka bir sorun 
	varsa bunun size veya kullanıcıya gösterilmesini istemiş olabilirsiniz. Ne 
	tür hatalar çıktığını gördükten sonra bunlara özgü çözümlerinizi de 
	ayarlayabilesiniz diye On Error GoTo kalıbını kullanabilirsiniz.&nbsp;Sonrasında ya If blokları ile yakalarsınız veya bir etikete 
	göndererek işlem yaparsınız. Örnek bir kod bloğu şöyle olacaktır:</p>
	<pre class="brush:vb">Sub onerrorgotosfır()
'Dim ile değişken deklerasyonları

'1000 diye sabit bir rakam olmaz tabi, bunu son satır gibi bir yöntemle yakalamanız lazım _
örnek basit olsun diye 1000 yazdım
For i=1 to 1000 

On Error Resume Next 'sadece döngü içinde hata olursa bir sonraki satıra geçsin

'mail gönderim kodları

ooMail.Send
'On Error GoTo 0 buraya da konabilir, aşağıya da

Next i

'buradan sonra bir hata çıkarsa hata mesajını görelim
On Error GoTo 0

'diğer kodlar

End Sub</pre>
	<h3>
	On Error GoTo -1</h3>
	<p>
	Yukarda
	bir programda <strong>sadece</strong> <strong>bir</strong> adet hata yakalama bölümü aktif 
	olabilir demiştik, dolayısıyla bir Error Handler bloğunda başka bir Error 
	handlera yönlendirme yapılırsa bu işe yaramıyordu. İşte <strong>On Error GoTo -1</strong> ile bu sorunu 
	aşmış oluyoruz. Bu kalıp, varolan Error Handler'ı resetler(<strong>Err</strong> 
	nesnesini yokeder) ve sizin yeni bir Error handler yaratmanıza imkan verir. 
	(Bunun, aşağıda anlatacğımız <strong>Err.Clear </strong>ile karıştırılmaması lazım.
	<strong>Err.Clear 
	</strong>sadece hata bilgilerini yokeder, <strong>Err</strong> nesnesini değil).<span>
	</span><strong>Sözün özü</strong>; bu kalıp, başka bir Error Handler bloğu 
	içinde kullanılır ve varolan Error handler içinden başka bir Error 
	handlera yönledirmenin tek yolu budur. On Error Resume Next bile burada işe yaramaz.</p>
	<p>
	Şimdi yukarda örneği tekrar alalım ve hata1'den hemen sonra kalıbımızı 
	yerleştirelim.</p>
	<pre class="brush:vb">Sub cokluhata()
On Error GoTo hata1

x = 0
MsgBox 1 / x
MsgBox "İşlem sonucu başarıyla gösterildi"

Exit Sub
hata1:
On Error GoTo -1 'bu satırın başına yorum işareti koyarak etkisini görün
On Error GoTo hata2


dosya = "C:\deneme.txt" 'bu dosyayı yönetici izni sağlayarak oluşturdum, yani dosya var,
'ancak izin isteyen bir konum olduğu içn dosyaya yazma sırasında hata alacağım
Open dosya For Append As #1
Write #1, Now, Err.Number, Err.Description, Environ("username")
Close #1

Exit Sub
hata2:
MsgBox Err.Description, , "Hata"
End Sub</pre>
	<h3>İkisi bir arada örnek</h3>
		<p>Inputbox konusunu anlatıkren görmüştük ki, dönüş tipi Range olan bir 
		Inputbox sorusuna Cancel diyip çıkarsak program hata alıyordu. Bunun 
		için "On Error Resume Next" diyerek hata mesajını yokediyorduk. Peki ya 
		sonraki hatalarda devam etmesini istemiyorsak, o zaman yeni bir hata 
		yakalama bloğu set etmeliyiz. Ör:</p>
		<pre class="brush:vb">
Sub hataikili()
Dim a As Range

On Error Resume Next 'geçici olarak tanımlanır
Set a = Application.InputBox("Bir hücre seçin", Type:=8) 'buna esc dersek diye, commenteleyip görelim

On Error GoTo hata 'esas hata yakalama kodumuzu yazıyoruz
If Not a Is Nothing Then
	MsgBox "Seçim yapıldı"
	'Diğer kodlar buraya
Else
	MsgBox "Bir seçim yapılmadan çıkmayı tercih ettiniz"			    
End If

MsgBox 1 / 0
MsgBox 2 / 4 'buraya ulaşmaz

Exit Sub
hata:
	MsgBox Err.Description
End Sub</pre>
		<h3>Kapsamlı bir örnek</h3>
	<p>Şimdiki örnekte 4 çeşit <strong>On Error </strong>ifadesini de 
	kullanıyoruz. Bu açıdan bu örneğin çok faydalı olacağını düşünüyorum. Özeti 
	şu:Networkte duran ve her gece refresh olması beklenen bir dosya schedule 
	edilmiş durumda. Schedule saati gelmiş ve dosya açılınca onun içindeki 
	ThisWorkbook_Open makrosu devreye girecek ve dosyadaki tablolar refresh 
	olacak. Refresh sonucunda data dönüyorsa(yani 
	veritabanına ilgili data yüklenmişse) mailing işlemine devam edecek, henüz 
	data yüklenmediyse dosyayı kaydetmeden çıkacak. Aradaki detaylar kafanızı karıştırmasın, sadece 
	hata bloklarına odaklanın.</p>
	<pre class="brush:vb">Sub musteri_degisim()
Dim OutApp As Object
Dim OutMail As Object
Dim alicilar As Object
Dim scl As Range
Dim rng As Range
Dim alan As Range


Workbooks.Open Filename:= _
        gunlukyol + "\Müşterisi değişen miyler\Müşterisi_Değişen_Miyler.xlsm"

'outlook nesnemizi yaratıyoruz
Set OutApp = CreateObject("Outlook.Application")

'ilk hata yakalama kodumuzu yazıyoruz
On Error GoTo cleanup

'kodlarımız hızlı çalışsın diye ayarlamalarımızı yapıyoruz
With Application
   .EnableEvents = False
   .ScreenUpdating = False
End With

'ANA kod bloğu burada başyor, 
Sheets(1).Select

If IsEmpty(Cells(2, 1)) Then
   ActiveWorkbook.Close savechanges:=True
   Exit Sub 'hiç kayıt gelmediyse çıksın
End If

For i = 2 To Cells(1, 1).End(xlDown).Row
	Cells(i, 2).Select
	Set OutMail = OutApp.CreateItem(0)
	Set alicilar = OutMail.Recipients.Add(ActiveCell.Value)
	
	If ActiveCell.Value = ActiveCell.Offset(1, 0).Value Then
	'yani hem giren hem çıkan varsa
		If ActiveCell.Offset(0, 1).Value = "Portföye Giren" Then
		
		'ikinci hata bloğu, hata olma ihtimali var, çünkü mail gönderemye çalışacağım kişi bankadan istifa etmiş olabilir
		' ve mail önderimi sırasında hata oluşabilir, o yüzden hata oluşursa devam et diyorum
		On Error Resume Next
		With OutMail
			.SentOnBehalfOfName = "BBSatisPerformans@akbank.com"
			.Subject = "Portföyünüze giren ve portföyünüzden çıkan müşteriler hakkında"
			.htmlbody = "Değerli Miyimiz," &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "Dün itibarıyle portföyünüze " &amp; ActiveCell.Offset(0, 2).Value &amp; " adet müşteri girişi olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 3).Value), 0, ActiveCell.Offset(0, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 4).Value), 0, ActiveCell.Offset(0, 4).Value)) &amp; " TL'dir. (Bu listeye, hesap devriyle gelen ve yeni yaratılan mbbler dahil değildir.)" &amp; Chr(14)
			.htmlbody = .htmlbody + "Ayrıca portföyünüzden " &amp; ActiveCell.Offset(1, 2).Value &amp; " adet müşteri çıkışı olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(1, 3).Value), 0, ActiveCell.Offset(1, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(1, 4).Value), 0, ActiveCell.Offset(1, 4).Value)) &amp; " TL'dir." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "MBB detayına ulaşmak için OPERA'dan 'Müşterisi değişen Miyler' raporuna bakabilirsiniz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "Bilginizi rica ederiz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;Saygılarımızla, &lt;/B&gt;" &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;&lt;FONT COLOR=""Red""&gt;" &amp; "Bireysel Bankacılık Satış Yönetimi &lt;/FONT&gt;&lt;/B&gt;"
			alicilar.Resolve
			If Not alicilar.Resolved Then
				GoTo sonraki
			End If
			.Send
		End With
		
		Else ' ilk satır portföyden çıkansa
		
		With OutMail
			.SentOnBehalfOfName = "BBSatisPerformans@akbank.com"
			.Subject = "Portföyünüze giren ve portföyünüzden çıkan müşteriler hakkında"
			.htmlbody = "Değerli Miyimiz," &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "Dün itibarıyle portföyünüzden " &amp; ActiveCell.Offset(0, 2).Value &amp; " adet müşteri çıkışı olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 3).Value), 0, ActiveCell.Offset(0, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 4).Value), 0, ActiveCell.Offset(0, 4).Value)) &amp; " TL'dir." &amp; Chr(14)
			.htmlbody = .htmlbody + "Ayrıca portföyünüze " &amp; ActiveCell.Offset(1, 2).Value &amp; " adet müşteri girişi olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(1, 3).Value), 0, ActiveCell.Offset(1, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(1, 4).Value), 0, ActiveCell.Offset(1, 4).Value)) &amp; " TL'dir. (Bu listeye, hesap devriyle gelen ve yeni yaratılan mbbler dahil değildir.)" &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "MBB detayına ulaşmak için OPERA'dan 'Müşterisi değişen Miyler' raporuna bakabilirsiniz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "Bilginizi rica ederiz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;Saygılarımızla, &lt;/B&gt;" &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;&lt;FONT COLOR=""Red""&gt;" &amp; "Bireysel Bankacılık Satış Yönetimi &lt;/FONT&gt;&lt;/B&gt;"
			alicilar.Resolve
			If Not alicilar.Resolved Then
			   GoTo sonraki
			End If
			.Send
		End With
		
		'istifa eden kişiye mail gönderme hatası dışında bir hata olursa 
		'bu hatayı greyim ki düzlteyim istiyroum; tabi ekran o an takılı kalır, sornaki schdeul rogramlar aksıya alınır 
		'ama olsun, hatayı düzeltmek adıan bunu yapmaız lazım
		On Error GoTo 0
		End If
	i = i + 1
	Else 'yani tek satır varsa
		'yukardaki nedenlerin aynı sebeple resume next
		On Error Resume Next
		With OutMail
			.SentOnBehalfOfName = "BBSatisPerformans@akbank.com"
			.Subject = "Portföyünüze giren ve portföyünüzden çıkan müşteriler hakkında"
			.htmlbody = "Değerli Miyimiz," &amp; Chr(14) &amp; Chr(14)
			If ActiveCell.Offset(0, 1).Value = "Portföye Giren" Then
			   .htmlbody = .htmlbody + "Dün itibarıyle portföyünüze " &amp; ActiveCell.Offset(0, 2).Value &amp; " adet müşteri girişi olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 3).Value), 0, ActiveCell.Offset(0, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 4).Value), 0, ActiveCell.Offset(0, 4).Value)) &amp; " TL'dir. (Bu listeye, hesap devriyle gelen ve yeni yaratılan mbbler dahil değildir.)" &amp; Chr(14) &amp; Chr(14)
			Else
			   .htmlbody = .htmlbody + "Dün itibarıyle portföyünüzden " &amp; ActiveCell.Offset(0, 2).Value &amp; " adet müşteri çıkışı olmuştur. Bu müşterilerin 2 gün önceki TMV bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 3).Value), 0, ActiveCell.Offset(0, 3).Value)) &amp; " TL, Kredi bakiyesi " &amp; Int(IIf(IsError(ActiveCell.Offset(0, 4).Value), 0, ActiveCell.Offset(0, 4).Value)) &amp; " TL'dir." &amp; Chr(14) &amp; Chr(14)
			End If
			.htmlbody = .htmlbody + "MBB detayına ulaşmak için OPERA'dan 'Müşterisi değişen Miyler' raporuna bakabilirsiniz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "Bilginizi rica ederiz." &amp; Chr(14) &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;Saygılarımızla, &lt;/B&gt;" &amp; Chr(14)
			.htmlbody = .htmlbody + "&lt;B&gt;&lt;FONT COLOR=""Red""&gt;" &amp; "Bireysel Bankacılık Satış Yönetimi &lt;/FONT&gt;&lt;/B&gt;"
			alicilar.Resolve
			If Not alicilar.Resolved Then
			   GoTo sonraki
			End If
			.Send
		End With
		'yine aynı sbeebple goto 0
		On Error GoTo 0
		
	End If
	
	sonraki:
	Set OutMail = Nothing

Next i


cleanup:
On Error GoTo -1 'Burdada bir hata çıkarsa öncekini resetliyorum
On Error GoTo hata 'hata çıakrsa "hata" etiketine yönlendiriyoru

Set OutApp = Nothing
With Application
.EnableEvents = True
.ScreenUpdating = True
End With

Windows("Müşterisi_Değişen_Miyler.xlsm").Close savechanges:=True

rapor = "Müşteri-Miy değişim"
alici = "13245;32222"
Call Mailat2(rapor, alici)

Exit Sub
hata:
Logcu Now, Err.Description 'Looger isimli loglama proseürümle hata kaydı tutuyorum
End Sub
</pre>
	<h3>
	Error handlerdan çıkma yolları</h3>
	<p>
	GoTo -1 bölümünde yaptığım açıklamlar şöyle bir yanlış anlaşılmaya sebep olmamalı. 
	GoTo -1, <strong>bir error handler içinden 
	başka bir error handler içine göndermenin </strong>tek yoludur, <strong>error handler içinden 
	çıkmanın </strong>tek yolu değil. Error handler içinden On Error GoTo -1 
	kalıbına ilave 
	olarak <strong>Resume, Resume Next, Resume etiket </strong>veya <strong>Exit 
	Sub </strong>diyerek de 
	çıkabilirsiniz. Bu şekillerde çıkış yapıldığında da Error Nesnesi resetlemiş 
	olur.</p>
	<p>
	Şimdi bunlara bakalım:</p>
	<h4> Resume</h4>
	<p>Resume tek başına kullanıldığında hata olan yerden tekrar devam etmeye 
	çalışır. Bunun için tabi error handlerda hatayı düzeltmeniz gerekir. Duruma 
	göre kullanıcıya düzeltmeyle ilgili bilgi vermek de gerekebilir. </p>
	<p><span class="dikkat">Dikkat:</span>Hatayı 
	düzeltmeden Resume dersek program sonsuz bir kısırdöngüye girer. </p>
	<pre class="brush:vb">Sub resume1()
On Error GoTo hata

x = InputBox("0 ile 100 arasında bir sayı girin") '0 dahil mi belli dğeil

sayı = 1 / x
Debug.Print sayı

Exit Sub
hata:
If x=0 Then 
   x=1
   MsgBox "Girdiğiniz sayı aralık dışı olup en düşük 1 girmeniz gerkeirdi. Sayı otomatikman 1e yuvarlandı""
   Resume
Else
   MsgBox Err.Description 'mesela bir harf girilise hata versin
End If
End Sub</pre>
	<p>Başka bir örnek de şu olabilr. Biz biliriz ki nerdeyse tüm Workbooklarda 
	Sheet1 diye bir dosya vardır, ve kodumuz içinde kullanıcıdan bir dosya 
	seçmesini istedik diyelim, sonra gidip Sheet1 isimli sayfaya bazı şeyleri 
	otomatik yazacağız. Eğer Sheet1 diye bir sayfa yoksa hata verecektir. Bunu aşağıdaki gibi ele 
	alabiliriz.</p>
	<pre class="brush:vb">
Sub resume3()
On Error GoTo hata:

    'Kullanıcıdan dosya seçtiren kodlar
    Worksheets("Sheet1").Activate
    'diğer işlemler
    hata:
    If Err.Number = 9 Then 'sayfa yoksa
		Worksheets.Add.Name = "Sheet1"
		Resume 'Worksheets("Sheet1").Activate satırına geri döner
	End If
End Sub</pre>
	<p>Aşağıdaki örnekte de Error handler içinde hatalı kaydı düzeltip aynı 
	kayıtla(sonraki değil) devam etmesini sağlıyorum:</p>
	<pre class="brush:vb">Sub resume4()
On Error GoTo hata

Dim kareal As Variant
kareal = Array(9, 16, -64, 100) 'yanlışlıkla 64ün önüne - yazmışım diyelim

For Each a In kareal
MsgBox a &amp; " sayısının kökü:" &amp; Sqr(a)
Next a

Exit Sub
hata:
a = Abs(a)
Resume
End Sub</pre>
	<h4>Resume etiket</h4>
	<p>Hata bloğunda gerekli bir mesaj verdikten sonra bir etikete yönlendirme 
	işlemini <strong>Resume Etiket </strong>kalıbı ile yaparız. Burada alternatif olarak 
	<strong>GoTo Etiket </strong>kalıbı 
	kullanıldığı görülmekle birlikte bu kullanım beklenmeyen hatalara neden 
	olabilir. Bu yüzden Error handlerlar içindeyken GoTo yerine Resume kullanmanızı öneriyorum.</p>
	<pre class="brush:vb">Sub resume2()
On Error GoTo hata

yeniden:
x = InputBox("0 ile 100 arasında bir sayı girin") '0 dahil mi belli dğeil

sayı = 1 / x
Debug.Print sayı

Exit Sub
hata:
If x=0 Then 
   x=1
   MsgBox "Girdiğiniz sayı aralık dışı olup en düşük 1 girmeniz gerkeirdi. Sayı otomatikman 1e yuvarlandı""
   Resume 'burada da istersek kullanıcıya tekrar seçim hakkı vermek isteyebilri virz
Else
   If Err.Number=13 Then 'Type Mismatch hatasıysa
      MsgBox "Lütfen geçerli bir sayı giriniz"
      Resume yeniden 'GoTo etiket de alternatif olmakla birlikte kullanmayın   Else
   Else 'başka bir hataysa
      MsgBox Err.Description 
   End If
End If
End Sub</pre>
	<h4>Resume Next</h4>
	<p>Mesela otomatik mail programınızda sicili olmayan kişiler çıkarsa 
	program akışı durmasın diye <strong>On Error Resume Next </strong>demiştik ya, 
	bunun bir diğer 
	ve şık alternatifi de şu olabilir. <strong>On Error Goto hata </strong>dersiniz, hata 
	bölümünde bir log kaydı oluşturabilir ve/veya mail gitmeyen kişilerin listesi o an 
	açık olan bir excel dosyasında ise onları renklendirebilirsiniz, böylece 
	kimlere mail gitmediğini de görmüş olursunuz. Zira belki sorun To'ya 
	koyduğunuz kişide değil, cc'deki kişiden kaynaklı da olabilir. </p>
	<p>Tabi bu 
	örnekte başka bir hata durumunda istenmeyen sonuçlar elde edebilirsiniz, o 
	yüzden kodunuzu bu duruma göre ayarlamanız gerekecektir. Burada basitlik olması adına 
	sadece olmayan bir adrese mail gitme ihtimali olduğunu düşündük.</p>
	<pre class="brush:vb">Sub resumenext1()
'Dim ile değişken deklerasyonları
On Error GoTo hata

'1000 diye sabit bir rakam olmaz tabi, bunu son satır gibi bir yöntemle yakalamanız lazım _
örnek basit olsun diye 1000 yazdım
For i=1 to 1000 

'mail gönderim kodları

Next i

'varsa dieğr kodlar

hata:
'Loglama veya formatlama işlemi
Resume Next

End Sub</pre>
	<p>Aşağıdaki örnekte ise Error Handlerda hatayı yakaldıktan sonra durmasını 
	istemiyorum ve diğer sayılar için devam etmesini istiyorum, ama devam etmeden 
	önce de bi mesaj vermesini istiyorum.</p>
	<pre class="brush:vb">Sub resumenext2()
On Error GoTo hata

Dim kareal As Variant
kareal = Array(9, 16, -1, "A", 81, "B", 100)

For Each a In kareal
MsgBox a &amp; " sayısının kökü:" &amp; Sqr(a)
Next a

Exit Sub
hata:
MsgBox a &amp; ":Karekökü alınacak bir sayı değil"
Resume Next
End Sub</pre>
	<h3>Err nesnesi</h3>
	<p>Bir runtime hatası oluştuğunda Err nesnesi oluşturulur ve bu nesnenin 
	özellikleri bu hatanın detayıyla doldurulur. </p>
	<p>Bu nesnenin üyeleri, <strong>Exit</strong>'li ifadelerden(Exit Sub gibi) ve 
	<strong>Resume 
	Next</strong>'ten sonra resetlenir, yani 0 ve ""'a döner. Veya manuel resetlemek için 
	Err.Clear metodu kullanılır.</p>
	<p>Err'in default property'si(özelliği) Numberdır. Bir de Description vardır. 
	Bunları yukardaki 
	örneklerde bolca kullandık. Genelde bu descriptiondaki bilgi 
	son kullanıcıya yeterli bilgi vermez, hele bi de İngilizceleri yoksa birşey 
	anlamazlar, o yüzden hata numarasına göre kendi mesajlarınızı belirtebilirsiniz.</p>
<pre class="brush:vb">
Select Case Err.Number
Case Is = 5
  MsgBox "Negatif sayıların karekökü alınamaz"
Case Is = 13
   MsgBox "Metinsel ifadelerin karekökü alınamaz"
End Select
</pre>
	<p>Çıkan hatanın Number özelliği yoksa Description öelliğinde "Application-defined or object-defined error" mesajı 
	gösterilir. Herhangi bir anda hata yoksa Number=0 ve Description="" 
	değerindedir. O yüzden belirli bir anda hata olup olmadığını If Err.Number=0 
	ile kontrol edebilirsiniz.</p>
	<h4>Err.Raise</h4>
	<p>Err nesnesi oluşturmanın bir yolu da <strong>Raise</strong> metodunu kullanmaktır. 
	Bu metodla birlikte kendimize özgü Hata Kaynağı, Hata Numarası ve Hata 
	açıklaması üretiriz. Hata numarasını 513 ile 65535 arası bir değer 
	atayabilyoruz ve bunu vbObjectError ile birlikte 
	kullanıyoruz.</p>
	<p>Err.Raise'in bir alternatifi <strong>GoTo hata </strong>diyerek ilerlemek 
	olabilir ancak bu şekilde hatayı kayıt altına 
	alamayız. Özellikle birden fazla yönlendirme olacaksa GoTo etiket yerine 
	Err.Raise 
	diyip açıklayıcı bilgilerle de donatmak gerekir.</p>
	<pre class="brush:vb">Sub errraise()
On Error GoTo hata

x = InputBox("1000den küçük bir sayı girin") 'çok net değiliz, belki kullanıcı 0 girecek, belki 1000 dahil sanacak

'bunu direkt select case içinde kullanamadığımız için if blok içinde kullanıyorum
If Not IsNumeric(x) Then
   Err.Raise 517 + vbObjectError, "Sayısal olmayan veri", "Kullanıcı sayısal olmayan bir değer girdi"
End If

Select Case x
Case 0
  Err.Raise 513 + vbObjectError, "x nedeniyle 0'a bölme hatası", "Kullanıcı x=0 girdi"
Case 1000
  Err.Raise 514 + vbObjectError, "x nedeniyle 0'a bölme hatası", "Kullanıcı x=1000 girdi"
Case Is &gt; 1000
  Err.Raise 515 + vbObjectError, "aralık dışı değer", "Kullanıcı x&gt;0 girdi"
Case Is &lt; 0
  Err.Raise 516 + vbObjectError, "aralık dışı değer", "Kullanıcı x&lt;0 girdi"
Case Else
  sayı = (1 / x) * (1 / (1000 - x))
End Select
Debug.Print sayı

Exit Sub
hata:
Debug.Print Err.Source, Err.Number, Err.Description
End Sub</pre>
	<h4>Err.Source</h4>
	<p>Bu özellik, özellikle Log tutulurken faydalı olabilmektedir, ve 
	yukarıdaki örnekte görüldüğü gibi, Raise ile manuel yaratılan hata 
	mesajlarında bizim tarafımızdan belirtildiğinde anlamlıdır, yoksa her zaman 
	VBAProject gibi birşey ifade etmeyen bir sonuç üretir. (E neden var o zaman? 
	Çünkü bu özellik VB dilinin kendisinde var ve orada anlamlı ama VBA içinde 
	anlamsızlaşıyor, eğer ki biz kendimiz belirtmezsek!)</p>
	<h4>Err.Clear</h4>
<p> Bu metod yukarda da gördüğümüz gibi hata nesnesinin üyelerini sıfırlar. 
Pratik olarak buna ihtiyaç duyabileceğimiz tek yer On Error Resume Next dendiği 
durumlarda, karşılaşılan hata adedini saymak için kullanmak olacaktır. Zira ben 
başka bir kullanımını görmedim.</p>
	<p> Şöyle ki, döngüsel bir kodumuz var diyelim, döngünün bir yerinde Err.Number&lt;&gt;0 
	kontrolüne takılır ise hata oluşmuş demektir, bu durumda hatasayısı=hatasayısı+1 
	diyerek sayacı arttırırız. Ama arada birçok yerde hatasız geçiş olacağı için ama en sonki kod numarası da 
	hala hafızada olacağı için bir kez hata oluştuğunda döngünün sonraki 
	turlarında heryerde tekrar hata kontrolüne takılır. O yüzden sayacı 
	arttırdıktan sonra hata numasını sıfırlamak gerekir. </p>
	<p> Aşağıdaki örnekte, yine otomatik mailing yapacağız ve kaç adet mailin 
	gitmediğini MsgBox ile bildireceğiz. Örneği özellikle basit tuttum, sadece 
	bu kısma odaklanmanız için. Tabiki yukardaki örneklerle birleştirilerek mail 
	gitmeyen satırlar için renklendirme de yapılabilir.</p>
	<pre class="brush:vb">On Error Resume Next

son=Range("A1").End(xlDown).Row
For i=1 to son
'mail kodları

If Err.Number&lt;&gt;0 The
   hata=hata+1
   Err.Clear
End If   

Next i

MsgBox hata &amp; " adet mail gönderilememiştir"</pre>
</div>
	<h2 class="baslik">Prosedürler arası hata yakalama</h2>
	<div class="konu">
	<p>A1 ve A2 isminde iki prosedürümüz oluğunu düşünelim. A1'in içinde bir yerde A2'yi 
	çağırıyoruz. A2'de bir hata meydana geldiğinde, VBA A2'nin içnde Error 
	Handler var mı diye bakar. Varsa buna girer, yoksa geriye gidip A1'e bakar, A1'de varsa 
	A1'in Error Handler'ına gider.</p>
	<p>Aşağıdaki örnekleri F8 ile çalıştırıp görelim. İlk olarak A2'de Error Handler 
	olmadığında A1'deki Error Handlera geldiğini görelim.</p>
	<pre class="brush:vb">Sub A1()
On Error GoTo hata

MsgBox "selam"
Call A2
Exit Sub

hata:
MsgBox "A1'deki hata mesajı"
End Sub

Sub A2()
MsgBox 1 / 0

End Sub
</pre>
	<p>Şimdi de A2'de varsa A2'de kaldığını.&nbsp;</p>
	<pre class="brush:vb">Sub A1()
On Error GoTo hata

MsgBox "selam"
Call A2
Exit Sub

hata:
MsgBox "A1'deki hata mesajı"
End Sub

Sub A2()
On Error GoTo eh
MsgBox 1 / 0
Exit Sub

eh:
MsgBox "A2'deki hata mesajı"
End Sub
</pre>
	<p> Tabi bu durum<strong> sadece On Error Goto</strong> ifadeleri 
	için geçerlidir, Resume için geçerli değldir.</p>
	<p> Aşağıdaki örnekte sadece A1'de Resume satırı 
	var, dolayısıyla A2'de ilk hatayı alır almaz A1'e dönecek ve "devam" 
	mesajını çıkaracak.</p>
	<pre class="brush:vb">Sub A1()
On Error Resume Next

MsgBox "selam"
Call A2
MsgBox "devam"

End Sub

Sub A2()

MsgBox 1 / 0
MsgBox "selam2"
End Sub</pre>
	<p> Halbuki biz A2'de kaldığı yerden devam 
	etmesini istersek A2 içine de ayrıca <strong>On Error Resume Next</strong> yazmalıyız.</p>
	<pre class="brush:vb">Sub A1()
On Error Resume Next

MsgBox "selam"
Call A2
MsgBox "devam"

End Sub

Sub A2()
On Error Resume Next
MsgBox 1 / 0
MsgBox "selam2"
End Sub</pre>
</div>
	<h2 class="baslik">Hata kaydı tutma</h2>
	<div class="konu">
			<p> Yukardaki bazı örneklerde Logcu isminde bir 
			prosedür görmüştük. Bununla ortaya çıkan hataları detay bilgilerle kayıt 
			altına alabilriz. Hata takibi açısından oldukça faydalı bir yöntemdir. 
			Özellikle ortak kullanılan ve ara sıra hata veren bi dosyanız varsa şiddetle 
			tavsiye ederim. Böylece kim hangi aşamada ne yapmış da hata meydana gelmiş 
			bunu görebilrisiniz.(dosya refreshi sırasında mı hata oluştu, bir 
			dosyaya ulaşmaya çalışırken mi ha oluştu v.s)</p>
			<p> Tabi bunu sadece hata kaydı tutma amacıyla 
			kullanırsanız onun hakkını vermemiş olursunuz. O yüzden Logcu prosedürünü 
			programın kilit noktalarına yerleştirerek tam bir <strong>olay kaydı</strong> 
			tutabilirsiniz. Bu konunun detay anlatımı başka bir sayfada olduğu için burada 
			tekrar örneklendiremeye gerek görmüyorum. 
			<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx#logger">Buradan</a> bakıp incelenemenizi 
			tavsiye ediyorum.</p>
</div>
</asp:Content>
