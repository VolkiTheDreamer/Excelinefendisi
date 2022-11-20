<%@ Page Title='Olaylar WorkbookOlaylarievent' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Olaylar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Workbook Olayları(Eventleri)</h1>
    <h2 class="baslik">Giriş</h2>
    <div class="konu">
	<p>Sözkonusu Workbook nesnesi olduğunda, en temel eventler "açılma, kapanma, kaydolma" 
	üçlüsüdür. bunun dışında pek tabiki başka olaylar da bulunmaktadır. Ben 
	burada temelini vermeye çalışacağım. Gerisi kurcalama iştahınıza kalmış.</p>
	<h3> 	Dosya açılması</h3>
	<p> 	Bir dosya açıldığında <span class="keywordler">Workbook_Open</span> olayı devreye girer. Bu event 
	ile kullanıcıya çeşitli mesajlar verebileceğiniz gibi,
	<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx#logger">log kaydı</a> 
	oluşturma, veritabanı işlemleri gibi başka birçok işlem yapabilirsiniz.</p>
	<pre class="brush:vb">Private Sub Workbook_Open()
  MsgBox "x dosyasına hoşgeldiniz. Dosyayı kullanırken şunlara dikkat edin" &amp; vbCrLf &amp; "falan filan" &amp; vbCrLf &amp; "falan filan"
  Logger Me.Name, Environ("username"), Date 'log kaydı tutuyoruz, kim ne zaman girmiş diye
End Sub</pre>
	<h4> 	Olay sırası</h4>
	<p> 	Bazı olaylar meydana geldiğinde birden çok event tetiklenebilir. Mesela 
	bir dosya açıldığında sadece <strong>Open</strong> eventi devreye girmez, aynı 
	zamanda <strong>Activate</strong> eventi de devreye girer. Bunlardaki sıra önce 
	<span class="keywordler">Workbook_Open </span>sonra <span class="keywordler">
	Workbook_Activate </span>olacak şekildedir.</p>
	<h4> 	Auto_open</h4>
	<p> 	
	Workbook_Open makrosuna benzer bir de Auto_open makrosu vardır. Onla ilgili 
	detaya <a href="DortTemelNesne_Workbook.aspx#autoopen">buradan</a> 
	ulaşabilirsiniz.</p>
	<h3> 	Kapanma ve Kaydetme</h3>
	<p> 	Bir dosya kapanırken de önce <span class="keywordler">Workbook_BeforeClose</span>, sonra 
	<span class="keywordler">Workbook_Deactivate<strong> </strong></span>eventi 
	devreye girer. Kapatma sırasında kaydetme de olacaksa süreç şöyle olur:</p>
	<pre class="brush:vb">Workbook_BeforeClose
Workbook_BeforeSave
Workbook_AfterSave
Workbook_Deactivate</pre>
	<pre class="brush:vb">Private Sub Workbook_BeforeClose(Cancel As Boolean)

  MsgBox "Dosyadan ayrılıyorsunuz. Falan filan yapmayı unutmayın."

End Sub</pre>
	<h4> 	İptal parametresi</h4>
	<p> 	
	Kapanma ve Kaydetme olayları öncesinde bazen bu olayın geçekleşip 
	gerçekleşmeyeceğini de kontrol altına almak isteyebilriiz. Bunu da 
	varsayılan değeri False olan <strong>Cancel</strong> parametresine kod içinde True değeri 
	atayarak yaparız. </p>
	<p> 	
	Mesela kaydetme işlemini, sadece belli bir şifreyi bilen kişilere yaptırmak 
	isteyebilirsiniz. Bunun için aşağıdaki gibi bir kod yazılabilir.</p>
	<pre class="brush:vb">
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim şifre As String
şifre = InputBox("Kaydetmek için yetkili şifresini girin")
If şifre <> "1234" Then
    MsgBox "Dosyayı kaydetmeye yetkili değilsiniz."
    Cancel = True
Else
    MsgBox "dosya başarılı bir şeklide kaydedildi" 'belki bir de hata kontrolü konulabilir buraya
End If
End Sub	</pre>
	<h4> 	Kapanma/Deactive olma durumuna göre farklılaşan mesaj gösterme örneği</h4>
	<p> 	Aşağıdaki örnekte ise dosya kapanırken farklı, deaktive olurken farklı 
	bir mesaj verme örneği var. </p>
<pre class="brush:vb">
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Range("Z1000").Value = 1 'hangi hücre müsaitse
    Me.Saved = True
End Sub

Private Sub Workbook_Deactivate()
    If Range("Z1000").Value = 1 Then
        MsgBox "çıkma mesajı"
    Else
        MsgBox "deaktive mesajı"
    End If
End Sub</pre>
<p>
Burdaki süreç şöyle işleyecek. Diyelim ki kullanıcı dosya açıldıktan bir süre 
sonra başka bir dosyaya geçmek istedi, deactivate olayı ilk kez devreye girer, 
Z1000 hücresinde 1 yazıyor mu diye bakar, yazmadığı için sadece "deaktive mesajını" 
verir. Sonra diyelim ki yine aktif oldu, bi süre sonra tekrar başka dosyaya geçtiğinde bu süreç 
aynen tekrar eder. Ne zaman ki dosyadan çıkmak isterse, önce BeforeClose devreye girer ve 
Z1000'e 1 değerini atar, sora Deactivate olayı devreye girer, Z1000=1 olduğu için 
sadece çıkma mesajı verilir. Burada BeforeClose içine hiçbir mesaj 
yazmadık, böyle yapsaydık hem "Çıkma mesajı" hem de "Deactive mesajı" verilmiş 
olurdu, bu da arka arkaya 2 farklı mesaj kutusu demek olurdu ki pek şık bir 
durum olmazdı. Burada tabi ben sadece mesaj farklılaşması yaptım, sizin 
ihtiyacınıza göre farklı işlemler de yaptırılabilir.</p>
	<h4> 	Auto_close</h4>
	<p> 	
	Auto_open için yazılan bilgilerin benzeri aynı mantıkta Auto_close için de 
	geçerlidir. </p>
	<h4> 	
	Kapanmadan önce sayfayı default ayarlara getirmek</h4>
	<p> 	
	Başlangıç ayarları özellikle belli değerlere ayarlanmış bir dosyanız olsun. 
	Kullanıcılar, bu dosyada çeşitli oynamalar yaptıktan sonra kapatacaklar 
	diyelim. Dosyayı ilk ayarlarına getirmek için BeforeClose olayı 
	kullanılabilir.</p>
	<pre class="brush:vb">Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Application.EnableEvents = False
  Range("B1") = vbNullString
  Range("B2") = 400
  Range("C21") = "Toplam"
  '....
  Application.EnableEvents = True
End Sub</pre>
	<p>Application.EnableEvents kullanmamızın sebebi, sayfada Worksheet_Change 
	eventi varsa bu tetiklenmesin diyedir. Bu konuda daha detay bilgi için
	<a href="Olaylar_WorksheetOlaylari.aspx">buraya</a> bakabilirsiniz.</p>
	<h3>Diğer olaylar</h3>
	<h4>Yeni sayfa eklenmesi</h4>
	<p>Yeni bir sayfa yaratıldığında devreye girer. Bazen birilerine 
	gönderdiğiniz bir dosyada yeni sayfa yaratılmasını istemiyor ve 
	kullanıcıların mevcut sayfa(lar) üzerinden çalışmasını istiyorsanız bu 
	eventi aşağıdaki gibi kullanabilirsiniz.</p>
	<pre class="brush:vb">Private Sub Workbook_NewSheet(ByVal Sh As Object)
	Application.ScreenUpdating = False
	MsgBox "yeni sayfa yaratamazsın"
	Application.DisplayAlerts = False
	Sh.Delete
	Application.DisplayAlerts = True
	Application.ScreenUpdating = True
End Sub</pre>
	<p>NOT:Screenupdating'i kullanma sebebi, kullanıcının geçici de olsa 
	sayfanın yaratıldığını görmesini engellemek içindir.</p>

        <p>Başka olaylar da bulunuyor, ancak bunları araştırmayı size bırakıyorum.</p>
        </div>

    
        <h2 class="baslik">Detaylar</h2>
    <div class="konu">
	<h3>
	Event tetiklenmesini bastırmak(Geçici olarak durdurmak)</h3>
	<p>
	Her ne kadar workbook eventlerinde kısırdöngüye giren event durumu pek 
	rastlanılan bir durum olmasa da, yine de belirli şartlarda eventleri geçici 
	olarak durdurmak, o özel durum geçince tekrar aktive olmalarını sağlamak 
	isteyebilirsiniz.</p>
	<p>
	Ben mesela QuickAccessToolbar üzerine eventleri toggle buton mantığı ile bir 
	aktif bir pasif yapan bir düğme koydum. Zira sıklıkla schedule edilmiş dosyaları 
	açıp içlerinde günelleme yapma ihtiyacım oluyor. Normalde bunlar açıldıklarında 
	Workbook_Open makrolarının otomatikman çalışmaları gerekiyor, ama güncelleme yapacaksam 
	bunların çalışmasını 
	engellemek için geçici olarak eventleri pasifleştiriyorum, işimi bitirince 
	tekrar aktifleştiriyorum.
	</p>
	<p>
	İlgili düğmeye atadığım kod şöyle:</p>
	<pre class="brush:vb">
Sub toggle_event()
If Application.EnableEvents = True Then
    Application.EnableEvents = False
    Application.StatusBar = "EnableEvents=False****************************************EnableEvents=False******************************************EnableEvents=False****************************************EnableEvents=False****************************************"
Else
    Application.EnableEvents = True
    Application.StatusBar = "EnableEvents=True*****************************************EnableEvents=True*******************************************EnableEvents=True*******************************************EnableEvents=True******************************************"
End If
End Sub</pre>
	<h3>
	<span style="font-family: 'Trebuchet MS'">
	<a href="https://eksisozluk.com/voltrani-olusturmak--395838">
	<span style="color: #800000">Voltranı</span></a></span> oluşturmak</h3>
	<p>
	Workbook_Open eventini, <a href="DortTemelNesne_Application.aspx#OnTime">
	Application.OnTime</a>,
	<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx#logger">Logger fonksiyonu </a>ve
	<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">Veritabanı</a>(Connection 
	veya ListObject) refresh işlemleriyle birlikte kullandığınızda Voltranı 
	oluşturmuş olursunuz. Bu dörtlü sizi ve kurumunuzu inanılmaz bir verimlilik 
	sürecine sokar. Zira bu şekilde istediğiniz sayıda raporu(yeterli cihazın 
	olması varsayımı altında) otomatiğe bağlayabilirsiniz. Diğer üçüyle ilgili detaylara ilgili linklerde 
	ulaşabilirsiniz, ben buraya sadece küçük bir örnek koymak istiyorum.</p>
	<p>
	Application.OnTime ile schedule edilmiş bir makro düşünün. Bu, Voltranın 
	1.üyesinin devreye girdiği andır. Bu makro çalıştığında(ve çalışırken kod 
	boyunca önemli yerlerde) Log kaydı oluşturacaktır:2.üye de tamam. Bu makro 
	ile belirli bir excel dosya açılmaktadır ve devreye bu dosyanın 
	Workbook_Open eventi girer, işte 3.üye. Son olarak da bu makro içinden 
	çeşitli veritabanından data çekme işlemi yapılarak 4.üye de devreye sokulmuş 
	olur. Son 2 üye aşağıda gösterilmiştir.</p>
	<pre class="brush:vb">
Private Sub Workbook_Open()
'Önceki kodlar
'Çeşitli ön kontroller(kullanıcı, pc adı v.s gbi) burada yapılabilir

For Each cn In ActiveWorkbook.Connections
    cn.ODBCConnection.BackgroundQuery = False
    cn.Refresh
Next

'Sonraki kodlar
'dosyayı kaydetme ve alıcılara mail gönderme kodları burada yer alabilir

End Sub</pre>
	<h3>
	Kısıtlama uygulama ve 
	Veri Güvenliği sağlama</h3>
	<p>
	Diyelim ki gizli bilgiler içeren bir dosyayı birine gönderdiniz, veya genel 
	kullanım için bir add-in yazdınız ve bu add-in gizli/hassas bilgiler içeren 
	bir veritabanı doyasından belli bir sorgu çalıştırıp getiriyor. Herneyse, bu 
	bilgilerin kullanıcılar tarafından yazdırılmasını, copy/paste ile başka 
	dosyaya kopyalanıp orada da yazdırılmasını engellemek istiyorsunuz ve 
	herhangi bir şekilde mail olarak gönderilmesini de engellemek istiyorsunuz. 
	Aşağıda, bunlara ait çözümleri bulabilirsiniz.</p>
	<h4>
	Sayfanın yazdırılmasını engelleme</h4>
	<p>
	Bunu aşağıdaki basit kod ile yapmak mümkündür.</p>
	<pre class="brush:vb">
Private Sub Workbook_BeforePrint(Cancel As Boolean)

  Cancel = True
  MsgBox "Bu dosyayı print almanıza izin verilmemiştir."

End Sub	</pre>
	<p>
	Eğer belli bir kişinin/kişilerin print alma yetkisi olsun isterseniz belli 
	bir sicili/sicilleri kontrol edebilirsiniz. Aşağıda tek sicil numarasının 
	kontrol örneği var, siz isterseniz bir diziye(veya collection) tüm 
	yetkilileri atıp o dizi içinde var mı diye de kontrol edebilirsiniz.(Username'in 
	sicil döndürdüğü varsayımı ile hareket edilmiştir. Sizin kurumda isim de 
	kullanılıyor olabilir, o zaman daha farklı bir yöntem denemeniz gerekebilir. 
	Sadece özel bir şifreyi bilenlerin kaydetmeye yetkisi olması gibi)</p>
	<pre class="brush:vb">
Private Sub Workbook_BeforePrint(Cancel As Boolean)

If Environ("username")&lt;&gt;12345 Then
  Cancel = True
  MsgBox "Bu dosyayı print almanıza izin verilmemiştir."
End If

End Sub	</pre>
	<h4 id="cutcopyengel">Cut/Copy engelleme</h4>
	<p>
	Cut/Copy işlemlerini engellemek için biraz daha uğraşmamız gerekiyor. 
	Öncelikle yasaklama işlemini hem Workboook_Open hem de Workboook_Activate 
	eventlerinde kullanmamız gerekiyor. Ayrıca kullanıcı başka bir dosyayı 
	açtığında o dosyada yasakların kalkıp normale dönmesi gerektiği için de 
	işlemlerin tersinin Workboook_Deactivate eventinde yapılması gerekiyor.</p>
	<p>
	Aşağıdaki örnek kullanıcıya bir de mesaj veriyoruz. Mesaj vermek 
	istemiyorsak aşağıda "cutcopyengel" olan herşeyi "" olarak değiştirmeniz 
	yeterlidir.</p>
	<pre class="brush:vb">
Private Sub Workbook_Open()

With Application
    .CutCopyMode = False
    .OnKey "^c", "cutcopyengel" 'Ctrl+C ile copy
    .OnKey "^x", "cutcopyengel" 'Ctrl+X ile cut
    .OnKey "^{INSERT}", "cutcopyengel" 'Ctrl+Insert ile copy
    .OnKey "+{DELETE}", "cutcopyengel" 'Shift+Delete ile cut
    .OnKey "+{DEL}", "cutcopyengel" 'Shift+Del ile cut
    'mousela taşıma iptali
    .CellDragAndDrop = False
    'sağ tıklama ile engel
    .CommandBars("Cell").Controls(1).Enabled = False 'cut
    .CommandBars("Cell").Controls(2).Enabled = False 'copy
End With

End Sub


Private Sub Workbook_Activate()
With Application
    .CutCopyMode = False
    .OnKey "^c", "cutcopyengel" 'Ctrl+C ile copy
    .OnKey "^x", "cutcopyengel" 'Ctrl+X ile cut
    .OnKey "^{INSERT}", "cutcopyengel" 'Ctrl+Insert ile copy
    .OnKey "+{DELETE}", "cutcopyengel" 'Shift+Delete ile cut
    .OnKey "+{DEL}", "cutcopyengel" 'Shift+Del ile cut
    'mousela taşıma iptali
    .CellDragAndDrop = False
    'sağ tıklama ile engel
    .CommandBars("Cell").Controls(1).Enabled = False 'cut
    .CommandBars("Cell").Controls(2).Enabled = False 'copy
End With
End Sub


'Deactivate eventinde OnKey'in ikinci parametresini boş bırakarak 
'işlemi tersine çeviriyoruz
Private Sub Workbook_Deactivate()
With Application
    .CutCopyMode = True
    .CellDragAndDrop = True
    .OnKey "^c"
    .OnKey "^x"
    .OnKey "^{INSERT}"
    .OnKey "+{DELETE}"
    .OnKey "+{DEL}"
    .CommandBars("Cell").Controls(1).Enabled = True
    .CommandBars("Cell").Controls(2).Enabled = True
End With
End Sub

'Bu da ayrı bir modül içindeki kodumuz
Sub cutcopyengel()
   MsgBox "Bu dosyada cut/copy yapamazsınız"
End Sub	</pre>
	<h4>
	Mail göndermeyi engelleme</h4>
	<p>
	Dosyayı mail ile göndermeyi engellemek için iki iş yapmak lazım. </p>
	<ol>
		<li>Excel içinden File menüsündeki Send As Attachment düğmesini 
		kullanmayı engellemek</li>
		<li>Dosyayı kaydedip Windows Explorer üzerinden veya Outlook ile 
		göndermeyi engellemek</li>
	</ol>
	<p>
	1.yöntemde Ribbonu özelleştirmek gerekiyor, aynı zamanda ilgili mail 
	gönderme butonu QAT(QuickAccessToolbar) üzerinde de olabilir diye QAT'ı 
	göstermeyi de engellemek lazım. Bunları 
	<a href="Ileriseviyekonular_MenuveRibbon.aspx">ileri yöntemler </a>sasyfasında 
	göreceğiz.</p>
	<p>
	2.yöntemde ise dosyanın kaydolmasını engellemeliyiz.</p>
	<pre class="brush:vb">Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Cancel = True
  MsgBox "Bu dosyayı kaydedemezsiniz"
End Sub</pre>
	<p>Tabi yine belli kişiler için kaydetme izni vermek isterseniz 
	Environ("username") özelliğini kullanabilirsiniz.</p>
	<h3>
	Tüm sayfaları ilgilendiren olaylar</h3>
	<p>
	Her ne kadar tüm sayfaları ilgilendiren olaylar Workbook eventleri olsa da 
	onları burada değil Worksheet olayları içinde anlatmayı uygun buldum. Bu 
	sayfaya alttaki İleri butonuna tıklayarak ulaşabilirsiniz.</p>
</div>
</asp:Content>
