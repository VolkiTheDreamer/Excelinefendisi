<%@ Page Title='DortTemelNesne Workbook' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server'
 Text='Dört Temel Nesne'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td>
 </tr></table></div>
 <h1>Workbook</h1>
    <h2 class="baslik">Giriş</h2>
<div class="konu">
	<p>Excel uygulamasında o anda açık olan tüm dosyalar Workbooks koleksiyonunu 
	oluşturur ve <span class="keywordler">Workbook</span> nesnesi de bu
	<span class="keywordler">Workbooks</span> koleksiyonunun bir üyesidir.Bu site boyunca bazen Workbook ifadesini bazen dosya ifadesini 
	kullanacağım.</p>
	<p>Bu arada bir de <span class="keywordler">Window</span> nesnesi vardır ki, 
	çoğu durumda Workbook nesnesi ile özdeş kullanılabilir. Window'un farkı 
	şudur, bir dosya açtığınızda bu tek bir workbooktur, ancak bunu birkaç 
	pencere şeklinde gösterebilirsiniz, işte bunların her biri Window 
	nesnesidir. Bu nesneye de yer yer değinip onun özelliklerini de 
	incelyeceğiz. Bu window(pencere) konusunu tam bilmiyorsanız Google'da Excel 
	Window araması yapabilir, ek bilgi alabilirsiniz.</p>
    </div>
<h2 class="baslik">Temeller</h2>
<div class="konu">
	<h3> Erişim</h3>
	<p> Workbooklar bir koleksiyon üyesi oldukları için onlara koleksiyonun 
	<strong>Item</strong> 
	özelliği ve bu özelliğin index numarası ile ulaşabiliriz. Item özelliği 
	koleksiyonların default özelliği olduğu için tüm diğer koleksiyonlarda 
	olduğu gibi bunda da yazılmadan es geçilebilir. Yani <strong>Workbooks.Items(1)</strong> ile 
	<strong>Workbooks(1)</strong> tamamen özdeştir. 
	Koleksiyonlarda indexler 1'den başlar, ilk açılan dosyanın indexi 1'dir.</p>
	<p>Workbooklara index numarasıyla olduğu gibi dosya adı ile de, <strong>Workbooks("Bütçe2016.xlsx")</strong> 
	gibi, ulaşabiliriz. Dosya ismini Window nesnesi ile de kullanabiliriz. 
	Windows("Bütçe2016.xlsx") gibi.</p>


	<h3>Mevcut bir Workbook'u Açma</h3>
	<p>Diskte kayıtlı olan bir dosyayı <span class="keywordler">Open</span> metodu ile açarız. Hemen makro 
	kaydedici ile bir dosya açıyorum ve koduma bakıyorum.</p>

	<pre class="brush:vb">
Sub Macro3()
     Workbooks.Open Filename:="C:\Users\Volkan\Desktop\deneme.xlsx
End Sub</pre>
	<p>NOT:Bunu yazarken Filename ifadesini kullanmadan veya parantez içinde 
	şöyle de yazabilirdim: <strong>Workbooks.Open 
	("C:\Users\Volkan\Desktop\a.xlsx"). </strong>Parametre yazımının detayları 
	için <a href="Temeller_Birazdahaterminoloji.aspx">buraya</a> bakınız. </p>
	<p>Open metodunun birçok parametresi var ama ben burada sadece önemli 
	olduğunu düşündüğüm 2 parametreden bahsetmek istiyorum. Bunlardan ilki <span class="keywordler">ReadOnly</span> parametresi. 
	Şimdi diyelim ki, kullanıcılar için UserForm kullanarak bir arayüz veya bir 
	Add-in oluşturdunuz. Oradaki bir buton da bir dosyayı açacak, ama dosyada 
	ara ara güncellemeler yapmanız gerekiyor, hatta dosyayı schedule programına 
	aldınız ve gece belirli bir saatte açılıp refresh olacak ve kaydedilecek. Bunun için dosyanın kimsede açık olmaması lazım. O yüzden kullanıcılar bu 
	butona tıkladıklarında bunun Readonly açılmasını sağlamak akıllıca olurdu. 
	İşte bu parametre ile bunu yapabiliyoruz, tabi buna True değerini atayarak. </p>
	<p>Peki, diyeceksiniz ki, "Bizim bölümdeki bazı kullanıcılar çok cin veya 
	umarsız, onlara hazırladığım arayüzü/add-ini değil de doğrudan dosyanın 
	bulunduğu klasörden dosyayı açabilirler". Bunun da çözümü var, Dosyanın Workbook_Open event modülüne dosyanın Readonly açılıp açılmadığını kontrol 
	eden bir kod yazarsınız. <span>(Bu detayları yeri geldikçe göreceğiz).
	</span>Kodlar şöyle olacaktır:</p>
	<pre class="brush:vb">
'dosyayı açmaya yarayan kod
Sub dosyaac()
    Workbooks.Open "C:\Users\Volkan\Desktop\deneme.xlsx", ReadOnly:=True
End Sub
'----------------------
'bu dosyanın Workbook_Open modülü
Sub Workbook_Open()
'ilk başta açan kullanıcının sizin dışında bir kullanıcı olup olmadığı kontrolü yapaım
If Environ$("computername") = "bilgisayar adınız" then Exit Sub 'veya username ile de kontrol edilebilir

If Me.Readonly=False Then 
   MsgBox "Lütfen Add-in üzerinden giriş yapınız"
   Me.Close SaveChanges:=False
End If
End Sub
</pre>
	<p>İkinci önemli parametre de <span class="keywordler">Updatelinks'</span>tir. 
	Bu parametre ile, dosya açıldığında içinde başka workbooklara link varsa 
	bunların güncellenip güncellenmeyeceğini belirtmiş oluruz. Eğer bu parametre 
	belirtilmediyse Excel bize sorar. <strong>DisplayAlerts=False</strong> kullanımının 
	baskılayamadığı tek uyarı şekli budur, o yüzden bunu kendi içinde 
	halletmemiz gerekir. <strong>0</strong>, güncelleme yapılmasın; <strong>3</strong>, 
	yapılsın demektir.</p>
	<p>Linklerle ilgili bir başka önemli metod da <span class=" keywordler">LinkSources</span> metodu olup, 
	bir dosyanın başka bir dosyadan(veya MS uygulamasından) link alıp 
	almadığını gösterir. Bu metodun dönüş değeri dizidir. Tüm linkleri bir dizi 
	içinde ayrı elemanlar olarak depolar. Eğer link yoksa dizi boştur ve IsEmpty 
	ile yakalanabilir.</p>
	<pre class="brush:vb">
linkler = ActiveWorkbook.LinkSources(xlExcelLinks) 
If Not IsEmpty(linkler) Then 
 For i = 1 To UBound(aLinks) 
    MsgBox "Link " & i & ":" & Chr(13) & linkler(i) 
 Next i
Else
  MsgBox "Dosya başka dosyalara link içermemektedir."
End If	</pre>
	<h4 id="autoopen">Auto_open makrosu</h4>
	<p>Workbook_Open <span>event handlerına </span>benzer bir de Auto_open 
	prosedürü vardır. Bu, daha 
	çok eskiden 
	kullanılırdı, ancak Excel 2000'den sonra Workbook_Open event makrosu devreye 
	girdiği için buna pek gerek kalmadı, eski makroları destek adına yaşamaya 
	devam ediyor. O yüzden olur da elinizde sizden önce 
	birilerinin yazdığı bir Auto_open kodu varsa ve bunu değiştirmek 
	istemiyorsanız aşağıdaki bilgi önemli olacaktır.</p>
	<p><span>Biz henüz eventlere(olaylara) gelmedik, neden bundan bahsediyorum.
	</span>İşte, VBA içinden çalışan bir Open metodu Auto_open makrosunu tetiklemez, 
	bu makro sadece dosya manuel olarak yani Excel içinden veya Windows 
	Explorerdan açıldığında tetiklenir. Bunun için ayrıca dosyayı açtıktan sonra 
	bir de şu kodu eklemek lazım.</p>
	<pre class="brush:vb">Workbooks.Open "C:\Users\Volkan\Desktop\deneme.xlsx"
ActiveWorkbook.RunAutoMacros xlAutoOpen</pre>
	<p>Bu arada Workbook_Open makrosunda dosyanın nasıl açıldığı önemli 
	değildir. Manuel de açılmış olsa, kod ile de açılmış olsa tetiklenir.</p>
	<h3>Dosya Kapatma</h3>
	<p><span class="keywordler">Close</span> metodu ile dosyalar kapatılır.</p>
	<p>Bu metodun en önemli parametresi kapatırken kaydedip kaydetmemeye yarayan
	<span class="keywordler">Savechanges</span> parametresidir.</p>
	<p>Aşağıdaki kod ile, aktif dosyayı kaydederek kapatıyoruz.</p>
	<pre class="brush:vb">Sub kapama()
   Activeworkbook.Close savechanges:=True
End Sub</pre>
	<p>Open metodunda olduğu gibi, eğer bir Auto_close makrosu varsa, bu makro Workbook.Close 
	metodu ile tetiklenmez, dosya X işaretine basılarak kapatıldığında 
	tetiklenir. Bunun için Open metodunda olduğu gibi şu kod eklenmelidir, 
	ancak bu sefer close işleminden önce eklenmelidr.</p>
	<pre class="brush:vb">ActiveWorkbook.RunAutoMacros xlAutoClose
ActiveWorkbook.Close Savechanges:=False</pre>
	<p>Aşağıdaki örnekte, açık olan tüm dosyaları kapatan ama kapatmadan önce 
	kaydetip kaydedilmeyeceklerini soran bir kod bulunuyor. Tabi Personal.xlsb 
	gibi gizli dosyalarda bu işlemi atlıyoruz, yoksa tüm Excel kapanır. Ayrıca 
	henüz kaydedilmemiş dosyalar için ayrı bir işlem uyguluyoruz.</p>
	<pre class="brush:vb">
Sub tümdosyaları_kapat()
Dim wb As Workbook

cevap = MsgBox("Kapatırken save edeyim mi", vbYesNoCancel)
If cevap = vbYes Then
    A = True
ElseIf cevap = vbNo Then
    A = False
Else
    Exit Sub
End If

For Each wb In Application.Workbooks
    Workbooks(wb.Name).Activate
    If Windows(wb.Name).Visible = True Then
        If InStr(ActiveWorkbook.Name, ".") = 0 Then 'henüz kaydedilmemiş bir dosyaysa
            wb.SaveAs Filename:=Application.DefaultFilePath & "\" & wb.Name & ".xlsx", _
                FileFormat:=Application.DefaultSaveFormat, CreateBackup:=False
                ActiveWindow.Close
        Else
            wb.Close savechanges:=A
        End If
    End If
Next wb
End Sub
</pre>
	<h3>Yeni Dosya Yaratma</h3>
	<p>Yeni bir dosya yaratacağımız zaman <span class="keywordler">Add</span> 
	metodunu kullanırız. Bağımsız bir satırda kullanacağımız gibi aynı anda bir 
	değişkene de atayabiliriz.</p>
	<pre class="brush:vb">Sub yenidosya()
Dim wb as Workbook

Set wb=Workbooks.Add
End Sub</pre>
	<h3>Dosya Kaydetme</h3>
	<p>Basit kaydetme işlemi <span class="keywordler">Save</span> metodu ile, 
	farklı kaydetme işlemi ise <span class="keywordler">SaveAs</span> metodu ile 
	yapılır. Save metodu parametre almaz, dosyayı sadece kaydeder. SaveAs 
	metodunun birkaç parametresi vardır. </p>
	<p>Tüm syntax şöyle: <strong>SaveAs(FileName,&nbsp; FileFormat,&nbsp; 
	Password,&nbsp;WriteResPassword,&nbsp;ReadOnlyRecommended,&nbsp;
	CreateBackup,&nbsp;AccessMode,&nbsp;ConflictResolution,&nbsp;AddToMru,&nbsp;
	TextCodepage,&nbsp;TextVisualLayout,&nbsp;Local)</strong></p>
	<p>FileName ile dosyaya ne isim vereceğimizi belirtiriz, bu isim klasör 
	ismini de içerebilir, klasör belirtilmezse Default kaydetme klasörüne 
	kaydedilir.</p>
	<p>FileFormat ile dosya formatının ne olacağına karar veririz.</p>
	<p>Belli başlı dosya formatları şöyle olup, tüm listeye
	<a href="https://msdn.microsoft.com/en-us/library/office/ff198017.aspx?f=255&amp;MSPPError=-2147217396">
	buradan</a> ulaşabilirsiniz. Parantez içindeki değerler enumaration 
	sabitleridir, kodlarınızda parantez içi de dışı da kullanılabilir.</p>
	
<table class="alterantelitable">
<th>İsim(Değer)</th><th>Açıklama ve uzantı</th>
<tr><td>xlExcel12(50)</td><td>2007 sonrasında binary format, xlsb</td>
<tr><td>xlWorkbookDefault(51)</td><td>2007 sonrasında gelen klasik format, xlsx</td>
<tr><td>xlOpenXMLWorkbookMacroEnabled(52)</td><td>2007 sonrasında gelen makrolu format, xlsm</td>
<tr><td>xlExcel8(56)</td><td>2007 sonrasında eski klasik format, xls</td>
<tr><td>xlWorkbookNormal(-4143)</td><td>2007 öncesinde klasik format, xls</td>
</table>

	<p>Dosya kaydetme işlemlerinde Excel versiyon kontrolünü de yapmak, zorunlu 
	olmasa da, hataları ele alması açısından akıllıca bir yol olacaktır. 
	(Versiyon numaralarına <a href="../Excel/Giris_ExcelinTarihselGelisimi.aspx">
	buradan</a> ulaşabilirsiniz.)</p>
	<p>Aşağıda örnek bir kod bulunmaktadır.</p>
	<pre class="brush:vb">
Sub versiyonkontrol()
Dim wb As Workbook
Dim uzantı As String
Dim frmt As Integer

Set wb = ActiveWorkbook
 
If Val(Application.Version) &lt; 12 Then
'2007 öncesini kullanıyorsunuzudur
   uzantı = ".xls": frmt = -4143
Else
'2007 sonrasını kullanıyorsunuzudur
   cevap = MsgBox("Eski formatta mı kaydetmek istiyorsunuz?", vbYesNo)
   If cevap = vbYes Then
      uzantı = ".xls": frmt = 56 'belki bu dosyayı 2003 kullanan bir alıcıya göndereceksinizidr
   Else
      uzantı = ".xlsx": frmt = 51
   End If
End If
 
wb.SaveAs Filename:="Yenidosya" &amp; uzantı, FileFormat:=frmt
End Sub
</pre>
	<p>
	Bir kaydetme metodu daha 
	vardır, bu metod dosyanın o anda bulunan halini 
	(kendi üzerinde kayıt işlemi yapmadan) başka bir yere kaydetmeye 
	yarayan <span class="keywordler">SaveCopyAs</span> metodudur. Mevcut dosya 
	üzerinde kayıt işlemi yapmadan, ama bu halini de saklaması adına güzel bir 
	metoddur. Yani özetle, SaveAs açık dosyayı o andaki değişiklikleriyle 
	kaydederken, SaveCopyAs, açık dosyayı aynen bırakır, buna herhangi bir 
	kaydetme işlemi uygulamaz, ancak son değişklikleri 
	içeren halini farklı bir isimde kaydeder. Bu da SaveCopyAs'i <strong>yedekleme 
	amacına</strong> hizmet etmesi için harika bir metod yapar.</p>
	<p>
	Mesela aşağıdaki kod Ribbonunuzda kısayol tuşu olarak bile atayabileceğiniz 
	güzel bir koddur.(FileDialog işlemlerini ayrıca göreceğiz)</p>
	<pre class="brush:vb">
Sub anlık_yedekal()
'kodun sağlıklı çalışması için References içinde Microsoft Scripting Runtime kütüphanesi eklenmiş olmalı
'eğer bu kütühane ekli değilse ve şuan eklemeyi tam bilmiyorsanız aşağıdaki satırı diğer iki set cümlesinin altına ekleyin
'Set fso = CreateObject("Scripting.FileSystemObject")

Dim adres As String
Dim isim As String
Dim fd As FileDialog
Dim wb As Workbook
Dim fso As New FileSystemObject

Set wb = ActiveWorkbook
Set fd = Application.FileDialog(msoFileDialogFolderPicker)

fd.ButtonName = "Seçin"
fd.Title = "Kayıt yeri seçin"
fd.InitialFileName = wb.path

If fd.Show = -1 Then
    adres = fd.SelectedItems(1)
    isim = InputBox("Dosya ismi ne olsun", Default:=fso.GetBaseName(wb.Name) & ".... yapmadan önceki yedek")

    wb.SaveCopyAs (adres + "\" + isim + "." + fso.GetExtensionName(wb.Name))
End If
End Sub	</pre>
	<h4>
	Network üzerinde kaydetme işlemi</h4>
	<p>
	Network üzerinde bir yere kayıt yaparken bazen bir bug(hata) nedeniyle 
	kaydetme işlemi ekranda takılı kalmaktadır. Bunun ne zaman olacağı asla 
	kestirilememekte(%3-5 ihtimal diyebilirim) ve Microsoft da malesef bu soruna 
	bir çözüm bulabilmiş değildir. Eğer schedule edilmiş bir iş sırasında 
	başınıza bu gelirse sonrasında çalışması gereken tüm makrolar durmaktadır, 
	bu da gününüzün iyi geçmemesi için yeterli bi gerekçe olabilir.		</p>
	<p>
	Bununla beraber bu konuda bir çözüm bulunmaktadır. Sırayla şu adımlar 
	uygulanmalıdır.</p>
	<ul>
		<li>Dosya önce yerel makinaya kaydedilir</li>
		<li>Sonra bu dosya kapatılır</li>
		<li>Sonra dosya kapalıyken yerel diskten networkteki klasöre kopyalanır</li>
		<li>Yerel diskteki dosya silinir</li>
	</ul>
	<p>Bunun kod örneği aşağıdaki gibidir. </p>
	<pre class="brush:vb">Sub çağıran()
  '.....
  '.....
  hedef=ActiveWorkbook.FullName
  saveas_islemi(hedef)
End Sub

Sub saveas_islemi(ByVal hdf As String)
	Application.DisplayAlerts = False
	ActiveWorkbook.SaveAs Filename:="C:\geçici\geçici.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled 'burda dosya isminin ne olduğu önemli dğeil bile
	Kaynak = ActiveWorkbook.FullName
	
	ActiveWorkbook.Close 'filecopy yapmadan önce dosyayı kapatmak lazım yoksa izin vermiyor
	FileCopy Kaynak, hdf
	Kill Kaynak
	Application.DisplayAlerts = True
End Sub</pre>
	<p>NOT:Eğer bu işlemin Workbook_Open prosedürü içinde yapılması gerekirse 
	ikinci bi geçici dosya daha yaratmak gerekecektir.</p>


</div>	

<h2 class="baslik">İleri Seviye işlemleri</h2>
<div class="konu">

	<h3>Workbook kaydedilmiş mi kontrolü</h3>
	<p>Dosya üzerinde, son açıldığından beri bir değişiklik yapılıp yapılmadığını 
	görmek için <span class="keywordler">Saved</span> özelliği kullanılır. True 
	dönerse, son açılışından beri bir değişiklik yapılmadığını gösterir.</p>
	<p>Bu property, özellikle Excel'i kapatırken işe yarar. Şöyle ki, diyelim ki 
	belli bir saatte Excel'i kapatmak istiyorsunuz, kapatırken de o anda tüm 
	açık dosyaları sanki kayıtlıymış gibi göstermeniz gereken durumlar olabilir. Olayı daha somutlaştırmak adına yine kendi işimden örnek vermek istiyorum.</p>

<p>Normalde işimi yaptığımı PC'm dışında <a href="DortTemelNesne_Application.aspx#OnTime">Application.Ontime</a> örneklerinde bahsettiğim bir de laptop'ım var, orda da scheduled makrolar çalışıyor. Bunların dışında bir üçüncü bilgisayar daha var ki, bu da bir televizyona bağlı. Televizyonda 10 dk'da bir refresh olan birkaç rapor var, oradan tüm departman anlık olarak bu rapor sonuçlarını görüyor.</p>

<p>Bu bilgisayarı her akşam 21:00'de kapanacak, her sabah 08:30'da da açılacak şekilde ayarladım. Kapama işlemini makroyla, sabah açma işlemini Windows Task Scheduler'a yaptırıyorum. Sabah açılır açılmaz, 21:00'de kapatacak makroyu schedule ediyorum. Schedule etme işinin detaylarını <a href="DortTemelNesne_Application.aspx#OnTime">bu sayfada</a> detaylıca göreceğiz. </p>

<p>Bu projeyi 
<a href="DigerUygulamalarlailetisim_VeriTabani-ConnectionListObjectveQueryTable.aspx#hayaletprotokol">
Hayalet Protokol</a> bölümünde detaylıca ele alacağım. Şimdi sadece bizi ilgilendiren kısmının kodlarını inceleyelim.</p>

<p>Ön bir bilgi daha:Task Scheduler Exceli açar açmaz, XLSTART klasöründeki schedule.xlsb dosyası da açılıyor. Bunun Workbook_Open eventinde da birkaç dosyayı açan bir kod var. Açılan ilk dosya "Kredi raporu", aşağıdaki kodlar da bu kredi raporuna aittir.</p>


<pre class="brush:vb">
'Öncelikle Workbook_Open makrosuna bakalım
Private Sub Workbook_Open()
'Öncül kodlar(tanımlamalar vs)

If Environ$("computername") = "A12345" Then 'TV bilgisayarıysa diye bakıyor. Çünkü bu dosyayı düm departmandakiler açabilir. Sadece TVye bağlı bilgisayrda otomatik refresfh devreye girmeli
   'kodlar
    Application.OnTime Now + TimeValue("00:10:00"), Procedure:="TVRefresh"
   'kodlar
End If

Logger "Giris", 0, "Giris yapildi ve 21:00e 'kapat' schedule ediliyor"
Application.OnTime TimeValue("21:00:00"), Procedure:="kapat" &#39;10 dk sonrasına tekrar refreshliyorum

'Diğer kodlar

End Sub

'-------------------------------------------------------------------------
'şimdi de Workbook_BeforeClose'a bakalım
Private Sub Workbook_BeforeClose(Cancel As Boolean)
On Error GoTo hata 'kapatma tuşuna basıp tüm bunlar yapıldıktan sonra hiç iptal edilecek schedule işşemi kalmıyor ama en son "şu dosyaları save edeyim mi" diye bi dialog çıkarıyor ya, buna cancel dediğimde Excelde kalmış oluyoruz, sonra ikinci bi kapat dediğinde iptal edecek bi schedule olmadığı için hata oluyor, o yüzden hata kontrolü koydum

If Hour(Now) < 21 Then
    Application.OnTime TimeValue("21:00:00"), Procedure:="kapat", Schedule:=False '21den önce kapatılırsa kapat'ın schedulını iptal ediyoruz
    Logger "Kapanis", 0, "21den once BeforeClose'a girildi ve SCH iptal edildi"
Else
    Logger "Kapanis", 0, "21de BeforeClose'a girildi ve kapanıyor"
End If
Exit Sub
hata:
If Err.Number = 1004 Then Resume Next
End Sub

'---------------------------------------------------------------------------
'son olarak da kapat makrosuna bakalım, bizim esas ilgilendiğimiz kısım burası
Sub kapat()
On Error GoTo hata

    If Environ$("computername") = "A12345" Then
        '.....
        
        For Each wb In Workbooks
            wb.Saved = True 'işte Saved özelliğini burada kullanıyorum,
        Next wb

        Logger "OtoKapanma", 0, "kapat makrosu ile TV'de kapaniyor"
        Application.Quit 'Burası ile de Excel kapatılıyor
    Else
        Logger "OtoKapanma", 0, "kapat makrosu ile kapaniyor"
        Windows("Anlık Kredi Raporu.xlsm").Close savechanges:=False
    End If
Exit Sub

hata:
Logger "Hata", Err.Number, "Kapat makrosunda hata:" & Err.Description
End Sub

</pre>


<p>Neden böyle yapıyoruz, bunu açıklayayım. </p>


<ul>
<li>Kapat makrosu içine direkt <span class=" keywordler">Application.Quit</span> dersem, Excel bana "şu şu dosyaları kaydedeyim mi" diye sorar ve ekran öylece kalır. Bu istediğim birşey değil.</li>
<li>Neden <span class=" keywordler">Workbook.Close</span> deyip <strong>Savechanges</strong> özelliğine False atamıyorum? Çünkü böyle dersem kodun çalıştığı kredi dosyası kapanırsa(ki ilk o kapanacaktır çünkü ilk açılan dosya o) kalan dosyalar için kapatma işlemi uygulanmaz.</li>
<li>Tüm dosyaları kaydedip neden kapatmıyorum? Çünkü dosyalar readonly açılıyor, kaydedilemez; gerçi readonly açılmasa bile kaydetme işlemini TV bilgisayarında yaptırmazdım.</li>
<li>O yüzden dosyaların <strong>Saved</strong> özelliğine True atayıp, onları sanki kaydedilmiş gibi gösteriyorum ve böylece Excel'in bana soru sormamasını sağlıyorum.</li>
</ul>


	<h3>Açık dosyalar arasında dolaşma</h3>
	<p>Tüm açık dosyalarda dolaşmak ve onlarda işlem yapmak için aşağıdaki basit 
	kodu kullanabilrisiniz.</p>
	<pre class="brush:vb">
Sub wbisimleri()
Dim i As Integer

For i = 1 To Workbooks.Count
	'buraya diğer kodlar yazılır
Next i

End Sub	</pre>
	<p>Bazen Personal.xlsb gibi dosyaları hariç tutarak işlem yapmak isteriz.&nbsp; 
	Örneğin, açık olan tüm dosyalarınızı kapatmak isterseniz yukardaki kodu alıp 
	araya Workbooks(i).Close dersek, bu kodu yerleştirdiğimiz Personal.xlsb de 
	kapandıktan sonra hiç açık dosya kalmayacağı için Excel de kapanmış olur. O 
	yüzden gizli dosyalar hariç bakalım demeliyiz. Ancak Workbook nesnesinin 
	<span class="keywordler">Visible </span>diye bir özelliği yok, bunun yerine Window nesnesinin bu özelliğini 
	kullanacağız. Bu durumda kodumuz aşağıdaki gibi olur.</p>
	<pre class="brush:vb">
Dim wb As Workbook
For Each wb In Workbooks
    If Windows(wb.Name).Visible = True Then
	'diğer kodlar buraya
    End If
Next wb	</pre>
	<p>Açık dosyalar arasında Window nesnesnin metodları aracılığı ile de 
	dolaşabiliriz. Window nesnesi birçok açıdan Workbook'a benzese de, bir 
	workbook içnde birden fazla window olduğu durumlarda biraz farklılaşma 
	yaşanabilir. Pencereler arasında dolaşma <span class="keywordler">
	ActiveWindow.ActivateNext</span> ve <span class="keywordler">
	ActiveWindow.ActivatePrevious</span> metodları aracılığı ile olur, ilki bir 
	sonraki pencereyi(workbooku) aktive ederken ikincisi bir öncekini aktive eder. İki dosya 
	arasında git gel yaptığınız durumlarda oldukça kullanışlıdır. Ancak birden fazla dosyanın açık olduğu durumlarda 
karışıklığa neden olacağı için kullanmanızı tavsiye etmem.</p>
	<h3>ActiveWorkbook &amp; ThisWorkbook &amp; Me</h3>
	<p>Çoğu zaman birbiri yerine kullanılabilecek olan bu terimler arasında 
	küçük farklar bulunmaktadır. <span class="keywordler">ActiveWorkbook</span>, o anda aktif olan dosya iken 
	<span class="keywordler">ThisWorkbook</span>, kodun çalıştığı dosyadır. 
	Mesela, bizim ana kod dosyamız olan Personal.xlsb dosyası gizli bir dosya 
	olduğu için genelde activeworkbook o 
	olmayacaktır, ancak eğer ondaki bir kod çalışıyorsa ThisWorkbook o olacaktır. </p>
	<p>Sayfa gibi bir nesnenin önünde Workbook ifadesi belirtilmediyse Excel 
	bunun ActiveWorkbookun bir sayfası olduğunu düşünür. Bunun bir istisnası 
	vardır: Eğer bir Workbook modülüne(ThisWorkbook) girildiyse ve burada bir 
	Workbook adı belirtilmeden bir nesne mesela bir sayfa adı kullanılıyorsa, bu 
	durumda bu nesne Activeworkbooka ait değil kodun çalıştığı workbooka ait 
	olarak algılanır.</p>
	<p>Başka bir örnek ise şöyle olabilir: Diyelim ki Kredi Format.xlsm diye bir 
	dosyanız var, gece belli bir saatte çalışacak şekilde programlanmış olsun. 
	Bu dosya ilgili saatte açıldığında, Workbook_Open makrosu devreye girsin ve 
	bazı kodları çalıştırsın. Kodların bir bölümünde PersonelBilgileri.xlsx diye 
	bir dosyayı açıp bundan sicil kodlarının yanına personelin ismini getirsin ve eğer kredi dosyasında personelbilgi dosyasında olmayan bir sicil varsa 
	bunu bu dosyaya eklesin. Bu durumda kredi dosyasından copy yapılıp personel 
	dosyasına paste işlemi yapılacağı için son açık görünen dosya yani 
	activeworkbook, personelsicil dosyası olacaktır. O yüzden paste işleminden 
	hemen sonra Activeworkbook.Close denirse, personelsicil dosyası kapatılmaya çalışılır. 
	Thisworkbook.Close denirse, aktif dosya personel dosyası olduğu halde kredi 
	dosyası kapatılmaya çalışılır, zaten o kapanırsa kod da durur, ve personel 
	dosyası açık bir şekilde bekler. </p>
	<p>Bu örneği biraz da aşama aşama görelim, yazım tekrarı olmaması adına 
	ThisWorkbook'un kod boyunca hep Kredi dosyası olacağını söyleyelim ve sadece 
	Activeworkbook'un değişimini izleyelim:</p>
	<ul>
		<li>Kredi dosyası açılır açılmaz Workbook_Open devreye girdi --&gt; 
		Activeworkbook=Kredi</li>
		<li>Personel dosyası açıldı--&gt;Activeworkbook=Personel </li>
		<li>Personelden krediye lookup yapılacak, kredi aktive edilir--&gt; 
		Activeworkbook=Kredi</li>
		<li>#N/A gelenler Personel dosyasına eklenecek, personel aktive edildi 
		ve yeni siciller yapıştırıldı --&gt; Activeworkbook=Personel.</li>
		<li>Önce personel dosyasını kapatmak için Activeworkbook.Close 
		Savechanges:=True denir</li>
		<li>Akabinde ThisWorkbook.Close Savechanges:=True denir ve kredi dosyası 
		da kapanır</li>
	</ul>
	<p>Bir de <span class="keywordler">Me</span> ifadesi vardır. Eğer kodumuz 
	ThisWorkbook modülü içindeyse bu ifade, ThisWorkbook anlamında kullanılabilir. (Ancak 
	Me'nin başka anlama geleceği durumlar da olabilir. Eğer, kodumuz bir sayfa 
	modülü içindeyse mesela Sheet1 modülünüdeyse, Me=Sheet1'dir. Onun dışında 
	Me'nin en çok kullanıldığı yer sanırım UserFormlardır. Bu durumda da Me, formun kendisi 
	olmaktadır. Bunu formlar bölümünde göreceğiz.)</p>
	<h3>Dosya ismi ve adresi/klasörü</h3>
	<p>Dosyanın ismine <span class="keywordler">Name</span> özelliği ile 
	ulaşırız. İsimden kastımız uzantı dahil isimdir. <strong>Ör:"Krediler.xlsx".</strong> 
	Worksheet'in Name özelliğinden farklı olarak bu özellik Readonly'dir(Nesne model tanımınıda "Sets" yok, sadece 
	"Returns" vardır), yani dosya adı dosya 
	açıkken değiştirilemez. Ancak <strong>SaveAs</strong> veya <strong>SaveCopyAs
	</strong>ile dosyaları farklı 
	isimde kaydetme imkanımız olabilir.</p>
	<p>Name özelliği dosyanın sadece adı ve uzantısını verirken, 
	<span class="keywordler">Fullname</span> 
	özelliği tüm path'i de verir. <strong>Ör: "C:\Users\Volkan\Desktop\deneme.xlsx"</strong></p>
	<p>Bazen de dosya adını uzantısı olmadan almak isteriz, mesela uzantısız 
	ismi alıp bu ismin sonuna bir ek ekleyip sonra tekrar uzantıyı eklemek 
	isteyebilirsiniz. Bunun için birkaç yöntem var. Hepsini de göstermeye 
	çalışacağım, başka yöntemler de bulunabilir tabiki, VBA'de bir sonucu 
	elde etmenin binbir türlü şekli var sonuçta.</p>
	<p>
	<strong>1.yöntem:</strong> <span class="keywordler">FileSystemObject</span> objesinin
	<span class="keywordler">GetBaseName</span> özelliğini kullanmak. Bu kanımca 
	en ideal yöntemdir, zira her versiyon ve uzantıda işe yarar. Bunu yukarda 
	anlık_yedek al makrosunda kullanmıştık.</p>
	<pre class="brush:vb">CreateObject("Scripting.FileSystemObject").GetBaseName(ActiveWorkbook.name)</pre>
	<p><strong>2.yöntem:</strong> Uzantının ne olduğunu biliyorsak, bunu "" yani sıfır uzunluklu 
	metinle replace ederiz.</p>
	<pre class="brush:vb">isim = Replace(ActiveWorkbook.Name, ".xlsx", "") </pre>
	<p><strong>3.yöntem:</strong> Dosya isminde nokta işareti olmadığından eminsek, 
	noktanın yerini bulup oraya kadar olan ismi almaktır.(Dosya isminde nokta 
	varsa bu yöntem işe yaramaz)</p>
	<pre class="brush:vb">
isim=Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
</pre>
	<p><strong>4.yöntem:</strong> Uzantı kadar karakteri hariç tutup soldan yeterli sayıda 
	karakter seçmek. Bu örnekte bir de versiyon kontrolü yapılabilir.</p>
	<pre class="brush:vb">
Dim i As Integer
If Application.Version = 12 Then 'Excel 2007 ve sonrasıysa
	If ActiveWorkbook.FileFormat=56 then 'xls uzantılıysa
		i=4
	Else
		i=5
	End If
Else 
	i = 4 
End If
   uzantısızisim= Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - i) 
</pre>
	<p>Ve isterseniz bütün bunları bir fonksiyona da atayabilirsiniz.</p>
	<pre class="brush:vb">
Function uzantısızisim() As String
Dim i As Integer 
If Application.Version &gt;= 12 Then
	If ActiveWorkbook.FileFormat=56 then
		i=4
	Else
		i=5
	End If
Else 
	i = 4 
End If 
uzantısızisim= Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - i) 
End Function 
</pre>
<p>
Dosyanın bulunduğu klasör ismine <span class="keywordler">Path </span>özelliği, uzantısına da yine 
	isminde olduğu gibi birkaç yolla ulaşılabilir, ben örnek olarak sadece <span class="keywordler">FileSystemObject</span> objesinin <span class="keywordler">GetExtensionName</span> özelliğini söyleyeyim, 
	diğer yolları siz düşünün.</p>
	<p><a href="Dosyaislemleri_DosyaveKlasorerisimi.aspx">FileSystemObject</a> nesnesine ayrıca bakacağımız için burada daha fazla 
	detaya girmiyorum.</p>
	<h3>Workbook Protection</h3>
	<p>Excel'in Review menüsünden konulup kaldırılan şifre işlemleri elebetteki 
	VBA ile de yapılabilmektedir. </p>
	<p>Mesela genel kullanıma açık olan bir dosyanızda otomatik refresh 
	işlemleri olduğunu düşünelim. Dosya gece çalışıp içindeki sorgular refresh 
	edilecek ve bazı copy paste işlemleri olacak diyelim. Dosya genel kullanıma 
	açık olduğu ve kimsenin gizli sayfalardaki veriye ulaşmasını istemediğiniz 
	için dosyayı korumaya almış olabilirsiniz. Ancak gece dosya otomatik 
	açıldığında işlemlerin öncesinde korumayı kaldırmalı, işler bitince tekrar 
	korumayı aktive etmelisiniz. İşte kodumuz:</p>
	<pre class="brush:vb">Private Sub Workbook_Open()

ThisWorkbook.UnProtect (1234)
'işlemler yapılır

ThisWorkbook.Protect (1234)
ThisWorkbook.Close Savechanges:=True

End Sub</pre>
	<p>Bir dosyada koruma olup olmadığını görmek için de <span><strong>
	ProtectStructure</strong> </span>özelliği kullanılır. 
	Aşağıdaki örnekte belirli bir grup dosyada toplu connection şifre 
	değişikliği gerçekleşecek ancak öncesinde her dosya için protection kontrolü 
	yapılıyor.</p>
	<pre class="brush:vb">
'Belirli dosyaları files isminde bir collectiona atadım ve bu collection içinde dolanarak çeşitli işlemler yapacağım
For Each file In files
    Workbooks.Open Filename:=file
    Set wb = ActiveWorkbook
    If wb.ProtectStructure = True Then
        protectlimi = True
        wb.Unprotect (1234) 
    End If
        
    For Each cn In wb.Connections
        If InStr(cn.ODBCConnection.Connection, eski) > 0 Then
            cn.ODBCConnection.Connection = Replace(cn.ODBCConnection.Connection, eski, yeni)
        End If
    Next cn

    'protection olanlarda tekrar koyalım
    If protectlimi = True Then wb.Protect (1234)
    protectlimi = False 'tekrar false yapıyorum ki, döngüye tekrar girdiğinde True olarak gitmesin    
Next file	</pre>
	<h3>Diğer</h3>
	<p>Workbook'un başka özellik ve metodları da bulunmaktadır. Ben burada 
	önemli olduğunu düşündüklerimi vermeye çalıştım. Sizler diğer üyeleri
	araştırabilir ve kendiniz deneyebilirsiniz.</p>
	<p>Bununla beraber <span class=" keywordler">RefreshAll</span> ve <span class=" keywordler">Connection</span> gibi üyeleri ise <a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">VeriTabanı</a> 
	işlemlerinde ayrıca ve detaylıca ele alacağım için burada değinmedim.</p>
	
	</div>
</asp:Content>
