<%@ Page Title='Olaylar WorksheetOlaylarievent' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Olaylar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Worksheet Olayları(Eventleri)</h1>
    	<h2 class="baslik">Giriş</h2>
	<div class="konu">
	<p>Bir workbook'un sayfalarındaki çeşitli olaylara tepki vermek adına devreye giren olaylar Worksheet olayları olarak 
	adlandırılır. Bunları da yine Worbook olaylarını 
	seçer gibi seçip içlerini doldurmaya başlayabilirsinz. İlgili combobox 
	seçildiğinde aşağıdaki gibi eventlerin bir kısmı görünür.</p>
	<p><img src="../../images/eventws1.jpg"></p>
	<p>Bunlardan en sık kullanacaklarımız;</p>
	<ul>
		<li>Change</li>
		<li>SheetChange</li>
		<li>BeforeDoubleClick</li>
		<li>Activate/Deactivate</li>
		<li>Calculate</li>
	</ul>
	<p>eventleridir.</p>
	<p>Pivot tablolarla ilgili olanlar da önemli olup bunlara
	<a href="Ileriseviyekonular_PivotTableveSlicer.aspx">Pivot İşlemleri </a>konusunda değineceğiz. 
	Şimdi sırayla önemli eventlere bakalım.</p>
</div>

	<h2 class="baslik">Temel olaylar</h2>
	<div class="konu">
	<h3>Worksheet_Change Event</h3>
	<p>Kuşkusuz en önemli sayfa olayı sayfada bir hücrenin değişimiyle meydana 
	gelen <strong>Change </strong>olayıdır. (Bu 
	eventin adını AfterChange gibi düşünmeniz yerinde olur. Zira olay, hücre içi 
	değiştikten sonra meydana gelir. Microsoft geliştiricileri olayın adını keşke böyle 
	yapsalarmış. Ne de olsa After ve Before ile başlayan bir sürü event var.) 
	Syntax'ı aşağıdaki gibidir. </p>
	<pre class="brush:vb">Private Sub Worksheet_Change(ByVal Target As Range)

End Sub</pre>
		<p>Küçük bir örnek yapalım. Bu örnekte, her değişim oldukça sayfanın 
		rengi değişsin. Bu örneği alıp istediğiniz bir dosyanın Sheet1 modülüne 
		yapıştırın ve sonra gidip sayfada rasgele hücrelere birşey girin. Her 
		Enter'a basışınızda sayfa rengi değişecektir.</p>
	<pre class="brush:vb">Private Sub Worksheet_Change(ByVal Target As Range)
  x = WorksheetFunction.RandBetween(1, 1000000)
  ActiveSheet.Cells.Interior.Color = x
End Sub</pre>
		<h4>
		Tetikleyiciler ve özel hususlar</h4>
		<p>
		Change olayı kullanıcının 
		manuel bir işlemi sonucunda tetiklenebileceği gibi bir 
		makro kodu sonucunda da tetiklenebilir.</p>
		<p>
		Bazı özel durumlar da vardır:</p>
		<ul>
			<li>Manuel hesaplama durumundan otomatik hesaplama durumuna 
			geçildiğinde de hücrelerin içi değişir ama bu durum Change olayını 
			tetiklemez. Yine de yeni duruma göre içerik kontrolü yapacaksanız bu 
			sefer <a href="#calculate">Calculate</a><strong> </strong>olayını kullanmanız gerekir. </li>
			<li>Bir hücrenin içini silmek de değişiklik olduğu için <strong>Change
			</strong>olayı 
			tetiklenir.</li>
			<li>Merge butonu ile hücre birleştirmek tetiklemez.</li>
			<li>Bir alanı sıralamak tetiklemez</li>
			<li>Goal Seek kullanarak bir hücrenin değişimi tetiklemez</li>
		</ul>
	<h4>Target Parametresi</h4>
	<p>Target parametresi, belli bir hücrenin içeriğini değişip 
	değişmediği öğrenmek amacıyla kullanılabileceği gibi ilgili hedefin tek bir hücre mi yoksa bir range mi olduğunu belirlemek 
	için de kullanılabilir. Aslında Range nesnesinin tüm özelliklerini kontrol 
	etmek için kullanılabilir.</p>
		<pre class="brush:vb">
If Target.Address="$A$1" Then 'bu bir adres kontrolüdür
If Target.Cells.Count=1 Then 'bu da tek bir hücre mi yoksa bir range mi kontrolüdür</pre>
	<p>Target'ın belirli bir aralıkta olup olmadığını öğrenmek için özel bir 
	kullanım şekli vardır: <span class="keywordler">If Not Intersect(Target, 
	Range("..")) Is Nothing Then</span></p>
		<p>
		Aşağıdaki örnekte değişen hücrenin C3 veya C4'te olması beklenmektedir. 
		Bununla ilgili daha detaylı örnek Çeşitli Örnekler bölmünde 
		yapılacaktır.</p>
		<pre class="brush:vb">
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("C3:C4")) Is Nothing Then
        'ana kod bloğu
    Else
        MsgBox "Yanlış yerden seçim yapıyorsunuz, sadece C3 ve C4 hücrelerini kullanınız"
    End If
End Sub</pre>
		<h4>
		Aynı hücredeki değişimlerde bir önceki değeri elde etme</h4>
		<p>
		Değişen hücrenin bir önceki değerini elde etmek istiyorsak <strong>
		<a href="Temeller_DegiskenlerveVeriTipleri.aspx#static">Statik</a></strong> değişken kullanırız.		</p>
		<pre class="brush:vb">
Private Sub Worksheet_Change(ByVal Target As Range)

Static öncekiDeğer As String
Static öncekiAdres As String

If öncekiDeğer &lt;&gt; "" And öncekiAdres = Target.Address Then
   MsgBox "Önceki:" &amp; öncekiDeğer
End If

öncekiDeğer = Target.Value
öncekiAdres = Target.Address

MsgBox "yenisi:" &amp; Target.Value

End Sub</pre>
		<p>
		Bu örnekte statik değişkenlerimiz ilk başta boş olacaktır, zira henüz 
		"öncesi" yoktur. İlk işlemden sonra önceki statik değişkenler dolmaya 
		başlayacaktır. Akabinde, yeni hücre ile öncekinin aynı olup olmadığı 
		kontrol edilir.</p>
	<h3>Worksheet_SelectionChange</h3>
	<p>Seçili hücre her değiştiğinde bu event oluşur. </p>
	<pre class="brush:vb">Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub</pre>
		<p>Target, seçilen hücreyi gösterir.</p>
		<p>Aşağıdaki örnekte, seçilen hücre pencerenin sol üst köşesindeki ilk 
		hücre olacak şekilde ayarlanır.</p>
<pre class="brush:vb">
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    With ActiveWindow
        .ScrollRow = Target.Row
        .ScrollColumn = Target.Column
    End With  
End Sub</pre>
		<h4>Önceki seçimi elde etme</h4>
		<p>Seçimden bir önceki 
	hücreye de ihtiyacımız olacaksa <strong>Statik</strong> bir değişken kullanırız. 
		İlk 
		seçimde çalışmaz, sonrakilerde çalışır, çünkü ilk seçimde henüz "öncesi" 
		yoktur.</p>
	<pre class="brush:vb">Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Static öncekiRange As String

If öncekiRange &lt;&gt; "" Then
  MsgBox "önceki:" &amp; Range(öncekiRange).Address
End If

öncekiRange = Target.Address

MsgBox "yenisi:" &amp; Target.Address
End Sub</pre>
		<p>Daha farklı bir örnek ise, önceki hücre ile yeni hücre arasındaki 
		alanı kırmızıya boyamak olabilir. "Ne işimize yarayacak" diye sormayın, 
		bu haliyle bir işinize yaramaz, ama farklı bir konuda size fikir 
		verebilir.</p>
	<pre class="brush:vb">
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Static öncekiRange As Range

If Not öncekiRange Is Nothing Then
    Range(öncekiRange, Target).Interior.Color = vbRed
End If

Set öncekiRange = Target

End Sub
</pre>
	<h3>
	Worksheet_BeforeDoubleClick</h3>
	<p>
	Bir hücreye çift tıklandığında bu olay olur ve Exceli'n o anda nasıl 
	davranmasını istiyorsak bu prosedüre bunları yazarız. Syntaxı aşağıdaki 
	gibidir.</p>
		<pre class="brush:vb">
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

End Sub</pre>
		<p>
		Target'ı 
		şimdiye kadar öğrenmiş olmalısınız; kullanım mantığı yine 
		yukardakilerle aynı. Cancel parametresine ise True değerini atayarak 
		eylemi iptal edebiliriz, yani Excele çift tıklama olmamış gibi 
		davrandırtabiliriz.</p>
		<p>
		En sık kullandığım caselerden birisi, toplanmış verileri tutan bir listede 
		ilgili hücreye çift tıklama sonucunda o grubun alt detayını gösteren verilerin uygun 
		miktarda satır açılarak araya eklenmesi; aynı hücreye tekrar çift tıklanması 
		durumunda ise bu kayıtların animasyonlu bir şekilde silinip(sanki bu 
		sitede bordo arkaplanlı başlıklara tıklandığında yavaşça katlanmasını 
		sağlayan Jquery kodlarına benzer) listenin ilk hale gelmesidir. Böyle bir örnek 
		kullanımı ADO içermesi sebebiyle bu sayfada vermeyip bunları veritabanı 
		uygulamaları bölümünde ele alıyor olacağım. İlgili örneğe
		<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx#doubleclickado">buradan</a> ulaşabilirsiniz. 
		Aynı örneği ilgili veriyi aynı sayfada gizlenmiş bir şekilde dururken 
		unhide ederek de yapabilirsiniz. Ancak az önceki linkteki örnekteki 
		liste dinamik bir yapıya sahip olduğu için hide etmek bir uygun bir 
		çözüm olmamaktaır.</p>
		<p>
		Başka bir örneği ise burada ele alabiliriz. Bunda da yine gruplu bir 
		liste var. Bu listede bir hücreye çift tıklayınca bu hücreye ait alt 
		veriler ayrı bir dosya olarak açılıyor olsun. Ör:En çok kredi düşüşü 
		yaşayan şube listesinde şube koduna çift tıklayınca bize en çok düşüş 
		yaşayan müşteriler dosyasını açıp bu şubeyi filtrelesin.</p>
		<pre class="brush:vb">Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

adres = "C:\...\"

If Not Intersect(Target, Range("B2:B20")) Is Nothing Then
sb = Target.Value
Workbooks.Open adres + "Kredisi en çok düşen müşteriler - " &amp; Date - 1 &amp; " Sonuçları.xlsm", ReadOnly:=True
ActiveSheet.ListObjects("Query_from_DWH").Range.AutoFilter Field:=2, Criteria:=sb
End If

End Sub</pre>
		<p>
		Şahsen ben bu eylemi çok önemsiyorum. Bununla ADO'yu birleştirerek 
		yüksek ücretli programlara alternatif programlar yazabilirsiniz. ADO 
		kısmında diğer detayları bulabilirsiniz.</p>
	<h3>
	Worksheet_Activate/Deactivate</h3>
	<p>
	Belli bir sayfa (yeniden) aktif(veya inaktif) olduğunda çalışmasını istediğiniz kodları 
	bu olayla tetiklenen olay prosedürleri içine yazabilirsiniz.</p>
		<pre class="brush:vb">Private Sub Worksheet_Activate()

End Sub</pre>
		<p>Örneğin, ana menü sayfası gibi bir sayfanız var ve buna sadece diğer 
		sayfalardaki <strong>Anamenü </strong>linki aracılığı ile ulaşmak istiyorsunuz, ve bu sayfalar 
		açıken bu menü sayfası görünmesin istiyorsanız, işte bu menü sayafasından ayrılırken 
		sayfanın gizlenmesini sağlayacak bir kod yazabilirsiniz.</p>
		<pre class="brush:vb">Private Sub Worksheet_Deactivate()
   Me.Visible = xlSheetHidden 
End Sub

'aşağıdaki kodu da diğer sayfalardaki Selection_Change eventine yazarsınız
If Target.Value = "Anamenü" Then
   Sheets("Anamenü").Visible = xlSheetVisible
   Sheets("Anamenü").Select
End If</pre>
		<h3 id="calculate">Worksheet_Calculate</h3>
		<p>Bu event, sayfadaki formüller yeniden hesaplandığında tetiklenir. 
		Özetle o formülü etkileyen hücrelerden birinde değişiklik olursa 
		tetiklenir. Mesela Bir hücre grubunun altında SUBTOTAL formülü ile 
		toplam/ortalama v.s alınmışsa ve hücre grubundaki filtrede bir 
		değişiklik yapılırsa formülün içeriği de değişeceği için bu event 
		tetiklenir.</p>
		<p>Bu eventte hedef bir hücre(Target) bulunmaz, zira tüm hücreler 
			yeniden hesaplanmıştır.</p>
		<p>NOT: Sayfa için aynı zamanda Change eventi de varsa kod bloğu içine 
		eventleri geçici olarak bastıran kodları eklemeyi unutmayın.(Bu konuyu 
		hemen aşağıda inceleyeceğiz)</p>
		<pre class="brush:vb">Private Sub Worksheet_Calculate()

'çeşitli işlemler

End Sub</pre>
		<p>Bu konuya ait güzel bir örneği
		<a href="Ileriseviyekonular_PivotTableChartveSlicernesneleri.aspx#OrnekUygulama">
		şurada</a> bulabilirsiniz.</p>
	</div>
	<h2 class="baslik">Diğer Hususlar</h2>
	<div class="konu">
	<h3>Event tetiklenmesini bastırmak(Geçici olarak durdurmak)</h3>
		<p>
		Makronuzda, bir yerlerde ilgili eventi tekrar tetikleyecek bir kod varsa 
		bu kod sonsuz döngüye girer ve Excel çökebilir(veya ayarlarınıza göre 
		100 civarı iterasyon sonucunda durabilir, bende 78.iterasyonda duruyor). 
		Change eventi içinde bir hücrenin içeriği değiştirilmesi veya 
		SelectionChange eventi içinde başka bir hücre seçilmesi gibi.</p>
		<p>
		Mesela aşağıdaki örneği F8 ile deneyip görün, her F8 yapışınızda kod hiç 
		durmadan bir aşağı inecektir.</p>
		<pre class="bush:vb">Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   Target.Offset(1, 0).Select
End Sub</pre>
		<p>Aşağıdaki kodda ise sürekli olarak Change olayı kendisini tetikliyor.</p>

		<pre class="bush:vb">Private Sub Worksheet_Change(ByVal Target As Range)
   Target.Offset(1, 0).Value = Target.Row
End Sub</pre>

		<p>
		İşte bu tür durumları önlemek için eventin başında <strong>Application.EnableEvents = False
		</strong>diyerek eventleri geçici olarak askıya alırız, sonra işlemleri 
		yaptırır, sonra da <strong>Application.EnableEvents = True </strong>
		diyerek evetnleri tekrar devreye sokarız. Tabi olur da kodumuzda bir 
		hata oluşur da sona gelmeden durursa Eventler askıda kalabilir, bu 
		yüzden bir hata yönetimi bloğu yazıp eventleri burda da tekrar aktive 
		etmeliyiz.</p>
<pre class="brush:vb">
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo hata
    Application.EnableEvents = False
    'tetiklemeye neden olabilecek işlemler
    Application.EnableEvents = True
    Exit Sub

hata:
    Application.EnableEvents = True
End Sub
</pre>
	<h3>Workbook'un sheet eventleri</h3>
	<p>Workbook eventleri workbookla ilgili bir eylem gerçekleşince devreye 
	giriyordu, Worksheet eventleri de sayfayla ilgili bir eylem gerçekleşince. 
	Bir de ikisinin karışımı gibi olan ama aslında bir Workbook eventi olan 
	event grubu var.</p>
		<p>Bunların bir listesi aşağıdaki gibi olup, <strong>belli bir sayfada 
		değil de herhangi bir sayfada</strong> bir eylem gerçekleştiğinde 
		tetiklenirler.</p>
		<p><img src="../../images/vbawsolay1.jpg"></p>
		<p>Mesela aşağıdaki kod ile hangi sayfa seçilirse onun adı bize MsgBox 
		ile gösterilir.</p>
		<pre class="brush:vb">Private Sub Workbook_SheetActivate(ByVal Sh As Object)
   MsgBox Sh.Name
End Sub</pre>
	<h3>
	Farklı kullanıcılarda eventlerin tetiklendiğinden emin olmak</h3>
	<p>
	Giriş bölümündeki <a href="Giris_MakroNedir.aspx#guvenlik">Güvenlik ayarları</a> 
	bölümünü okumadıysanız öncelikle orayı okumanızı öneririm. Orada 
	belirtildiği gibi makro ayarları <strong>Disable All </strong>şeklindeyse sonuçta bir makro 
	olan Event Prosedürleriniz de devreye girmez.</p>
		<h4>
		Örnek senaryo</h4>
		<p>
		Hazırladığınız bir dosyanın anlamlı olabilmesi için eventlerin 
		çalışması gerekmekte olsun. Ancak kullanıcının makro ayarları Disabled 
		ise kullanıcı dosyadan istenen verimi alamayacaktır, üstelik sizin 
		istemediğiniz şekilde yetkisi olmayan görüntülemeler bile 
		yapabilecektir.(Farklı şubenin rakamlarını görmek gibi)</p>
		<p>
		Bunu engellemek için benim geliştirdiğim yöntem aşağıdaki 
		gibidir(Daha iyi veya daha kötü yöntemler var olabilir, ben 
		araştırdığımda hiçbirşeyle karşılaşmadığım için kendi çözümümü böyle 
		geliştirmiştim)</p>
		<p>
		Çalışmanın tam üstüne denk gelecek şekilde bir düğme koyarım ve bu düğme 
		için bir kod yazarım. Eğer 
		makrolar enable ise düğme kaybolur, makrolar disabled ise aşağıdaki gibi 
		bi hata alır.</p>
		<p>
		<img src="../../images/vbawsolay2.jpg"></p>
		<p>
		Ayrıca düğmeyi silmesin veya başka bi yere taşımasın diye sayfaya 
		protection da koymamız gerekiyor. Makro sırasında dosyayı gizlerken 
		geçici olarak kaldırıyor, gizledikten sonra tekrar koyuyoruz, ki 
		protection'ı başka amaçlar için de kullanabilelim. Buna ait bir örneği 
		<strong>Çeşitli Örnekler </strong>bölümünde 2.örnekte bulabilrisinz.</p>
		<p>
		<img src="../../images/vbawsolay3.jpg"></p>
		<p>
		Düğmenin Click eventi ise şöyledir.</p>
		<pre class="brush:vb">Sub Button1_Click()
   Sheets(1).Unprotect Password:="1234" 
   ActiveSheet.Shapes("Button 1").Visible = msoFalse 'düğmeyi gizler
   Sheets(1).Protect Password:="1234"
End Sub</pre>
	<h3>
	Kısıtlar uygulamak</h3>
	<h4>
	Sayfanın yazdırılmasını engellemek</h4>
		<p>
		Diyelim ki kullanıcıların belli sayfaları basmasını istemiyorsunuz. 
		Aşağıdaki kodu ilgili dosyanın Workbook_BeforePrint eventine yazmanız 
		gerekir.</p>
		<pre class="brush:vb">
Private Sub Workbook_BeforePrint(Cancel As Boolean)	
	For Each s In ActiveWorkbook.SelectedSheets
		If s.Name = "Ham Data" Then
		    MsgBox ("Bu sayfayı basamazsınız!!!")
		    Cancel = True  
		End If
	Next
End Sub</pre>

		<p>
		Workbook içinde hiçbir sayfanın bastırılmasını istemiyorsanız bu sefer 
		hiç safya kontrolü yapmadan doğrudan MsgBox ve Cancel=True satırları 
		yeterli olacaktır.</p>
		<p>
		Gördüğünüz gibi bu işlemi bir worksheet eventi ile değil workbook eventi ile 
		yapıyoruz.</p>
		<h4>
		Sayfada cut/copy engellemek</h4>
	<p>
	Bu işlemin tüm dosya bazında yapılmasıyla ilgili örnek
	<a href="Olaylar_WorkbookOlaylari.aspx#cutcopyengel">şurada</a> olup, sayfa 
	bazında yapmak için Worksheet_Activate ve Worksheet_Deactivate olaylarında 
	kullanılması yeterlidir.</p>


	</div>
	
	
	<h2 class="baslik">Çeşitli Örnekler</h2>
	<div class="konu">	
		<h4 class="baslik">Mevduat fiyatlama hesap makinası(Animasyonlu)</h4>
			<div><p>Bu örnekte, 4 parametreden oluşan bir denklemin herhangi 3'ü 
				bilinirken diğer 4.sünün tespit edilmesine yönelik bir kod 
				yazacağız. Klasik Excel yöntemiyle yapmak istediğinizde 4 ayrı 
				çalışma yapmanız gerekirken VBA ile tek bir format ile tüm 
				senaryoları ele alabileceğiz.<p>Bunun için aşağıdaki gibi bir 
				form hazırladım. Dosyanın kendisine
				<a href="../../Ornek_dosyalar/Makrolar/mevduat%20makinesi.xlsm">buradan</a> 
				ulaşabilirsiniz.<p>
				<img src="../../images/vbaworksheetolay2.jpg"><p>Çalışmaya ait 
				kodlar şöyle:<h5>Önce Sheet1 modülü:</h5>
				<pre class="brush:vb">
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo çıkış
    Application.EnableEvents = False
    
    If Target = [stopaj] And IsEmpty(Target) Then
        Target.Value = "0,15 (TL, 6 aya kadar)"
    End If
    
    If Not Intersect(Target, [alan]) Is Nothing Then
        ActiveSheet.Unprotect 1234
        Call temizlik([alan])
        If [alan].Cells.SpecialCells(xlCellTypeBlanks).Count = 1 Then            [alan].SpecialCells(xlCellTypeBlanks).Select
            Select Case ActiveCell
                Case [anapara]
                    ActiveCell.Formula = "=365*NetGetiri/(Vade*Faiz*(1-value(left(stopaj,4))))"
                Case [Faiz]
                    ActiveCell.Formula = "=365*NetGetiri/(Vade*Anapara*(1-value(left(stopaj,4))))"
                Case [Vade]
                    ActiveCell.Formula = "=365*NetGetiri/(Anapara*Faiz*(1-value(left(stopaj,4))))"
                Case [NetGetiri]
                    ActiveCell.Formula = "=Anapara*Faiz*Vade*(1-value(left(stopaj,4)))/365"
                Case Else
                    MsgBox "Böyle bir seçenek bulunmamaktadır"
            End Select
            ActiveCell.Font.Color = vbRed
            [uyarı].Value = ""
            Call Fontsizedeğiş(24, 20)
            Call alancopypaste
        End If
    End If
    Application.EnableEvents = True
    ActiveSheet.Protect 1234
    
    Exit Sub
    
çıkış:
If Err.Description = "No cells were found." Then 'blank sayısı 0 ise, count=0 kontrolüne gelmediği için o noktayı kaldırdım
    [uyarı].Select
    ActiveCell.Value = "Lütfen hangi alanın yeniden hesaplanmasını istiyorsanız onu silin."
    Call Fontsizedeğiş(14, 10)
End If
Application.EnableEvents = True
ActiveSheet.Protect 1234
End Sub
'----------------------------------------------------
Sub temizlik(alan As Range)
For Each a In alan
    a.Font.Color = vbBlack
Next a
End Sub
Sub Fontsizedeğiş(x As Integer, s As Integer)
    For i = 1 To 5
        Call Module2.beklet(s)
        DoEvents
        ActiveCell.Font.Size = x + i * 2
    Next i
    
    For i = 1 To 5
        Call Module2.beklet(s)
        DoEvents
        ActiveCell.Font.Size = x + 10 - i * 2
    Next i
End Sub
'----------------------------------------------------
Sub alancopypaste()
    For Each a In [alan]
        a.Value = a.Value
    Next a
End Sub</pre>
				<h5>Standart Modül içeriği</h5>
				<p>Bunda sleep metodu kullanıldğı için 
				aşağıdaki özel kod en başa eklenmiştir.
<pre class="brush:vb">
#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) '64 Bit Sistemler için
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '32 Bit Sistemler için
#End If

Sub beklet(sure As Integer)
    Sleep sure
End Sub</pre>
				<h5>Temizlik butonunun kodu ise şöyledir.</h5>
				<pre class="brush:vb">Sub Button1_Click()
  Range("alan").ClearContents
  ActiveSheet.Unprotect 1234
  [uyarı].Value = ""
  ActiveSheet.Protect 1234
End Sub</pre>
				<h5>Bu hesap makinesinin 
				kullanımı şöyledir:</h5>
				<p>Kullanıcı diyelim ki, bilinen olarak 
				anapara, faiz ve&nbsp; vadeyi girip müşterinin net kazancını 
				hesaplamak istiyor olsun. Bu üçünü yazınca net kazanç bilgisi 
				otomatik hesaplanır. Bu hesaplamanın sonucu da bir döngü ile 
				font hacminin önce büyüyüp sonra da küçülmesiyle animasyonlu bir 
				şekilde gösterilir.<p>Kullanıcı diyelim ki sonradan kazanç 
				bilgisini de manuel değiştirdi, o zaman tüm alanlar dolu olacağı 
				için kodumuz neye göre hesaplama yapacağını bilmez ve kullancıya 
				"Lütfen hangi alanın yeniden hesaplanmasını istiyorsanız onu 
				silin" mesajını yine animasyonlu bi şekilde gösterir.<p>Çalışma 
				mantığı ise şöyledir:<p>Sayfada belli name'ler tanımlanmış durumda. Makronun 
				tetiklenmesi için "alan" isimli namede bir hücrenin değişmesi beklenmekte. 
				Tabi değişklikler sonucunda başka tetkilenme olmasın diye 
				eventler geçici olarak baskılanmakta. Değişlik sonucunda alan 
				isimli name'de boş hücre sayısının 1 olup olmadığına 
				bakılmaktadır([alan].Cells.SpecialCells(xlCellTypeBlanks).Count 
				= 1 kodu ile). Böylece bu boş olana uygun formül yazılmakta ve 
				sonuç copy-paste yapılmaktadır. 
				<p>Alan isimli namede 2 hücre doluyken 3.sünün doldurulması 
				durumunda da, 4 hücre doluyken birinin silinmesi durumunda 
				kontrol sonucu 1 dönecek ve esas işi yapan kod bloğu çalışmış 
				olacaktır.</div>
		<h4 class="baslik">Data Validation ve Yetki kontrolü</h4>
			<div><p>En kısa sürede eklenecek</p>
			</div>
				<h4 class="baslik">Seçimlere göre veritabanından sonuç getirmek</h4>
				<div>
				<p>Bu işlem veritabanı kodlama bilgisi de gerektirdiği için 
				<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">buraya</a> 
				konulmuştur.</p></div>			
	</div>
</asp:Content>
