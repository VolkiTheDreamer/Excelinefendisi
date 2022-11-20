<%@ Page Title='Application Olayları' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Olaylar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='5'></asp:Label></td></tr></table></div>
<h1>Application Olayları</h1>
	<h2 class="baslik">Genel</h2>
	<div class="konu">
	<p>Bundan önceki 3 bölümde Workbook, Worksheet ve Chart eventlerini gördük. 
	Hatırlarsanız Worksheet eventleri aynı zamanda Workbook eventleri içinde 
	de geçiyordu. Ör:Workbook_SheetActivate. Bu şu demekti; ilgili workbook içinde <strong>herhangi</strong> 
	bir sayfayla ilgili olay meydana geldiğinde bu event kodu tetiklenecekti.</p>
	<p>Aynı şekilde açık olan tüm dosya ve sayfaların tamamı için yani Excelin 
	kendisi için, bir diğer deyişle tüm uygulama(application) seviyesinde geçerli olacak şekilde, herhangi bir 
	Workbook veya Worksheet veya chart eventi meydana geldiğinde bir kodun 
	tetiklenmesini istiyorsak Application eventleri kullanırız.</p>
	<p>Ancak Application eventlerinin bi farkı var, bunlar diğer üçü gibi kod 
	penceresinde görünmezler, zira Excel Objects altında böyle bir nesne yoktur. 
	Aşağıda gördüğünüz üzere, Chart, Sheet ve Workbook var, o kadar.</p>
	<p>
	<img src="../../images/vbaeventapp1.jpg"></p>
	<p><span>Peki Application eventlerine nasıl ulaşırız? Bunun 2 yolu var:</span></p>
	<p>İlk yöntemde <span><strong>Object Browser</strong>'ı kullanıp Application nesnesine gelin, sonra sağ 
	tıklayıp <strong>Group Members </strong>diyin. Propertiler, Metodlar ve 
	Eventler sıralanacaktır. Eventleri en sonda görebilirsiniz.</span></p>
	<p>
	<img src="../../images/vbaeventapp2.jpg"></p>
	<p>Peki ulaştık ama bunları nasıl kullanıcaz? İşte hem bu eventleri görmenin ikinci 
	yöntemi hem de bunları kullanma yöntemine geldi sıra. Bunun da iki yolu var.</p>
	<h3>Class modülde Application nesnesi tanımlama</h3>
	<p>Şimdi ilgili projeye sağ tıklayıp <strong>Insert&gt;Class Module</strong> diyerek bir class 
	modülü ekleyelim. Evet, Application eventleri class modüller içinde tanımlanır, 
	zaten işlemleri normal modülde yaparsanız hata alırsınız. Sonra da tepeye şu 
	satırları yazın:</p>
	<pre class="brush:vb">Dim WithEvents myApp As Application</pre>
	<p> 	
	Şimdi sol üstteki combodan 
	myApp'i seçin ve sağdaki comboboxta eventleri 
	görün. Böylece eventleri görmenin ikinci yolunu da öğrenmiş olduk. Ama devam 
	edip nasıl kullanacağımıza bakalım.</p>
	<p> 	
	<img src="../../images/vbaeventapp3.jpg"></p>
	<p>
	NOT:Bu iki metodu workbbok ve worksheet 
	nesnelerinde de deneyebilirsiniz ama onların 
	eventi zaten kod peneceresinde üstteki comboboxlarda otomatik çıktığı için 
	bu zahmete gerek yok.</p>
	<p>
	Şimdi 
	devam ediyoruz, kulanım için birkaç aşamamız 
	daha var. Öncelikle class modülümüzde şu kodlar olsun. Bu arada properties 
	penceresinde classımızın <strong>Class1 </strong>olan adını <strong>myClass</strong> olarak 
	değiştirelim.</p>
	<pre class="brush:vb">
Private WithEvents myApp As Application
Private Sub Class_Initialize()
    Set myApp = Application
End Sub
Private Sub myApp_NewWorkbook(ByVal Wb As Workbook)
    MsgBox "Yeni açılan dosya adı:" & Wb.Name
End Sub</pre>
	<p>
	Şimdi de <strong>Workbook_Open</strong> eventine gidip, dosya açılır açılmaz bu classın 
	yaratılmasını sağlayalım.</p>
	<pre class="brush:vb">Dim myClassNesnesi As myClass
Private Sub Workbook_Open()
    Set myClassNesnesi = New myClass
End Sub</pre>
	<p>Dosyayı kaydedip, kapatalım ve tekrar açalım. Dosya açılınca bu classtan bir nesne yaratılacak, ve nesne 
	yaratılır yaratılmaz da bu classın Initialize eventi sayesinde kendine ait 
	bir eventi olan myApp nesnesi yaratılacaktır. Sonrasında her yeni bir dosya açılıdığında(Open 
	değil, New butonu ile) ilgili dosyanın adı görüntülenecektir. Siz isterseniz 
	bu eventin içini özelleştirebilirsiniz, zira bu haliyle pek kullanışlı 
	değil. Aşağıda Çeşitli Örnekler bölümünde biraz daha kullanışlı örnekler 
	bulabilirsiniz.</p>
		<p>&nbsp;NOT:Buraya kadar olan işlemlerde gördüğünz gibi Class'ları 
		kullandık. Class ve Class modüller hakkında daha fazla bilgiyi
		<a href="Ileriseviyekonular_ClassveClassModuller.aspx">buradan</a> 
		edinebilirsiniz.</p>
	<h3>Workbook modülde Application nesnesi tanımlama</h3>
	<p>Yukarda dedik ki App nesnesini tanımladığımız modülün Class 
	modül olması lazım. Bildiğiniz gibi Workbook ve Worksheet de aslında bir 
	class olup bunlara ait kod yazdığımız sayfalar da class modül sayfalarıdır. Dolayısıyla 
	bunlara ait sayfalarda da bu nesneyi tanımlayabiliriz. </p>

		<p>
		Şimdi, ThisWorkbook&nbsp;modülüne aşağıdaki kodu yazın.</p>
		<pre class="brush:vb">
Public WithEvents myApp As Application
Private Sub Workbook_Open()
    Set myApp = Application
End Sub</pre>
		<p>Şimdi de yine aynı penceredeyken sol üstte myApp'i seçerek 
		NewWorkbook eventine yine yukardaki gibi aynı kodu yazalım. Bu da bu 
		kadar.</p>
		<pre class="brush:vb">
Private Sub myApp_NewWorkbook(ByVal Wb As Workbook)
   MsgBox "Yeni açılan dosya adı:" & Wb.Name
End Sub</pre>
		<h3>
		Son Söz</h3>
		<h4>
		Hangi yöntem seçilmeli</h4>
		<p>
		Peki biz bu iki yöntemden hangisini kullanmalıyız. Tamamen kişisel 
		tercih olmakla birlikte üstatlara kulak verbiliriz. Derler ki, ayrı bir 
		class modül kodların derli toplu durması adına daha iyidir. Workbook ve 
		Worksheet sayfalarında onların kendine has kodlarını yazalım.</p>
	<h4>
	Olayların sırası</h4>
	<p>
	Elimizde myApp_SheetActivate,Workbook_SheetActivate ve Worksheet_Activate 
	olaylarının hepsi de varsa, hiyerarşik seviyeye göre işleme girerler. Yani 
	bir sayfa aktive olduğunda önce Application olayı, sonra Workbook olayı, en 
	son da sayfa olayı gerçekleşir.</p>
	</div>
	
	
		<h2 class="baslik">Çeşitli Örnekler</h2>
		<div class="konu">
				<h4 class="baslik">myApp_WorkbookOpen ile dış bağlantı tespiti</h4>
				<div class="konu">
							<p>Aşağıdaki örnekte açılan dosyalarda harici bir bağlantı var mı diye bize söyleyen 
							bir kod bulunuyor.</p>
							<p>Bunu yine yukardaki örnekte olduğu gibi Personal.xlsb dosyası içinde ele alalım.</p>
							<pre class="brush:vb">
					Private Sub myApp_WorkbookOpen(ByVal Wb As Workbook)
					    If Wb.Connections.Count > 0 Or Wb.Queries.Count > 0 Then
					        MsgBox "bu dosya harici bağlantı içeriyor"
					    End If
					End Sub
					</pre>
				</div>
					<h4 class="baslik">myApp_WorkbookBeforePrint ile Print alma engeli</h4>
					<div class="konu">
							<p>Olur da acil bir iş nedeniyle bilgisayarımızı açık bıraktık gittik, birileri gelip bizim 
							iznimiz olmadan bilgisayarımızdan birşeylerin printini almak isteyebilir. Bunu aşağıdaki kod
							ile engellemiş oluruz.</p>
							<p>Bunu da yine yukardaki örnekte olduğu gibi Personal.xlsb dosyası içinde ele alalım.</p>

							<pre class="brush:vb">
					Private Sub myApp_WorkbookBeforePrint(ByVal Wb As Workbook, Cancel As Boolean)
					    şifre = InputBox("Yazdırma şifresini girin")
						If şifre &lt;&gt; "1234" Then
						   Cancel = True
						MsgBox "Şifreyi bilmiyorsanız bu bilgisayardan hiçbir Excel dosyasının çıktısını alamazsınız"
					    End If
					End Sub
					</pre>
					</div>
		</div>
</asp:Content>
