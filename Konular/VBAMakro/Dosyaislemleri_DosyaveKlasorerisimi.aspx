<%@ Page Title='Dosyaislemleri DosyaveKlasorerisimi' Language='C#' MasterPageFile='~/MasterPage.master' 
AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
	<div id='gizliforkonu'>
<table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Dosya işlemleri'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr>
</table>
</div>
	<h1>Dosya ve Klasör erişimi</h1>
    <h2 class="baslik">Giriş</h2>
		<div class='konu'>
		<p> 
		Makro yazarken zaman zaman dosyaların/klasörlerin adını değiştirmek, taşımak, kopyalamak, 
		silmek v.b gibi manipule edici; bazen de bunların adını, 
	bulunduğu klasörü, adresini, dosya uzantısını öğrenmek v.b gibi bilgi 
		edinici işlemler yapmak isteriz.</p>
	<p> 
	Bu tür işlemler için iki temel modül var. <strong>FileSystemObject</strong>(<strong>FSO</strong> 
	olarak anılacak)<strong> </strong>Class Modülü ve <strong>FileSystem(FS)</strong> 
	Modülü. </p>
	<p> 
	FS normal bir modül olduğu için bunlardaki fonksiyonları doğrudan 
	kullanabiliyoruz. O yüzden FSO ile yapacak başka işlemimiz yoksa her zaman 
	FS fonksiyonlarını tercih etmeliyiz, tabi FS ile yapılamayan işlemleri de 
	mecburen FSO ile yapmamız gerekecektir.</p>
		<p> 
		Şimdi bunlara detaylıca bakalım. (Her zaman olduğu gibi sadece işimize 
		yarayan veya yarama ihtimali olan üyelere bakıyor olacağız. Ayrıca başka 
		sayfanın konusu olan üyelere o sayfalarda yer vereceğiz.)</p>
		<p> 
	<strong>NOT:</strong>Klasör/Dosya açma/seçme v.s gibi işlemleri bu bölümde değil
	<a href="DortTemelNesne_Application.aspx#filefolder">burada</a> görüyoruz.</p>
		</div>
    
    <h2 class="baslik">FileSystemObject(FSO)</h2>
		<div class='konu'>
		<h3> 
		Giriş</h3>
		<p> 
		FSO 
		nesnesi, web server üzerinde VBScript diliyle kullanılmak için 
		yaratılmıştır(ancak tabiki VBA dünyasından da kullanıma girmiştir). Bu yüzden ayrı bir dll(library) içinde bulunur: 
		<strong>Scripting Runtime</strong>. Bu nesneyi kullanabilmek için Tools&gt;Reference 
		menüsünden aşağıdaki gibi eklemek gerekir.(Tabi ki bu iş 
		<a href="Ileriseviyekonular_ObjelerDunyasi.aspx#binding">early binding</a> 
		için geçerli, late binding için böyle bir işleme gerek yoktur.)</p>
		<p> 
		<img src="/images/vbaiofsodll.jpg"></p>
	<pLate binding şekli aşağıdaki gibi olup intellisense sevdalısı olduğumuz için 
	şimdilik biz late binding kullamayacağız.</p>
		<p>Late binding için ise aşağıdaki gibi bir kod yeterli olacaktır. Ancak 
		biz kodlarımızda genellikle early binding kullanacağımız için devam 
		etmeden önce bu library'yi eklemenizi tavsiye ederim.</p>
			<pre class="brush:vb">Set FSO = CreateObject("Scripting.FileSystemObject") 'Late binding</pre>
	<h4> 
	Nesne 
	Hiyerarşisi</h4>
		<p> 
		Dosya işlemleriyle ilgili olarak tepede FSO 
		objesi bulunmaktadır. FSO kendi altında 
	sırasıyla şu nesneleri&amp;collectionları bulundurur.</p>
		<ul>
			<li>
			Drive(s)</li>
			<li>
			Folder(s)</li>
			<li>
			File(s)</li>
			<li>
TextStream 
			</li>
		</ul>
			<p>
			FSO'yu doğrudan kullanmak yerine ya bu alt nesneleri üretmek için 
			veya bu nesneleri temsil eden string ifadeler üzerinden işlem yapmak 
			için kullanırız.</p>
			<p>
			Mesela bir klasörü silmek için ya fso'nun DeleteFolder metodunu 
			kullanıp parametre olarak da ilgili klasörün adresini yazarız, ya da 
			GetFolder metodu ile bir Folder nesnesi yaratıp, sonra bu Folder 
			nesnesinin Delete metodunu kullanırız.</p>
			<p>
			Yanlız Folder nesnesini yaratırken dikkatli olmak lazım. Eğer ki, 
			üzerinde çalıştığınız dosyada aynı zamanda Outlook library'si için 
			de refearns tanımladıysanız, onun da bir Folder class'ı vardır. İki 
			class'ın karışmaması ve hata almamanız için Folder class'ının önüne 
			library adını yazmamız gerekir: <strong>Scripting.Folder</strong> 
			şeklinde.</p>
			<p>
			Zaten şu şekilde ikisi arasındaki ayrımı görmek kolaydır:</p>
			<pre class="brush:vb"> Dim fld As Scripting.Folder
Dim outfld As Folder 'outlook folder, bunu kullanmayacağız</pre>
			<p>İki değişkenin de intellisense'ie baktığımızda çıkan üyelerin 
			farklı olduğunu görüyoruz. </p>
			<p>Dosya Klasörü olan Folder'ın üyeleri böyle iken,</p>
			<p>
			<img src="../../images/fsoref1.jpg"></p>
			<p>
			Outlook Klasörü olan Folder'ın üyeleri böyledir.</p>
			<p>
			<img src="../../images/fsoref2.jpg"></p>
			<p>
			Madem ki bu nesneyi doğrudan kullanmayacağız, ikide 
		bir bu nesneden yaratmamak için bunu global seviyede public olarak 
		tanımlamak akıllıca olacaktır. Global tanımlamayacaksak her prosedürün 
		sonunda buna Nothing değerini atamak bellek yönetimi açısından iyi 
		olacaktır.</p>
		<p>
		Şimdi bu silme işlemine ait örneğe bakalım.</p>
		<pre class="brush:vb">
Public fso As New FileSystemObject 'global tanımlama
Sub foldersil()
    Dim fld As Scripting.Folder
    Dim outfld As Folder 'outlook folder, bunu kullanmayacağız
       
    fso.DeleteFolder "C:\Users\Volkan\Desktop\sil1"
    Set fld = fso.GetFolder("C:\Users\Volkan\Desktop\sil2")
    fld.Delete
End Sub
</pre>
<h4>Metodlar</h4>
			<p>FSO'nun en sık kullanacağımız 2 temel metodu şöyledir:</p>
		<ul>
			<li><span class="keywordler">GetFolder</span>:Folder(Klasör) nesnesi döndürür.</li>
			<li><span class="keywordler">GetFile</span>:File(Dosya) nesnesi döndürür.</li>
		</ul>
		<p>
		Bu ikisinden başka metodlar da var tabi ama bu ikisinin özelliği, diğer 
		2 temel 
	nesneyi yaratıyor olmaları. </p>
	<p> 
	Bu nesnenin metodlarına daha genel bir bakış ise şöyle olacaktır.</p>
	<p> 
	<table class="alterantelitable">
		<tr>
			<th>Metodlar</th>
			<th>Görevleri</th>
		</tr>	
		<tr>
			<td>
			GetDrive, GetFolder, GetFile</td>
			<td>
			Yukarda bahsettik.</td>
		</tr>
		<tr>
			<td>
			CreateFolder, CreateTextFile</td>
			<td>
			Yeni klasör ve dosya yaratır.</td>
		</tr>
		<tr>
			<td>
			DeleteFile, DeleteFolder</td>
			<td>
			Klasör ve Dosya siler</td>
		</tr>
		<tr>
			<td>
			CopyFile, CopyFolder</td>
			<td>
			Klasör ve Dosya kopyalar</td>
		</tr>
		<tr>
			<td>
			MoveFile, MoveFolder</td>
			<td>
			Klasör ve Dosya taşır</td>
		</tr>
		<tr>
			<td>
			DriveExists, FolderExists, FileExists</td>
			<td>
			İlgili birim mevcut mu kontrolü yapar</td>
		</tr>
	</table>
	</p>
		<p> 
		<strong>NOT</strong>:Fso işlemlerinde ilgili dosya/klasör adresi verilirken son 
		karakterin "/" olup olmaması önem arzet<span style="text-decoration: underline"><strong>me</strong></span>mektedir. Yani "C:\deneme" 
		ile "C:\deneme\" özdeştir.(Ancak daha 
		aşağıda göreceğimiz gibi "Dir" ile kullanırken durum farklıdır.)</p>
		<h3> 
		Folder ve File 
		işlemleri</h3>
	<p> 
	Folder/File işlemlerinde akılda tutulması gereken en önemli şey, öncelikle 
	dosyanın varolup olmadığını öğrenmeye çalışmaktır. Dosya özellikle 
	<strong>Application.FileDialog</strong> ile kullanıcıya seçtirilmemişse bu kontrol işlemini 
	mutlaka yapın derim. (FileDilaog ile yapılan seçimlerde bu kontrole gerek 
	yoktur.)</p>
		<h4> 
		Var mı? kontrolü	</h4>
		<p>Bunun içn <strong>FileExists</strong> ve <strong>FolderExists</strong> metotlarını 
	kullanırız.&nbsp; </p>
	<pre class="brush:vb">
Sub kontrol_fso()

If fso.FileExists("C:\Users\Volkan\Desktop\denemeler\deneme.xlsx") Then
    Debug.Print "Dosya var"
End If

If fso.FolderExists("C:\Users\Volkan\Desktop\denemeler\") Then 'sonda \ olup olmaması farketmez
    Debug.Print "Klasör var"
End If

End Sub</pre>
<h4>Klasör içindeki klasörleri elde etme</h4>
		<p>Bu işlem için <strong>SubFolders</strong> property'si kullanılır, bu 
		özellik bize Folders collection'ı döndürür. Gerisi ForEach yapmaktan 
		ibarettir.</p>
		<pre class="brush:vb">
Sub KlasördekiKlasörler()

Dim anaklasorStr As String
Dim fol As Folder, alt As Folder

anaklasorStr = "C:\Users\Volkan\OneDrive\Dökümanlar"
Set fol = fso.GetFolder(anaklasorStr)

For Each alt In fol.SubFolders
    Debug.Print alt.Name
    'diğer fso işlemleri
Next

End Sub</pre>
			<p>
			<strong>NOT</strong>:Bu işlemi FS modülündeki Dir ile de yapabiliyoruz. Başka FSO 
			işlemi yapacaksak(FolderExists kullanmak gibi) bu yöntemi kullanalım, 
			yoksa Dir yöntemini.</p>
<h4>Klasör içindeki dosyaları elde etme</h4>
		<p>Bu işlem için <strong>Files</strong> property'si kullanılır. Folder'da 
		olduğu gibi bi Collection elde eder ve Foreach uygularız.</p>
		<pre class="brush:vb">
Sub KlasördekiDosyalar()

Dim anaklasorStr As String
Dim fol As Folder
Dim f As File

anaklasorStr = "C:\Users\Volkan\OneDrive\Dökümanlar"
Set fol = fso.GetFolder(anaklasorStr)

For Each f In fol.Files
    Debug.Print f.Name
    'diğer fso işlemleri
Next

End Sub</pre>
			<p>
			<strong>NOT</strong>:Bu işlemi FS modülündeki Dir ile de yapabiliyoruz. Başka FSO 
			işlemi yapacaksak(FolderExists kullanmak gibi) bu yöntemi kullanalım, 
			yoksa Dir yöntemini.</p>
		<h4>
		Klasör içindeki (tüm alt klasörlerin içindekiler dahil) dosyaları elde etme</h4>
		<p>
		Bu örnek üsttekilere göre biraz daha karmaşıklık içerir ancak mantığı 
		açısından güzel bir örnektir. </p>
		<p>
		Alt klasörleri işleme dahil etmek için <strong>Collection</strong> 
		tipinde bir yığın oluşturuyoruz ve her defasında bu yığının ilk üyesi 
		üzerinde işlem yapıyoruz. İşlem yapmadan önce bu ilk elemanı yığından 
		dışarı atıyoruz ki bir daha işleme girmesin. Sonra yığındaki 
		elemanların(klasörlerin) her biri için işlemi yineliyoruz. Bu işlemi 
		anlamanın en iyi yolu, Local penceresi açıkken F8 ile ilerlemek 
		olacaktır.</p>
		<p>
		NOT:Bunun daha hızlı ve basit you aşağıda Dir bölümünde ele alınacaktır. 
		Ancak fso nesnesiyle ilgili başka kontroller veya işlemler yapılması 
		gerekirse bu yöntemin kullanılması tercih edilmelidir. Bu arada 
		internette araştırırsanız başka yöntemler olduğunu da görebilirsiniz. 
		Kullanım tercihi size kalmış.</p>
			<p>
			Bu örnekte bi klasördeki dosyaların adını, bulunduğu klasörü ve 
			dosya boyutunu yazdırıyoruz. Kendi diskinizdeki bir klasör ile yer 
			değiştirerek deneyebilrisiniz.(Örnek klasörü
			<a href="../../Ornek_dosyalar/Makrolar/FSO%20Örnek.rar">buradan</a> 
			indirebilirsiniz)</p>
		<pre class="brush:vb">
Sub KlasördekiAltKlasörDahilDosyalar()
    Dim kls As Scripting.Folder, altkls As Scripting.Folder
    Dim dosya As File
    Dim klasörYığını As New Collection
    Dim i As Integer
    Dim YığındaNeVar As String, sabtiStr As String
    Dim adım As Integer
        
    klasörYığını.Add fso.GetFolder("C:\Users\Volkan\Videos\Movavi Screen Capture Studio\Udemy Kurslar\2-ileri vba-makro\FSO Örnek")
    sabitstr = "C:\Users\Volkan\Videos\Movavi Screen Capture Studio\Udemy Kurslar\2-ileri vba-makro\"
    i = 1
    adım = 1
    
    'yığındaki eleman sayısı 0 olana kadar yani tüm alt klasörler bitene kadar devam edicez
    Do While klasörYığını.Count &gt; 0
                
        Set kls = klasörYığını(1)
        klasörYığını.Remove 1 'ilk klasörü yığından çıkarıyoruz
        
        'alt klasörleri yığına ekliyoruz
        For Each altkls In kls.SubFolders
            klasörYığını.Add altkls
        Next altkls
        
        'bu if bloğu informativedir, silinebilir
        If klasörYığını.Count &gt; 0 Then
            For Each k In klasörYığını
                YığındaNeVar = Replace(k, sabitstr, "") &amp; ";" &amp; Replace(YığındaNeVar, sabitstr, "")
            Next k
            ActiveCell(i, 1).Value = Mid(YığındaNeVar, 1, Len(YığındaNeVar) - 1)
        End If
        
        For Each dosya In kls.files
            ActiveCell(i, 2).Value = adım &amp; ". adımdaki klasör:" &amp; kls.Name 'informativedir, silinebilir
            ActiveCell(i, 2).Offset(0, 1).Value = Replace(dosya.ParentFolder, sabitstr, "")
            ActiveCell(i, 2).Offset(0, 2).Value = dosya.Name
            ActiveCell(i, 2).Offset(0, 3).Value = dosya.Size
            i = i + 1
        Next dosya
        YığındaNeVar = vbNullString
        adım = adım + 1
    Loop
End Sub
</pre>
			<p>
			Bu işlemi yapmanın bir diğer yolu da işlemi recursive bir şekilde ele 
			almaktır.</p>

	<pre class="brush:vb">'ana prosedür
Sub recursive_fulldosya()
    Dim anaklasorStr As String
    anaklasorStr = "C:\inetpub\wwwroot\aspnettest\excelefendi2\"
    Recursiveİlerle fso.GetFolder(anaklasorStr)
End Sub

'recursive prosedür
Sub Recursiveİlerle(kls As Variant) 'variant çünkü ilk girereken Folder sonra Folders olacak
    Dim altKlasorler As Variant
    Dim dosya As file
    Static i As Integer '<span>&nbsp;her defasında bir önceki değerini koruması için</span>
    
    On Error Resume Next 'erişim izni olmayan yerlerde hata almasın diye
    For Each altKlasorler In kls.SubFolders
	Debug.Print altKlasorler 'bilgi amaçlıdır
        Recursiveİlerle altKlasorler
    Next
    
    For Each dosya In kls.Files
        ActiveCell(i, 1).Value = dosya.ParentFolder
        ActiveCell(i, 1).Offset(0, 1).Value = dosya.Name
        ActiveCell(i, 1).Offset(0, 2).Value = dosya.Size
        i = i + 1
    Next
End Sub</pre>

		<h3>
		Diğer FSO ve File/Folder işlemleri</h3>
	<h4>Dosyaları ReadOnly yapmak</h4>
			<p>
			Bu işlemi de FS modülü ile yapabiliyoruz. Neden böyle 
			bir işlemi yapmak isteyeceğimi orada açıklıyorum. Burada sadece 
			kısaca bu işlemin nasıl yapıldığına bakalım. Önceki örneklerde 
			olduğu gibi eğer File nesnesi ile FileSystem ile yapılamayacak başka 
			işlemler yapacaksanız bu yöntemi kullanın, yoksa en hızlısı 
			FileSystem olduğu için onu kullanın.</p>
			<pre class="brush:vb">dosyaStr = "C:\deneme.xlsx" 
Set dosya = fso.GetFile(dosyaStr)
dosya.Attributes = 1 'bu özellik hem okunur hem yazılırdır</pre>
		<h4>
		Silme işlemi</h4>
		<p>
		Yine FS ile de yapılabilir ve öncekilerde olduğu gibi FS'nin Kill 
		fonksiyonunu kullanmak FSO'nun metodlarından daha efektiftir, özellikle 
		büyük çaplı işlemlerde.</p>
			<p>
			Bazı eylemler için FSO'nun metodlarını da File/Folder'ın metodlarını 
			da kullanabiliyoruz,&nbsp; FS'nin fonksiyonlarını da. Tıpkı silme 
			işleminde olduğu gibi. Aşağıdaki örnekte bu 3 yönteme de bakalım.&nbsp;</p>
		<pre class="brush:vb">
Sub silmeler()
Dim f As file

'1.yöntem:Filesystem modülündeki Kill fonk ile
a = "C:\Users\Volkan\Desktop\a.txt"
FileSystem.Kill a 'FileSystem yazmaya gerek yoktur

'2.yöntem:fso nesnesi ile. Fso'yu başka amaçla da kullancaksak
b = "C:\Users\Volkan\Desktop\b.txt"
If fso.FileExists(b) Then
    fso.DeleteFile b
End If

'3.yöntem:file nesnesi ile. File bilgisi lazımsa
c = "C:\Users\Volkan\Desktop\c.txt"
Set f = fso.GetFile(c)
If f.Size > 1024 Then
    f.Delete
End If

End Sub
</pre>
	<h4>Kopyalama</h4>
	<h5>Tek dosya kopyalamak</h5>
			<p>Kaynak olarak her zaman dosya adı belirtilir. Hedef olarak klasör 
			adı veya dosya adı belirtilebilir.(fso.CopyFile kaynakdosya, hedefklasör)</p>
			<pre class="brush:vb">'hedef: klasör
fso.CopyFile "C:\Users\Volkan\Desktop\denemeler\Şubeliste.xlsx", "C:\Users\Volkan\Desktop\ıvır zıvır\" 'sonda \ olmalı
'hedef: dosya adı, dosyanın adı dğeiştirilebilir
fso.CopyFile "C:\Users\Volkan\Desktop\denemeler\Şubeliste.xlsx", "C:\Users\Volkan\Desktop\ıvır zıvır\ŞubelisteforBölgeler.xlsx"</pre>
	<h5>Aynı tipteki çoklu dosya kopyalama</h5>
			<p>Aşağıdaki kod ile kaynak klasördeki tüm xlsx, xlsm, xlsb,xls 
			uzantılı dosyaları hedef klasöre kopyalamış oluyoruz.</p>
			<pre class="brush:vb">
kaynak = "C:\Users\Volkan\Desktop\denemeler"
hedef = "C:\Users\Volkan\Desktop\ıvır zıvır\"

fso.CopyFile Kaynak &amp; "\*.xl*", Hedef</pre>
	<h4>Yeniden adlandırmak ve taşımak</h4>
			<p><span class="keywordler">MoveFile </span>kullanılabileceği gibi MSDN'de dokumente edilmemiş bir 
			<span class="keywordler">Name </span>fonksiyonu var bu da kullanılabilir. MoveFile kullanımı 
			CopyFile'a benzer.</p>
			<pre class="brush:vb">Sub rename()

'kaynak dosya ve hedef klasör belirterek
fso.MoveFile "C:\Users\Volkan\Desktop\denemeler\Şubeliste.xlsx", "C:\Users\Volkan\Desktop\ıvır zıvır\" 'sonda "\" olmalı
'veya kaynak dosya, hedef dosya adı(farklı isim olabilir)
fso.MoveFile "C:\Users\Volkan\Desktop\denemeler\Şubeliste.xlsx", "C:\Users\Volkan\Desktop\ıvır zıvır\ŞubelisteforBölgeler.xlsx"
'veya
Name "C:\Users\Volkan\Desktop\denemeler\Şubeliste.xlsx" As "C:\Users\Volkan\Desktop\ıvır zıvır\ŞubelisteforBölgeler.xlsx"

End Sub</pre>
	<h4>Dosya son değişim tarihi hakkında bilgi almak</h4>
			<p>Bir dosyanın gün içinde birkaç kez açılma durumu varsa ve dosyada 
			bir kaydetme işlemi uygulanıyorsa, dosyanın daha önce kaydedilip 
			edilmediği bilgisine bakarak sonraki açılışlarda kodun çalışmamasını sağlayabilrisiniz.(veya 
			tamamen başka sebeplerle)</p>
	<pre class="brush:vb">Set f = fso.GetFile(gunlukyol &amp; adres)

If DateValue(f.DateLastModified) = Date Then
   Exit Sub
Else
   'diğer kodlar
End If</pre>
			<h4>Dosya isim, uzantı ve adresleri</h4>
			<p>Dosya isim, uzantı ve adreslerine sıklıkla ihtiyaç duyuyor 
			olacağız. Bunların açıklamasını doğrudan kod içinde vermek daha kolay 
			olacaktır.</p>
			<pre class="brush:vb">
Sub cesitli_fsofilefolder()
    Dim f As file
    Dim k As Folder
    dosya = "C:\Users\Volkan\Desktop\denemeler\deneme.xlsx"
    
    Debug.Print fso.GetAbsolutePathName(dosya) 'C:\Users\Volkan\Desktop\denemeler\deneme.xlsx
    Debug.Print fso.GetBaseName(dosya) 'deneme
    Debug.Print fso.GetDriveName(dosya) 'C:
    Debug.Print fso.GetExtensionName(dosya) 'xlsx
    Debug.Print fso.GetFileName(dosya) 'deneme.xlsx
    Debug.Print fso.GetParentFolderName(dosya) 'C:\Users\Volkan\Desktop\denemeler
    
    Set f = fso.GetFile(dosya)
        
    Debug.Print f.Name 'deneme.xlsx
    Debug.Print f.ParentFolder 'denemeler
    Debug.Print f.Path 'C:\Users\Volkan\Desktop\denemeler\deneme.xlsx
    Debug.Print f.ShortName 'DENEME~1.XLS
    Debug.Print f.ShortPath 'C:\Users\Volkan\Desktop\DENEME~1\DENEME~1.XLS
    
    Set k = f.ParentFolder
    
    Debug.Print k.Name 'denemeler
    Debug.Print k.ParentFolder 'C:\Users\Volkan\Desktop
    Debug.Print k.Path 'C:\Users\Volkan\Desktop\denemeler
    Debug.Print k.ShortName 'DENEME~1
    Debug.Print k.ShortPath 'C:\Users\Volkan\Desktop\DENEME~1
    
End Sub			</pre>
			</div>
		<h2 class="baslik">FileSystem(FS)</h2>
		<div class="konu">
		<p>
		Yukarıda belirttiğimiz gibi FS modülü, class modül olmayıp bunun içindeki 
		fonksiyonları kullanmak için bir FS nesnesi yaratmaya gerek 
		bulunmamaktadır, daha da önemlisi FSO gibi başka bir library'yi reference 
		olarak göstermeye gerek yoktur.</p>
			<p>
			Yine yukarıda belirttiğimiz gibi FSO ve FS'nin ortak üyeleri 
			bulunmaktadır. Amacımız sadece bu metodu kullanmak ise tercih her 
			zaman FS'den yana olmalı, ancak FSO'nun diğer üyeleriyle de işlem 
			yapılacaksa işte o zaman FSO kullanılmalıdır.</p>
			<p>
			Şimdi FS'nin metdolarına bakalım.</p>
			<h3>
			<strong>Dir</strong></h3>
		<p>
		<strong>Dir</strong> heralde FS'nin en önemli fonksiyonudur. Parametre 
		olarak dosya/klasör adı alır. Bunu kullanırken bir klasörle işlem 
		yapacaksak <strong>en sonda hep 
		"\" olmasına dikkat edilmelidir</strong>. Zira bu fonksiyoni aldığı 
		parametrenin dosya mı klasör mü olduğunu sondaki "\" ile anlıyor. O yüzden Dir ile ilgili işlem yapmadan önce sağ tarafa "\" 
		olup olmadığına göre aşağıdaki gibi bi düzeltme işlemi yapılmasında fayda var.(tabi 
		bir klasör ile uğraştığımızı düşünüyorsak)</p>
		<pre class="brush:vb">
If Right(klasor, 1) <> "\" Then
    klasor = klasor & "\"
End If	</pre>
		<p> 
		Bu fonksiyonün dönüş değeri String'tir. Peki ne döndürür? </p>
			<ul>
				<li>Parametre olarak Klasör adresi alıyorsa, klasördeki ilk 
				dosyanın adını</li>
				<li>Parametre olarak Dosya adresi alıyorsa, dosyanın kendisini</li>
				<li>Parametre olark ne alırsa alsın, eğer böyle bir dosya/klasör 
				yoksa ZLS, yani sıfır uzunluklu bir metin döndürür</li>
			</ul>
			<pre class="brush:vb">
Sub fs_dir()

kls1 = "C:\Users\Volkan\Desktop\denemeler" 'sonunda \ yok
kls2 = "C:\Users\Volkan\Desktop\denemeler\"
dsy = "C:\Users\Volkan\Desktop\denemeler\algo.xlsx"

Debug.Print Dir(kls1) 'hiçbişrey döndürmez
Debug.Print Dir(kls2) 'klasördeki ilk dosyayı döndürür
Debug.Print Dir(dsy) 'dosyanın kendisini döndürür

Debug.Print Len(Dir("var_olmayan_dosya_veya_klasör"))

End Sub			
			</pre>
			<p> 
		Bu fonksiyon FSO'nun FileExists/FolderExists özellikleri yerine de 
		kullanılabilir ancak bunlardan farklı olarak True/False döndürmez, 
			yukarıda belirttiğimiz gibi bir string döndürür, ilgili parametreyi 
			bulamazsa ZLS döndürür demiştik. O 
		yüzden sonucun vbNullString olup olmadığına bakmalıyız.</p>
		<h4> 
		Var mı? kontrolü</h4>
	<pre class="brush:vb">Sub kontrol2()

If Dir("C:\Users\Volkan\Desktop\denemeler\deneme.xlsx") &lt;&gt; vbNullString Then
    Debug.Print "Dosya var"
End If

If Dir("C:\Users\Volkan\Desktop\denemeler\") &lt;&gt; vbNullString Then 'sondaki \ işaretine dikkat. Bu klasör mevcut olsa bile \ işareti olmazsa Null döndürür
    Debug.Print "Klasör var"
End If

End Sub</pre>
		<p>Aşağıda Excel gurularından Ken Puls tarafından bir fonksiyon haline 
		dönüştürülmüş(benim de bir zamanlar sıklıkla kullandığım) versiyonunu görüyorsunuz.</p>
	<pre class="brush:vb">
Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists

    If strFullPath = vbNullString Then Exit Function
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function

'Kullanımı
    Sub varmı_test()
    'adres olarak dosya adresi de klasör adresi de verilebilir
    Const adres As String = "var_olmayan_dosya_veya_klasör"
    If FileFolderExists(adres) Then
        Debug.Print "var"
    Else
        Debug.Print "yok"
    End If
    
    End Sub</pre>
		<p>
		Yanlız bu tür şeylerin üzerine kafa yorunca siz daha basit çözümler 
		bulabiliyorsunuz. Mesela ben aşağıdaki çözümü buldum. Yaptığım şey 
		basit:Parametre olarak verilen şeyin hem file hem folder için işleyen
		<strong>GetAttr </strong>metodu ile attribute bilgisini almaya 
		çalışıyorum; eğer bilgi alabilirsem demek ki dosya veya klasör 
		mevcuttur; bilgi alamıyorsam, ki öyleyse hata üretir, o zaman hata 
		yönetim bloğu ile de fonksiyona False değerini atarım.</p>
			<pre class="brush:vb">Function Mevcutmu(adres As String) As Boolean
On Error GoTo hata
If GetAttr(adres) &gt;= 0 Then Mevcutmu = True
Exit Function

hata:
Mevcutmu = False
End Function

'kullanım
Sub mevcutmu_test()
Debug.Print Mevcutmu("C:\Users\Volkan\") 'True
Debug.Print Mevcutmu("C:\Users\Volki\") 'False
Debug.Print Mevcutmu("C:\Users\Volkan\Desktop\denemeler\deneme.xlsx") 'True
Debug.Print Mevcutmu("C:\Users\Volkan\Desktop\denemeler\deneme5.xlsx") 'false
End Sub</pre>
			<h4>
		İkinci parametre</h4>
			<p>
			Dir'in ikinci parametresi çeşitli constantları alır, bunların en 
			önemlisi vbDirectory'dir. İkinci parametre girilmezse(veya vbNormal 
			olarak girilirse, ki default olanı budur) ilk dosyayı elde ederiz 
			demiştik, işte bu ikinci parametre vbDirectory olarak girilirse 
			klasörün kendisini elde ederiz, 
			ve bunu "." olarak görürüz.</p>
			<p>
			Aşağıdaki klasör gözönüne alındığında;</p>
			<p>
			<img src="/images/vbafso2.jpg"></p>
		<pre class="brush:vb">
Sub dir_ikinciparametre()

Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\")  'klasördeki ilk dosyayı verir, ALGORİTMA.xlsx
Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\ALGORİTMA.xlsx")  'dosyanını kendisini verir
Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\", vbDirectory)  'Klasörün kendisi yani '.'
Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\ALGORİTMA.xlsx", vbDirectory) 'dosyanını kendisini verir

End Sub</pre>
			<h4>
			Çoklu kullanım</h4>
			<p>
			Dir'in iki parametresi de opsiyonel olmakla birlikte, ilk kullanımı 
			mutlaka parametreli olmalıdır. Ondan sonra hiç parametre almadan da 
			kullanılabilir. Parametresiz kullanım, <strong>"sonraki 
			dosya/klasöre geç"</strong> demektir.</p>
			<p>
			Dir'i bu şekilde arka arkaya kullandığımızda bi yerde artık sona 
			ulaşılır ve yeni bir dosya/klasör elde edilemez; işte bu noktada 
			sonuç ZLS'dir. O yüzden bu noktaya ulaşıldığında ya bu tekrarlı 
			çağırma işini sonlandırmalı, ki bu genelde bi döngü içinde 
			yapıldığından döngüdnen çıkmak anlamına gelir, ya da yeniden 
			parametreli bir kullanıma geçmelisiniz.</p>
			<pre class="brush:vb">
Sub çok_kullanım()

Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\")
Debug.Print Dir 'parantezsiz de kullanılabilir
Debug.Print Dir() 'parantezli de 
Debug.Print Dir("") 'klasördeki ilk dosyaya döner

Debug.Print Dir("C:\Users\Volkan\Desktop\denemeler\ALGORİTMA.xlsx")
Debug.Print Dir 'Dosyalarda bu ZLS döndürür
'Debug.Print Dir() 'Bu hata verir
Debug.Print Dir("") 'klasördeki ilk dosyaya döner

End Sub
</pre>
<h4>Klasör içindeki klasörleri elde etme&nbsp;</h4>
			<p>
			Bu işlemi FSO altında da incelemiştik. Anak eğer FSO ile ilgili başka bir 
			işlem yapılmayacaksa Dir ile çok daha hızlı yapılır. (Koddaki GetAttr'le ilgili satıra çok takılmayın, 
			bunu şimdilik ezbere kullanın, zira konu biraz karışık, merak edenler için
			<a href="https://msdn.microsoft.com/en-us/library/ee200232.aspx">burada</a> detaylı açıklama var.)</p>
		<pre class="brush:vb">
Sub LoopinKlasör()
Dim i As Integer
i = 1

anaklasor = "C:\inetpub\wwwroot\aspnettest\excelefendi2\"
klasor = Dir(anaklasor, vbDirectory)   'kendisini alarak başlıyoru

Do Until klasor = ""   'Dir'den klasor dönmeyene kadar ilerliyoruz
  If (GetAttr(anaklasor & klasor) And vbDirectory) = vbDirectory Then
    ActiveSheet.Cells(i, 1).Value = anaklasor & klasor
    i = i + 1
  End If
  klasor = Dir()   'Parametresiz dir ile sonraki klasöre ilerliyoruz(dosyalarla çalışsaydık sonraki dosya)
Loop
End Sub	</pre>
			<p>
			NOT: Bu döngü sırasında görünen "..", parent	klasörü temsil eder.</p>
<h4>Klasör içindeki dosyaları elde etme</h4>
		<p>
		Yine bu işlem de FSO ile yaplabilmekteydi ancak FSO'nun başka kullanımı 
		olmacyasak Dir ile yapmak daha hızlıdır. </p>
			<p>
			İlk parametreli Dir'den sonraki parametresiz Dir 
		ile aynı klasörde ilerleriz. Ayrıca Dir, joker eleman da aldığı için tüm 
		dosyalarda değil belli dosyalar üzerinde de dolaşabilirsiniz.</p>
		<p>
		Aşağıdaki örneklerde, bi klasörde "vba" ifadesi içieren dosya isimlerini 
		yazırıyoruz.
		350 dosya içeren bu klasörde Dir'in çalışma hızı 0,17 sn iken FSO 
		0,53 sn olmuştur. Daha kalabalık klasörlerde fark daha da açılmaktadır.</p>
			<pre class="brush:vb">'Dir ile
Sub dosyalarda_gezin_dir()
    Dim bas As Single, bts As Single
    bas = Timer
    f = Dir("C:\Users\Volkan\Desktop\denemeler\*al*") 'içide al geeçen ilk dosyayı bulur
    Debug.Print f
    Do While Len(f) &gt; 0
        f = Dir 'sonraki dosyaya geçiyoruz. Sonunda () gerekli değil, olsa da our olmasa da
        Debug.Print f
    Loop
    bts = Timer
    Debug.Print bts - bas
End Sub

'FSO ile daha uzun
Sub dosyalarda_gezin_fsofilefolder()
   Dim bas As Single, bts As Single

   Dim fso As New FileSystemObject
   Dim kaynak As Scripting.Folder
   Dim dosya As File
   
   bas = Timer
   Set kaynak = fso.GetFolder("C:\Users\Volkan\Desktop\denemeler\")
   For Each dosya In kaynak.files
      If InStr(dosya.Name, "al") &gt; 0 Then
            Debug.Print dosya.Name
      End If
   Next dosya
   bts = Timer
   Debug.Print bts - bas
End Sub</pre>
		<h4>
		Klasör içindeki (tüm alt klasörler dahil) dosyaları elde etme</h4>
			<p>
			Yine aynı açıklamamız geçerli. FSO ile de yapabiliriz, ancak Dir 
			daha hızlıdır. Eğer ki FSO'yu başka amaçla da kullanacaksak FSO 
			tercih edilmeli, yoksa Dir.</p>
			<p>
			Aşağıdaki örneği <a href="http://www.ammara.com">şuradan</a> aldım. 
			Kendi ihtiyacınıza göre değiştirebilirsiniz. F8 ve Local Windows 
			aracılığı ile değişklikleri izleyerek incelemenizi tavsiye ederim.</p>
			<pre class="brush:vb">
Sub recursive_Dir()
     Dim colFiles As New Collection
     RecursiveDir colFiles, "C:\deneme", "*.*", True
     Dim vFile As Variant
     For Each vFile In colFiles
         Debug.Print vFile
     Next vFile
End Sub

'recursive Fonksiyonumuz	
Public Function RecursiveDir(colFiles As Collection, _
                              strFolder As String, _
                              strFileSpec As String, _
                              bIncludeSubfolders As Boolean)
     Dim strTemp As String
     Dim colFolders As New Collection
     Dim vFolderName As Variant
     'Add files in strFolder matching strFileSpec to colFiles
     strFolder = TrailingSlash(strFolder)
     strTemp = Dir(strFolder & strFileSpec)
     Do While strTemp <> vbNullString
         colFiles.Add strFolder & strTemp
         strTemp = Dir
     Loop
     If bIncludeSubfolders Then
         'Fill colFolders with list of subdirectories of strFolder
         strTemp = Dir(strFolder, vbDirectory)
         Do While strTemp <> vbNullString
             If (strTemp <> ".") And (strTemp <> "..") Then
                 If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                     colFolders.Add strTemp
                 End If
             End If
             strTemp = Dir
         Loop
         'Call RecursiveDir for each subfolder in colFolders
         For Each vFolderName In colFolders
             Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
         Next vFolderName
     End If
End Function

'Dir kullanırken sondaki \ işareti önemli olduğu için bunu ele alma fonksiyonu
Public Function TrailingSlash(strFolder As String) As String
     If Len(strFolder) > 0 Then
         If Right(strFolder, 1) = "\" Then
             TrailingSlash = strFolder
         Else
             TrailingSlash = strFolder & "\"
         End If
     End If
End Function
</pre>
		<h3>
		Diğer FileSystem üyeleri</h3>
		<h4>
					
		ReadOnly yapmak</h4>
			<p>
					
			Bir dosyayı neden Readonly yapmak isteyesiniz ki? Sitenin anasayfasında 
			belirttiğim gibi herşeyi öğrenmeye çalışın, en azından böyle 
			birşeyin var olduğunu bilin, birgün lazım olabilir. Bendeki ihtiyaç 
			şöyle doğdu:</p>
			<p>
					
			İşyerinde 2 pc ile çalışıyorum. Ana pc'mdeki Personal.xlsb dosyamda 
			yaptığım değişikliklerin dosyanın Workbook_AfterSave eventi ile network üzerindeki 
			ortak alana kaydolmasını sağlıyorum. Diğer pc ve benimle çalışan 
			arkadaşımın da bendeki personal dosyasının en güncel verisyonuna 
			sahip olmasını istiyorum. Bunun için de onların pclerindeki Task 
			Scheduler ile her gece güncel dosyayı kendi pclerine kopyalamasını 
			sağlayan bir Task yarattım(Tabi önce Application.OnTime ile akşam belli bi saatte 
			Excelin kapatılmasını sağlıyorum). Süreç özetle şöyle:</p>
			<ul>
				<li>Personal.xlsb dosyamda bi değişklik yaparım</li>
				<li>Dosyanın güncel hali anında ortak alana kaydolur</li>
				<li>Akşam 21:00'de arkadaşımın pcsinde ve benim diğer pc'mde 
				Excel kapanır</li>
				<li>Gece 00:00da arkadaşımın ve benim diğer pc'mde Task 
				Scheduler önce güncel dosyayı bunların XLSTART klasörüne 
				kopyalar</li>
				<li>Hemen arkadasından Excel otomatik açılır</li>
			</ul>
			<p>
					
			Şimdi burda şöyle bi tehlike var. Eğer ki arkadaşım kendi pc'sindeki 
			Personal.xlsb dosyasında birşeyler denemek isteyip sonrasında bunu da 
			kaydetmek isterse ortak alana da bu değişkliklerle kaydedilmiş 
			olacak. Belki önemli bazı şeyleri bozmuş olacak. O yüzden 
			arkadaşımın bu dosya üzeinde Save etme işlemini engellemem gerekir. Bunun için 
			dosyanın ortak alana kaydolurken <strong>Readonly </strong>olarak kaydolmasını sağlayabilirim. 
			Böylece bu dosya, arkadaşımda ve benim diğer pc'de readonly 
			açılacak. PC'de sadece otomatik işler çalıştığı için orda zaten save 
			etme endişesi yok, ancak 
			arkadaşım save etmeye çalışınca ona uyarı çıkacak, ya başka 
			isimle kaydetmek zorunda kalacak ya da kaydetme işleminden vazgeçecek. İşte 
			bütün bu olay için kodumuz şöyle:</p>
			<pre class="brush:vb">
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  FileSystem.SetAttr Path & Me.Name, vbNormal 'Varolanın üzerine yazması için önce Readonly özelliğini kaldırıyorum
  fso.CopyFile Me.FullName, Path 'Dosya o an açıkken FilseSystem'in FileCopy metodu hata veriri. O yüzden fso nesnesini kullandık
  FileSystem.SetAttr Path & Me.Name, vbReadOnly 'en son readonly yapıyorum
End Sub
</pre>
			<h4>Diğer metodlar</h4>
			<p>

			<span class="keywordler">Kill ve RmDir:</span> Sırayla dosya ve 
			klasör silerler. Kill'in iki muadili var, FSO altında 
			karşılaştırması var. RmDir&nbsp; FSO'daki DelefeFolder muadilidir.</p>
			<p>

			<span class="keywordler">MkDir:</span>Klasör yaratır. FSO'daki CreateFolder muadili</p>
			<p>

			<span class="keywordler">FileCopy:</span> Fso'daki CopyFile'ın 
			muadilidir.</p>
			<p>

			<span class="keywordler">FileLen:</span> File'ın Size muadili.</p>
			<p>

			<span class="keywordler">FileDateTime:</span>Yine File'ın 
			DateLastModified muadili&nbsp;</p>
			<p>

			<span class="keywordler">GetAttr ve SetAttr:</span> File'ın 
			Attributes muadili.</p>
</div>
	
</asp:Content>
