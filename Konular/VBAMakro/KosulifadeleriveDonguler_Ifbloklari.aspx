<%@ Page Title='KosulifadeleriveDonguler Ifbloklari' Language='C#' MasterPageFile='~/MasterPage.master' 
AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
	<div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td>
<td><asp:Label ID='Label2' runat='server' Text='Koşul ifadeleri ve Döngüler'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>If'li Koşul Yapıları</h1>
<p>Programlamanın temel taşlarından biri koşullu yapılarsa, koşullu yapıların 
temeli de IF bloklarıdır. Birçok programlama dilinde de zaten bu yapı kullanılır, 
sadece syntaxı farklı olabilmektedir.</p>
	<p>Genel yapı şu şekildedir: <strong>If....Then .... </strong>Tabi bu yapı 
	farklı görünümler kazanabilmektedir. Hepsine tek tek bakalım, sonra da 
	örnekleri verelim.</p>
	<h2 class="baslik">Temeller</h2>
	<div class="konu">
	<h3>Kullanım şekli</h3>
	<table class="alterantelitable">
	<th>Kullanım şekli</th>
		<th>Syntax</th>
		<tr>
			<td>Tek satırda yazım(kısa bir işlem yaptıracaksak)</td>
			<td>If isim="Volkan" Then Goto devam</td>
		</tr>
		<tr>
			<td>Tekli if bloğu. Birden fazla işlem yaptıracaksak If-End If bloğu 
			arasında yazılır.</td>
			<td>If isim="Volkan" Then<br>&nbsp;&nbsp; MsgBox "merhaba volkan"<br>&nbsp;&nbsp; 
			Goto devam<br>End If</td>
		</tr>
		<tr>
			<td>Çoklu if bloğu: Eğer şöyleyse şunu yap yok değilse bunu yap, 
			diğer durumlarda ise şunu yap.</td>
			<td>If isim="Volkan" Then<br>&nbsp;&nbsp; MsgBox "Merhaba volkan"<br>
			ElseIf isim="Serkan" Then<br>&nbsp;&nbsp; MsgBox "Merhaba serkan"<br>
			Else<br>&nbsp;&nbsp; MsgBox "Merhaba değerli insan"<br>End If</td>
		</tr>
		<tr>
			<td>İç içe If: Eğer şöyleyse, ve yine eğer böyleyse şunları yap</td>
			<td>If isim="Volkan" Then<br>&nbsp; If şehir="İstanbul Then<br>&nbsp;&nbsp; 
			MsgBox "merhaba istanbullu volkan"<br>&nbsp; Else<br>&nbsp;&nbsp; 
			MsgBox "merhaba anadolulu volkan"<br>&nbsp;&nbsp; End If<br>Else<br>&nbsp;&nbsp; 
			MsgBox "merhaba değerli insan"<br>End If&nbsp; </td>
		</tr>
		<tr>
			<td>IIF(tek satırda pratik değer atama)</td>
			<td>YasDurum=IIF(yas&lt;18;"Çocuk";Yetişkin")</td>
		</tr>
	</table>
	<p>Şimdi hemen küçük bir örnek yapalım:</p>
		<pre class="brush:vb">Sub basitif()
    Dim segment As integer
    Dim segmentAd as String

    segment=Application.Inputbox("Segment kodunu girin",Type:=1)
       
    If segment= 1 Then
        segmentAd = "Bireysel"
    Else
        segmentAd = "Tüzel"
    End If
    
    MsgBox "Müşteri " &amp; segmentAd &amp; " segmenttedir"
End Sub</pre>
		<h3>Operatörler</h3>
		<p>Temel operatörler şunlardır: eşitlik(=), eşit olmama(&lt;&gt;), büyük ve 
		büyükeşit(&gt;,&gt;=) ile küçük ve küçükeşit(&lt;,&lt;=)</p>
		<p>Bunlardan başka ne tür operatörler olduğunu aşağıda göreceğiz.</p>
</div>
<h2 class="baslik">İleri Seviye işlemler</h2>
<div class="konu">
	<h3>Boolean değişkenler ve If blokları</h3>
	<p>"If X=4" mü diye bir sorgulama yaparken aslında X=4 sonucu True mu False mu(doğru 
	mu yanlış mı) şeklinde 
	bir sorgulama yapmış oluyoruz. Aksi belirtilmedikçe tüm 
	sorgulamalarda sonucun True mu olduğunu sorgulamış oluruz, ve True ise Then 
	ifadesinden sonraki kısım işletilir.</p>
	<p>If bloklarında False sorgulaması da yapılabilir, bunun için açıkça If x= 
	False diye sormak gerekir, tabi x'in boolean tipinde olması kaydıyla. Böylece ifadenin 
	False olması durumunda da Then'den sonraki kısmın işletilmesi sağlanabilir. 
	Unutulmamalıdır ki yaratılmış ancak henüz değer atanmamış Boolean 
	değişkenlerin değeri False'tur. Bazen bu bilgi çok kullanışlı olmaktadır. 
	Bi örnek yapalım: Kullanıcıdan ismini girmesini isteriz ve 
	bu ismi, arasında ";" olacak şekilde 10 kere yazdırmak istiyoruz diyelim. Şimdi ilk başta 
	hatalı bir kod yazalım. Bu kodda en başta da istenmeyen bir ";" olacak.</p>
	<pre class="brush:vb">Sub booleanif1()

isim = InputBox("İsminizi girin")

For i = 1 To 10
   kelime = kelime + ";" + isim
Next i

MsgBox kelime
End Sub
</pre>
	<p>Gördüğümüz gibi en baştaki ; fazladan oldu.</p>
	<p><img src="/images/vbaifboolean1.jpg"></p>
	<p>Ancak kodumuzda aşağıdaki gibi False sorgulamasını yaparsak istediğimizi 
	elde ederiz.</p>
	<pre class="brush:vb">
Sub booleanif1()
Dim ilkDegerAtandımı As Boolean

isim = InputBox("İsminizi girin")

For i = 1 To 10
	If ilkDegerAtandımı = False Then 'İlk etapta Falsetur, ve false ise bu kısım işletilir
	   kelime = isim
	Else
	   kelime = kelime + ";" + isim
	End If

	ilkDegerAtandımı = True 'False'tan True'ya ilk geçişi burada yapıyoruz, sonrasında zaten 9 kez zaten True olan değere True atamış olacak
Next i

MsgBox kelime
End Sub</pre>

	<img src="/images/vbaifboolean2.jpg">
	
	<p>Başka bir örnek daha yapalım, bunu Interaktivite bölümünde de görmüştük, 
	kullanıcı InputBoxa giriş yapmadan çıkarsa bunu yakalayalım. Detay için
	<a href="Temeller_Interaktivite.aspx">buraya</a> bakınız.<br></p>
	<pre class="brush:vb">Sub booleanif2()

a = Application.InputBox("Yaşınızı girin", Type:=1)
If a = False Then 'Variant her değeri alabileceği için False ifadesi aynen yazılır
  MsgBox "Bir giriş yapılmadan çıkmayı tercih ettiniz"
  'Diğer kodlar buraya
Else
  MsgBox "Giriş yapıldı"
End If

End Sub</pre>
	<p><strong>NOT</strong>:Boolean ifadelerde True sorgulaması yaparken "<strong>=True</strong>" yazmaya gerek yoktur, 
	o yüzden içinde = operatörü olmayan bir sorgulama şekli gördüğünüzde 
	şaşırmayın. Aşağıda bir örnek var.</p>

<pre class="brush:vb">Sub yaskontrol()
Dim yas As Integer
Dim resitmi As Boolean

yas = Application.InputBox("yaşınızı girin", Type:=1)

If yas &gt;= 18 Then resitmi = True 'else durumunda False atamaya gerek yok hiçbir atama olmaz, atanmamaış boolenalar da default değeri yani false değeirini alır

If resitmi Then '=True yazmadık, yazabilirdik de.
   MsgBox "reşitsin"
Else
   MsgBox "reşit değilsin"
End If

End Sub</pre>

<p>Bu yöntemi sadece değişkenlerle değil Booelan tipli dönüşü olan fonksiyon veya propertylerle de kullanabiliriz.</p>

<pre class="brush:vb">
If Application.DisplayAlerts Then '=True yazmamama gerek yok çünkü DisplayAlerts propertysi Booelan döndürür
'kodlar
End If

</pre>



	<h3>Mantıksal kontroller</h3>
	<p>If ile birlikte kullanılan bazı built-in kontrol yapıları vardır. Yani =, 
	&lt;, &gt; ve is operatörlerinden başka bir de bu ifadelerle kontrol 
	sağlanır, ki bunlar da aslında Boolean sorgulama şekilleridir.</p>
	<ul>
		<li><strong>IsEmpty</strong>:<a href="../Fasulye/NeNeredeNasil_NullNothingEmptyveIlkdegeratama.aspx">(Variant 
		tipli</a>) değişken boşmu, yani henüz bir değer atanmadı mı? 
		Veya bir hücrenin içi boş mu?</li>
		<li><strong>IsDate</strong>:Değişken, tarihsel bir değişken mi? 
		Pratikte, Inputboxa girilen bir 
		tarihin doğru formatta girilip girilmediğini kontrol etmek için 
		kullanılır.&nbsp;&nbsp;&nbsp;&nbsp;
		</li>
		<li><strong>IsNumeric</strong>:Değişken sayısal mı? Inputboxa sayısal girilmesi 
		gereken bir verinin doğru girilip girilmediğini kontrol amaçlı 
		kullanılabildiği gibi, bir hücredeki verinin sayısal veri içerip 
		içermediğini kontrol etmek için de kullanılır. Gerçi Inputboxı 
		Application.Inputbox şeklinde kullanıp bir de veritipi olarak 1 seçersek 
		zaten kullanıcıyı otomatikman sayısal girmeye zorlamış oluruz ama normal 
		Inputboxla sorulduğunda bunu kontrol amaçlı kullanabilirsiniz.</li>
		<li><strong>IsNull</strong>:Değişken, bir veri içeriyor mu?</li>
	</ul>
	<p>Bütün bunları <strong>Not</strong> operatör ile ters mantıkta 
	sorgulayabilirsiniz. Bu konuyu bir alt bölümde göreceğiz.</p>
	<p>Çeşitli örnekler:</p>
	<pre class="brush:vb">
1.Örnek
'bulunduğunuz yerden ilk dolu hücreye kadar olan tüm boş hücreleri silmek için
'tabi bu örneği bir döngü içinde yapmak daha şık olurdu ancak burda örnek vermek
'adına If ve GoTo ile yapılmıştır
bas:
If IsEmpty(ActiveCell) Then 
   ActiveCell.EntireRow.Delete
   ActiveCell.Offset(1,0).Select
   GoTo bas
End If</pre>

    <p></p>

<pre class="brush:vb">
2.Örnek
'Be'Belli bir anda aktif hücrenin değerinin sayısal olmaması durumunda programı durduruyoruz
If IsNumeric(ActiveCell.Value) Then Exit Subre>
</pre>

<h3>If Not</h3>
	<p>Bazen bir kodu günlük konuşma dilinde olduğu gibi önce negatifini 
	sorgulayarak yazmak isteyebiliriz. X=y değilse şu kod çalışsın. Biraz kötü 
	bir örnek olacak ama mesela aşağıdaki kodda, sadece 18 yaşından küçükler 
	için çalışacak bir kod yazmış oluyoruz.</p>
<pre class="brush:vb">'Öncül kodlar
If Not Yaş&gt;18 Then
'işletilecek kod
End If

'diğer kodlar</pre>
	<p>Kötü bir örnek dedik çünkü&nbsp; burda if Yaş&lt;=18 diye yazsaydık daha 
	az kodla yazılmış olurdu. Evet olurdu, ancak bazen öyle durumlar olacak ki 
	günlük konuşma diline göre kod yazmak size daha rahat gelecek ve negatifi sorgulamak 
	da daha mantıklı 
	olacaktır. </p>
	<p>Şimdi mybooelan adına bir Boolean tipli değişkeniniz olsun. Bu durumda 
	<strong>If 
	Not myboolean</strong> yazmak ile <strong>If mybooolean=False </strong>yazmak arasında bir fark yok 
	gibi görünebilir ancak&nbsp;hem konuşma diline yakınlığı dolayısıyla 
	anlaşılırlığı hem de performans açısından <strong>Not</strong>'lı versiyon biraz daha 
	öndedir. Çünkü önce mybooleanını ne olduğunu hesaplıyor, bu 1; sonra bunun 
	False mu olduğuna bakıyor, bu da 2. Halbuki <strong>Not </strong>diyince sadece 
	1 değerlendirme işlemi yapmış oluyorsunuz. Ancak bu performans farkı o kadar 
	da büyük bir fark değil. Bu konu, daha çok okunaklılık ve kişisel tercihle 
	alakalıdır.</p>
	<p><strong>Not</strong> yöntemi pratikte en çok Boolean karşılaştırmalarda kullanılır ve 
	özellikle de Booelan değer dönüren bir fonksiyonu sonucu ile. Mesela 
	DosyaMevcutmu diye bir fonksiyonunuz olsun, eğer mevcutsa True, değilse 
	False döndürüyor olsun. Hergün saat 12de bir dosyanın oluşup oluşmadığına 
	bakan scheduled bir kodunuz olsun. Eğer dosya henüz oluşmamışsa kullanıcılara mail atmadan çıksın, 
	oluştuysa mail atsın.&nbsp; </p>
	<pre class="brush:vb">If Not DosyaMevcutmu(dosyadı) Then
   Exit Sub
Else
   'uzunca bir mail atma kodu
End If</pre>
	<p>Şimdi yine bir itiraz gelebilir, bu örnekte de aslında Not kullanmasak 
	olabilir, ilk başa DosyaMevcutmu diye yazar mail işlemini Else öncesine 
	koyabilir, Exit Sub'ı da Else sonrasına. Ancak gördüğünüz gibi mail işemi 
	oldukça 
	uzun olabilir, mesela 50 satır. Böyle bir durumda ilk bloğa daha kısa olan kodu koymak daha mantıklı ve okunaklıdır. 
	Böylece dosya oluşmadıysa ne yapılacağını çok net görebilyorum, ancak başa 
	dosya oluşma durumunu koyup Else'den sonra oluşmama durumuu koysaydık 
	Else'den sonra ne olduğunu görmek için scrollbarı aşağı indirmek 
	gerekecekti. Küçük bir detay ama bu tür pratiklikler size hep zaman kazandırır. 
	Zaten makroların amacı da bu değil mi, bize her anlamda zaman kazandırmak. O 
	yüzden kodlarımızı düzenlerken en okunaklı, bi zaman sonra tekrar içine 
	baktığınıza anlaması en kısa sürecek şekilde düzenlemekte fayda var. Bu 
	arada bu kod da pek tabiki <strong>Not DosyaMevcutmu</strong>&nbsp; yerine 
	<strong>DosyaMevcutmu =False </strong>şeklinde yapılabilirdi ama yine okunurluk açısından 
	<strong>Not</strong>'lı halini gündelik dile daha çok 
	benzetiyorum. Tercih sizin.</p>
	<h4>Mantıksal fonksiyonlarla Not kullanımı</h4>
	<p>Yine mantıksal fonksiyonların ters sorgulaması yapıldığında <strong>False
	</strong>ile sorgulamak yerine gündelik dile daha yakın olması açsıından
	<strong>Not </strong>ile sorgularız. Ör:Kullanıcının boş mu 
	dolu mu bir hücre seçtiğini kontrol edecek bir kod yazalım.</p>

<pre class="brush:vb">
tekrar:
Set alan=Application.InputBox("Hücre seçin", Type:=8)

If Not IsEmpty(alan) Then 'boş değilse. If IsEmpty(alan)=False de olur ama biraz garip görünür
    'işletilecek kodlar
Else
    MsgBox "Boş bir hücre seçtiniz, lütfen dolu bir hücre seçin"
    GoTo tekrar
End If
</pre>
	
<h4>Nothing ve Not</h4>
	<p>Bir de Not'ın tamamen zorunlu olduğu bir durum var ki, o da Nothing ile 
	birlikte kullanım şeklidir. Genel Syntax şöyledir: <strong>If Not nesne Is 
	Nothing. </strong>Burada nesne Object tipindeki herşey olabilir. Range, 
	Worksheet, Outlook application v.s. Kelime kelime tercümesi "nesne hiçlik 
	değilse" şeklinde yapılabilecek olan bu cümlecik aslında "nesne birşeyse, 
	yani birşey atanmışsa veya henüz boşaltılmadıysa" anlamında daha Türkçe olarak 
	yorumlanabilir. Ancak VBA'da "<strong>Birşey</strong>" yani <strong>
	Something</strong> diye bir ifade olmadığı için bunun zıddı olan "<strong>Not 
	Nothing</strong>" şeklinde yazılır.</p>
	<p>Mesela Worksheet olayları bölümünde sıkça 
	kullanacağımız bir <strong>Intersect</strong> metodu var. Worksheet değiştiğinde eğer değişen 
	hücreler belli bir aralıkta mı diye kontrol ederiz, işte burda <strong>Not 
	ve Nothing </strong>kombinasyonu kullaırlır. </p>
	<p><span>Ör: <strong>If Not Intersect(Target, 
	Range("C3:C4")) Is Nothing Then</strong>&nbsp;aslında şu demektir:&nbsp;<strong>If Intersect(Target, 
	Range("C3:C4")) Is Something Then</strong>, yani değişen hücre(Target) C3:C4 alanı 
	ile kesişim kümesindeyse, yani kesişim kümesi boşküme değilse, yani 
	"birşeyse". Worksheet olaylarında daha detaylı örnekler için
	<a href="Olaylar_WorksheetOlaylarievent.aspx">buraya</a> tıklayınız.</span></p>
	<p>Aşağıda da 3 değişkenden 2sine değer atanma durumu sözkonusu, idğer 
	açıklamalar commentlerde bulunuyor.</p>
	<pre class="brush:vb">
Sub ifnotnothing()

Dim hucre As Range
Dim ws As Worksheet
Dim wb As Workbook

Set hucre = ActiveCell
Set ws = ActiveSheet

If hucre Is Nothing Then 'if not nothing, hücrenin boş olup olmadığını sorgulamak değildir.
'hucre değişkenine bir değer atanıp atanmadığnı sorgulamaktır. boş olup olmadığı IsEmpty ile sorgulanır
    Debug.Print "hucre değişkenine atama yapılmamış"
Else
    Debug.Print "hucre değişkenine atama yapılmış"
End If

If Not ws Is Nothing Then 'if ws is something yani birşey ise
    Debug.Print "ws değişkenine atama yapılmış"
Else
    Debug.Print "ws değişkenine atama yapılmamış"
End If

'sadece sometinh olma durumunda çalışacak kod bloğu
If Not wb Is Nothing Then 'if wb is something yani birşey ise
    Debug.Print "buraya girmeyecek"
End If
End Sub	</pre>
	<h4>Koşulu terse çevirme</h4>
	<p>Bu arada boolean tipli özellikleri tersine çevirmek için o özelliğin 
	mevcut durumunu öğrenmek amacıyla If yapısını 
	kullanmaya gerek yoktur; <strong>Not</strong> operatörü ile doğrudan 
	tersine çevirilebilir. Mesela sayfadaki gridler açıkken kapatan, kapalıyken 
	açan(bu işlemlere toggle işlemleri denir) bir buton yapıp quickaccessbara 
	ekleyebilrisiniz. </p>
	<p>If'li kod şöyle olacaktır.</p>
	<pre class="brush:vb">
Sub grid_toggle()
If ActiveWindow.DisplayGridlines Then 'Dikkat ettiyeniz =True yazmadım, yazsam da bişey farketmeyecekti
   ActiveWindow.DisplayGridlines=False
Else
   ActiveWindow.DisplayGridlines=True
End If
End Sub</pre>
<p>Onun yerine basitçe şöyle de yazabilirsiniz:</p>
<pre class="brush:vb">
Sub grid_toggle()
  ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub
</pre>

<p>Tabiki, bunu yapabilmeniz için ilgli propertynin(özelliğin) hem 
	okunabilir hem yazılabilir(Returns and Sets) karakterde olması lazım.</p>
	<h3>İçiçe If ve Bağlaçlar(Birden fazla koşul sorgulama)</h3>
	<p>Birden fazla koşulu sorgulamak için içiçe birkaç IF bloğu yazılabilir. 
	Ör: Açılış Tarihi bu yıldan önce olan şubelerden sadece Bireysel Şube 
	sayısını öğreneceğimiz bir kod yazalım.</p>
	<pre class="brush:vb">If Subeaçılıştarihi&lt; Year(Now) Then
   If tip="Bireysel" Then
      'işletilecek kodlar
   End If
End If</pre>
	<p>Bunu yapmanını bir diğer yolu da <strong>And</strong> bağlacını kullanmak olacaktır.</p>
	<pre class="brush:vb">If Subeaçılıştarihi&lt; Year(Now) And tip="Bireysel" Then
      'işletilecek kodlar
End If</pre>
	<p>And kullanmak sanki daha mantıklı gibi ama ikinci koşuldan önce ara bir 
	işlem yapmak isterseniz içiçe IF kullanmanız gerekir.</p>
	<pre class="brush:vb">If Subeaçılıştarihi&lt;Year(Now) Then
   'ara işlemler
   If tip="Bireysel" Then
      'işletilecek kodlar
   End If
End If</pre>
	<p>Koşul yapısı <strong>Veya </strong>şeklinde olacaksa da <strong>Or
	</strong>kullanılır. Ör: Şube tipi 
	bireysel veya Karma ise şunları yap gibi.</p>
	<pre class="brush:vb">If tip="Bireysel" or tip="Karma" Then
      'işletilecek kodlar
End If</pre>
	<h4>Çoklu If bloğu vs Ayrı If blokları</h4>
	<p>Bazen öyle anlar gelir ki, konuşma dilinde söylediğimiz şey sanki çoklu if kullanmamız gerektiğini ima eder, ancak yapılması gereken işlem ayrı if 
	blokları kullanmak olabilir. Hemen bir örnek yapalım.</p>
	<p>Diyelim ki, bir schedule prosedürünüz var ve Excel açılır açılmaz devreye 
	giriyor.(Scheduling işlemleri için
	<a href="DortTemelNesne_Application.aspx#OnTime">buraya</a> bakınız). O an 
	saat 08:01den büyükse A makrosu, 08:11'den büyükse B makrosu v.s çalışsın 
	istiyorsunuz. Söylerken kulağa çoklu if kullanılacakmış gibi geliyor ancak 
	öyle yaptığımız durumda ilk koşul gerçekleşince kalanı işletilmez. Aslında 
	burda düşüncemizi dile getirme şeklimiz de hatalı olabilir. O yüzden 
	düşündüğümüz şeyi kelimelere daha doğru dökelim: A makrosu, o an saat 08:01'den 
	büyükse çalışsın, B makrosu saat 08:11'den büyükse çalışsın v.s. Küçük bir nüans var, 
	farkedebilidiniz mi? Evet, önce makronun adını sonra zamanı dile getirdik. 
	</p>
	<p>İlk versiyona göre kodumuzu şöyle hazırlardık ve hatalı bir işlem yapmış 
	olurduk:</p>
	<pre class="brush:vb">
If TimeValue(Now) > TimeValue("08:01:00") Then
    Application.OnTime Now + TimeValue("00:02:00"), Procedure:="pyskontrol"
ElseIf TimeValue(Now) > TimeValue("08:11:00") Then
    Application.OnTime Now + TimeValue("00:02:30"), Procedure:="nrkontrol"
ElseIf TimeValue(Now) > TimeValue("08:21:00") Then
    Application.OnTime Now + TimeValue("00:03:00"), Procedure:="pdmkontrol"
ElseIf TimeValue(Now) > TimeValue("08:31:00") Then
    Application.OnTime Now + TimeValue("00:04:00"), Procedure:="pargkontrol"
End If</pre>


	<p>Doğru kod aşağıdaki gibi olacaktır.</p>
	<pre class="brush:vb">
If TimeValue(Now) > TimeValue("08:01:00") Then
    Application.OnTime Now + TimeValue("00:02:00"), Procedure:="pyskontrol"
End If

If TimeValue(Now) > TimeValue("08:11:00") Then
    Application.OnTime Now + TimeValue("00:02:30"), Procedure:="nrkontrol"
End If

If TimeValue(Now) > TimeValue("08:21:00") Then
    Application.OnTime Now + TimeValue("00:03:00"), Procedure:="pdmkontrol"
End If

If TimeValue(Now) > TimeValue("08:31:00") Then
    Application.OnTime Now + TimeValue("00:04:00"), Procedure:="pargkontrol"
End If	</pre>
	<h3>IIf</h3>
	<p>Bir değişkene, bir koşul sonucuna bakarak bir değer atamak istiyorsanız
	<span class="keywordler">IIf</span> yapısı çok kulanışlıdır ve kullanımı 
	oldukça basittir.</p>
	<p>Aşağıdaki kodda, segment kodu 1 ise Bireysel değilse Tüzel şeklinde bir 
	atama yapılmaktadır.</p>
	<pre class="brush:vb">Segment=IIf(segmentkodu=1,"Bireysel","Tüzel")</pre>
	<p>&nbsp;</p>
	</div>
</asp:Content>
