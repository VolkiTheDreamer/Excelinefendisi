<%@ Page Title='Temeller Interaktivite' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' 
runat='server' Text='Temeller'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='4'></asp:Label>
</td></tr></table></div>

<h1>Interaktivite</h1>
<p> Makrolar kullanıcı ile belli başlı 4 şekilde iletişim kurar.</p>
	<ul>
		<li>Mesaj kutuları(MsgBox)</li>
		<li>Bilgi sorma kutuları (InputBox)</li>
		<li>Formlar(Userforms)</li>
		<li>File/Folder dialog kutuları</li>
	</ul>
	<p>Son ikisini ayrı bölümlerde işleyeceğiz, biz şimdilik burada MsgBox ve InputBox ile haşır neşir olacağız.</p>


	<h2 class="baslik">InputBox</h2>
	<div class="konu">
	<p>Inputbox ile kullanıcıya çeşitli sorular sorar, ondan bir şeyler 
	yazmasını veya sayfa üzerinde birşeyleri seçmesini bekleriz. Kullanıcının bu 
	girdiği değeri de bir değişken içinde depolarız. O yüzden Inputboxları tek 
	başına kullanmak yerine her zaman bir değişkene atama şeklinde kullanırız.</p>
<pre class="brush:vb">
Dim ad As String
ad = InputBox("Adınızı girin")		
</pre> 

<p>Inputbox'a girilen her 
		değer bir metindir, sayı olsa bile bu metin olarak depolanır. Girilen değeri sayı olarak kullanmak istiyorsanız bunu
		<span class="keywordler">Val</span>veya buna benzer bir dönüştürme metodu (<span class="keywordler">Int, 
		CInt, CLng, CDbl gibi</span>) ile sayıya çevirmeniz 
		gerekir. Aksi halde istenmeyen sonuçlar ortaya çıkabilir. Bir 
		örnekle bakalım</p>

<pre class="brush:vb">
Sub input1()

a = InputBox("bir sayı girin")
b = InputBox("ikinci bir sayı girin")

Range("A1").Value = a + b

End Sub</pre> 
		<p>Kodu çalıştıralım, a için 5, b için 7 girelim. A1 hücresinde 12 
		rakamını görmeyi bekleriz ama 57 yazar. Çünkü VBA metinler için 
		birleştirme operatörü olarak + işaretini de kullanılır(bir de & işareti var).</p>

		<p>Şimdi aynı kodu aşağıda gibi çalıştıralım, aynı değerleri girelim, bu 
		sefer 12 sonucunu görebiliriz.</p>
		<pre class="brush:vb">
Sub input2()

a = InputBox("bir sayı girin")
b = InputBox("ikinci bir sayı girin")

Range("A1").Value = Val(a) + Val(b)

End Sub</pre> 
		<p>Bir diğer alternatif de, numerik olmasını istediğimiz değişkenleri 
		baştan numerik olarak tanımlamaktır.</p>
		<pre class="brush:vb">
Sub input3()
Dim a As Integer
Dim b As Integer

a = InputBox("bir sayı girin")
b = InputBox("ikinci bir sayı girin")

Range("A1").Value = a + b '12 yazar

End Sub</pre>
		<h3>Kullanım şekli ve diğer InputBox</h3>
		<p><a href="Giris_ExcelNesneModeli.aspx">Nesne Modelini</a> anlatırken Classlardan ve Library'lerden bahsetmiştim. İşte bu yukarda gördüğümüz InputBox da 
		<strong>VBA Library</strong>'si 
		içinde <strong>Interaction</strong> class'ına ait bir fonksiyondur. Birçok fonksiyon gibi bu da 
		parametre alır. </p>
		<p><abbr title="Fonksiyonların parametreleriye birlikte yazım şekli">Syntax'ı</abbr> şöyledir: <span class="keywordler">
		InputBox(Prompt[,title][,default][,xpos][,ypos][,helpfile,context])
		</span></p>
		<p>Bu parametrelerden sadece köşeli parantez içine alınmamış olan Prompt 
		parametresi zorunlu olup diğerleri opsiyonedir. Önemlilerin açıklaması 
		ise şöyledir:</p>
		<ul>
			<li><strong>Prompt</strong>:Kullanıcıya ne girişi yapmasını söyleyeceğimiz ifade. 
			Ör:Adınızı giriniz.
			<li><strong>Title</strong>:Inputbox kutusunun başlığını set edebilirsiniz</li>
			<li><strong>Default</strong>:Faydalı bir özelliktir, kullanıcıya bazı durumlarda kutu 
			içinde hazır bir değer sunabilirsiniz. Genelde, en çok girilen 
			değerleri tahmin ederek girebilirsiniz. Örneğin "Bir il kodu girin" 
			diyip, default değer olarak da İstanbul'un kodu olan 34ü 
			yazabilirsiniz.</li>
		</ul>
		<p>
		Yukarda Inputbox'ın bize aktif sayfadan bir seçim de 
		yaptırabileceğini söylemiştim. Ancak yukardaki kodları 
		çalıştırdığınızda bunu yapamazsınız, isterseniz bi deneyin, sonra tekrar 
		gelin. Peki neden böyle söyledim. Çünkü bir Inputbox'ımız daha var, bu 
		işlemi o yapar ve kendisi <strong>Excel Library</strong>'sindeki <strong>Application</strong> 
		nesnesinin bir <span style="text-decoration: underline">metodudur</span>. <a href="Temeller_Terminoloji.aspx">
		Terminoloji</a> sayfasında belirttiğimiz gibi, bu Inputbox metodu diğer 
		metodlar gibi bir nesneye ihtiyaç duyar, yani tek başına kullanılamaz, o 
		nesne de Application nesnesidir. İlki function iken ikincisi metoddur.</p>
		<p>
		Zaten aşağıdaki resimden de görüleceği üzere 
		bağlı oldukları classların iconları bile farklı. </p>
		<p>
		<img src="/images/vba_temeller_inputbox.jpg" class="zoomla" width="60%" height="60%">
		</p>
<p>
İkinci Inputbox'ımızın syntax'ı ise şöyledir:<span class="keywordler">Application.InputBox(Prompt,Title,Default,Left,Top,HelpFile,HelpContextID,Type)</span></p>

<p>
		Bir önceki InputBox'tan farklı olarak en sonda bir Type parametresi 
		görüyoruz. Bu parametrenin alabileceği değerleri ve anlamları aşağıda 
		verilmiştir. En sık kullanacaklarımız koyu gösterilmiştir.</p>
		<table class="alterantelitable">
			
			<th>Değer</th>
			<th>Anlam</th>
			
			<tr>
				<td>0</td>
				<td>Formül</td>
			</tr>
			<tr>
				<td><strong>1</strong></td>
				<td><strong>Sayı</strong></td>
			</tr>
			<tr>
				<td><strong>2</strong></td>
				<td><strong>Metin</strong></td>
			</tr>
			<tr>
				<td>4</td>
				<td>True/False</td>
			</tr>
			<tr>
				<td><strong>8</strong></td>
				<td><strong>Range(Bir hücre grubu)</strong></td>
			</tr>
			<tr>
				<td>16</td>
				<td>Hata değeri</td>
			</tr>
			<tr>
				<td>64</td>
				<td>Dizi</td>
			</tr>
		</table>
		<p>Tablodan da görüleceği üzere kullanıcıya bir hücre grubu seçtirmek 
		için Type parametresini 8 tipinde belirtmemiz gerekiyor. Eğer kullanıcı hem metin 
		hem sayısal birşey girebilecekse Type değerine toplam değer olan 3(1+2) 
		yazılır.</p>
		<p>Hemen bir örnek yapalım.</p>
		<pre class="brush:vb">
Dim sonHucre As Range
Set sonHucre = Application.InputBox(Prompt:="Son hücreyi seçin", Type:=8)		
		</pre>
		<p>Değişkenlerle ilgili <a href="Temeller_DegiskenlerveVeriTipleri.aspx">
		sayfadan</a> hatırlayacağınız üzere nesnelere değer atamak için 
		<span class="keywordler">Set</span> ifadesini kullanıyorduk, burada da öyle yaptık. </p>
		<h3>Boş geçilen kutular(Cancel veya Esc ile iptal)</h3>
		<p>Bazen kullanıcılar hiçbir değer girmeden çıkmak 
		ister, o zaman ne olur. </p>
		<ul>
			<li>Klasik Inputbox'ın dönüş değeri olan 
		Stringtir ve bu durumda ilgili değişkene String tipinin default değeri atanır, yani 
			<strong>""</strong>. O yüzden değişkenin değerinin "" olup 
			olmadığı kontrol edilir.</li>
			<li>Application.Inputbox metodununu dönüş değeri Varianttır, o yüzden default değer olarak <strong>Empty</strong> bekleriz ancak 
			MSDN bize bu Inputbox'ta boş geçilen değerler için atanan değerin <strong>False</strong> 
			olduğunu söylüyor. O yüzden değişkenin değerini False olup olmadığı kontrol 
			edilir, ama bu klasik Inputboxa göre biraz daha alengirlidir. 
			Aşağıdaki örneklere bakalım.</li>
		</ul>
		<p>Kodumuzda hatalı birşey olmaması için bazı kontroller yapmamız 
		gerekiyor. Bundan sonrasına devam etmeden önce koşullu yapıları 
		bildiğinizden emin olun, bilmiyorsanız
		<a href="KosulifadeleriveDonguler_Ifbloklari.aspx">buradan</a> kısa bir 
		bilgi edinip tekrar buraya gelin.</p>
		<pre class="brush:vb">
'klasik Inputbox
a=Inputbox("Bir değer girin")
If a<>"" Then
   Msgbox "Giriş yapıldı"
   'Diğer kodlar buraya
Else
   Msgbox "Bir giriş yapılmadan çıkmayı tercih ettiniz"			
End If

'Application'lı, String
Dim a As String
a=Application.Inputbox("Adınızı girin", Type:=2)
If a<>"False" Then 'False'ın tırnak içinde yazıldığına dikkat edin
	Msgbox "Giriş yapıldı"
	'Diğer kodlar buraya
Else
	Msgbox "Bir giriş yapılmadan çıkmayı tercih ettiniz"			
End If		

'Application'lı, Integer(değişken tanımlanmaz, yani Varianttır)
a=Application.Inputbox("Yaşınızı girin", Type:=1)
If a<>False Then 'Variant her değeri alabilecğei için False ifadesi aynen yazılır
	Msgbox "Giriş yapıldı"
	'Diğer kodlar buraya
Else
	Msgbox "Bir giriş yapılmadan çıkmayı tercih ettiniz"			
End If

'Application'lı, Integer(değişken tanımlanır)
Dim a As Integer
a=Application.Inputbox("Yaşınızı girin", Type:=1)
If a<>0 Then 'Sayısal ifadelerde False veya False'ın rakamsal karşılığı olan 0 kullanılabilir
	Msgbox "Giriş yapıldı"
	'Diğer kodlar buraya
Else
	Msgbox "Bir giriş yapılmadan çıkmayı tercih ettiniz"			
End If

'Applicationlu, Range
'Range seçiminde eğer kullanıcı seçim yapmazsa hata oluşur, bu yüzden bir hata kontrol mekanizması da ekleriz
've ayrıca bir seçim yapıp yapmadığını da Nothing ile kontrol ederiz		
On Error Resume Next 'burayı yazmassak hata alırız. Hata yönetim mekanizmaları için ilgili sayfaya gidip bilgi edinebilirsiniz
Dim a As Range
Set a = Application.InputBox("Bir hücre seçin", Type:=8)
If Not a Is Nothing Then
	Msgbox "Seçim yapıldı"
	'Diğer kodlar buraya
Else
	Msgbox "Bir seçim yapılmadan çıkmayı tercih ettiniz"			    
End If
</pre>
<p>Şimdi son olarak tam bir örnek yapalım. Bu örnekte kullanıcıdan açık olan 
dosyaya kaç sayfa eklemek istediğini soracağız, detaylara takılmayın, sadece 
yukardaki anlatılanları pekiştirmeye çalışın.</p>
		<pre class="brush:vb">
Sub Sayfaekle()
Dim i As Integer, syf As Integer

syf = Application.InputBox("Kaç sayfa ekleyelim", Default:=3, Type:=1)

If syf = False Then 'escape'e baıslıysa veya Cancel'a tıklandıysa. Bunu ayrıca if syf= 0 diye de yapabilrdik
    Exit Sub
Else
    For i = 1 To syf
        Worksheets.Add
    Next i
End If
End Sub</pre>
	</div>

	<h2 class='baslik'>MsgBox</h2>
<div class='konu'>
<p>MsgBox ile ya bilgilendirme yaparız, ya da cevabı Evet/Hayır gibi 
sorular sorup bilgi ediniriz. Bilgilendirme yaptığımızda bunu bir değişkene 
atamaya gerek yoktur, ancak bilgi topladığımızda Inputboxta olduğu gibi bir değişkene atamamız lazım.</p>

	<p>MsgBox da InputBox gibi VBA Library'sindeki Interaction sınıfı içinde yer 
	alır ve syntax'ı şöyledir: <span class="keywordler">MsgBox(prompt[,&nbsp;buttons] 
	[,&nbsp;title] [,&nbsp;helpfile,&nbsp;context])</span></p>

	<p>Burda prompt ve title InputBoxtaki gibidir, son iki parametreden 
	bahsetmeyeceğim, arzu eden araştırabilir. Burda önemli bir parametre var: buttons parametresi. Bu parametrenin alabileceği değerler 
	şöyledir(Liste daha uzun ama çoğu gereksiz olduğu için buraya almadım, hatta 
	bunlardan da en çok YesNo ve YesNoCancel düğmelerini kullanacağımızı söyleyebilirim)</p>
	<p><img src="/images/vba_interaktif_msgbox.jpg"></p>
	<p>
	Aşağıda bilgilendirmeye örnek bir kod var<pre class="brush:vb">
Sub MessageBox()
	'Uzunca bir kod bloğu

	MsgBox "İşlem tamamdır"
End sub	</pre>
	<p>Bilgi edinme örneği ise şöyle birşey olabilir.</p>
	<pre class="brush:vb">
Sub MessageBox()
	cvp = MsgBox("Ana diskinizde(Ör:'C:') 'böl' isminde bir klasörünüz var mı?", vbYesNo) ' bu bilgi toplama mesajı
	If cvp= 6 Then 'yes demek oluyor
		GoTo ilerle
	Else
		MsgBox "O ZAMAN O KLASÖRÜ YARATIP TEKRAR ÇALIŞTIR" 'bu bilgi mesajı
		Exit Sub
	End If
ilerle:
'diğer kodlar
End sub	
</pre>
	<p>Gördüğünüz üzere cvp değerinin değerini 6 gibi bir sayıyla ölçtük. İşte 
	VBA'da bazı sabitlerin(constant) böyle sayısal değerleri vardır, ikisi de 
	kullanılabilir. Tüm düğmeler ve değerleri şöyle.</p>
	<table class="alterantelitable">
			<th>Sabit</th>
			<th>Değer</th>

		<tr>
			<td>vbOK</td>
			<td>1</td>
		</tr>
		<tr>
			<td>vbCancel</td>
			<td>2</td>
		</tr>
		<tr>
			<td>vbAbort</td>
			<td>3</td>
		</tr>
		<tr>
			<td>vbRetry</td>
			<td>4</td>
		</tr>
		<tr>
			<td>vbIgnore</td>
			<td>5</td>
		</tr>
		<tr>
			<td>vbYes</td>
			<td>6</td>
		</tr>
		<tr>
			<td>vbNo</td>
			<td>7</td>
		</tr>
	</table>

<p>InputBox'ta olduğu gibi MsgBox'ın da iptal edilmesi sözkonusu olabilmektedir. Tabi eğer buton türü olarak Cancel varsa. Aksi halde Esc tuşu da işe yaramamaktadır.</p>

<p>Bu örnekten çıkış mümkün değilken,</p>

<pre class="brush:vb">
Sub msgbox1()

On Error GoTo hata
a = MsgBox("Cevap verirmisin", vbYesNo)
'Diğer kodlar
Exit Sub
hata:
Debug.Print Err.Description
    
End Sub

</pre>

<p>Ama bunu iptal edebilirsiniz.</p>
<pre class="brush:vb">
Sub msgbox1()

On Error GoTo hata
a = MsgBox("Cevap verirmisin", vbYesNoCancel)
If a = vbYes Then
    MsgBox "Evet denildi"
ElseIf a = vbNo Then
    MsgBox "Hayır denildi"
Else
    MsgBox "Seçimi iptal ettiniz"
End If
Exit Sub
 
hata:
Debug.Print Err.Description
    
End Sub

</pre>

<p>Önemli bir detay da, MsgBox'ın bilgi toplama formundayken mutlaka ()'ler içinde kullanılmasıdır. Mesaj verirken ise genelde () olmadan kullanılır, ama parantezli kullanımı da 
sorunsuz çalışır.</p>

</div>

	<h2 class='baslik'>Kullanıcı dostu mesajlar</h2>
<div class='konu'><p> Şimdi kod yazmada biraz deneyim kazandığımıza göre uzun 
	kodlar yazarken nelere dikkat etmemiz gerekir ona bir bakalım. </p>
	<h3> Kullanıcı dostu kodlama</h3>
	<p> Kullanıcılara bazen MsgBox ile bazen Inputbox ile çeşitli mesajlar 
	yayınlamak gerekecek. Kullanıcı bu mesajları rahat okusun diye gerekli 
	yerlerde satır geçişlerini yapmanız lazım. Bir örnekle ne demek istediğimiz 
	daha iyi anlatabilirim.</p>
	<p> Şimdi aşağıdaki kodu, bir modül içine yazıp F5 ile çalıştıralım. Görüntü 
	aşağıdaki gibi olup, kullancının okuması açısından çok kolay değildir.</p>
	
	<pre class="brush:vb">
Sub satırgeçiş()
  a = InputBox("Müşteri segmenti için bir değer giriniz. Bireysel müşteriler için 1, Ticari müşteriler için 2, Kurumsal müşteriler için 3")
End Sub
	</pre>

<p>	<img src="/images/vba_interaktivite_satır1.jpg"></p>
<p>
	Şimdi bir de bu kod nasıl daha düzenli hale getirilir ona bakalım: Her cümle 
	ve seçenek arasına bir ifade koyarak. Bu ifade vbCrLf ifadesidir ve 
	cümleleri bir alt satıra taşır, bunun yerine vbCr veya vbLf veya vbNewLine 
	veya Chr(10) ifadeleri de 
	kullanılabilir. (Bunların dördü de Msgbox ve InputBox kullanımında aynı etkiye 
	sahiptir, ancak hücre içine birşey yazdırırken farklı etkilere sahiptir, 
	bunu deneyip görebilrisiniz.)</p>
	
		<pre class="brush:vb">
Sub satırgeçiş2()
a = InputBox("Müşteri segmenti için bir değer giriniz. " &amp; vbCrLf &amp; "Bireysel müşteriler için 1," &amp; vbCrLf &amp; "Ticari müşteriler için 2," &amp; vbCrLf &amp; "Kurumsal müşteriler için 3")
End Sub
	</pre>

	<p> 
	<img src="/images/vba_interaktivite_satır2.jpg"></p>
	<h3> 
	Kodlamacı dostu kodlama</h3>
	<p> 
	Şimdi yeri gelmişken bir de kullancı dostu olmakla ilgili değil ama 
	kodlamacı dostu olmakla ilgili bir notum olacak. Yine yukardaki kodu örnek 
	alalım, bu kod biz kodlamacılar için de okuması zor, çünkü VBE içinde kod 
	sağa doğru uzuyor, ama kodlamacı olarak benim bunu ekranda, scroolbarı sağa 
	sürüklemeden görebilmem lazım. Hadi gelin bunu 
	okunaklı hale getirelim.</p>
	<p>Yapacağımız şey basit, cümleyi nerden kesmek istiyorsak oraya bir 
			boşluk ve sonrasında bir alt çizgi(_) koymak. Buna<span class="keywordler"> Line Contination 
			Character</span> adı verilir.</p>
			
<pre class="brush:vb">
Sub satırgeçiş3()
a = InputBox("Müşteri segmenti için bir değer giriniz. " &amp; vbCrLf &amp;  _
"Bireysel müşteriler için 1," &amp;  vbCrLf &amp;  _
"Ticari müşteriler için 2," &amp;  vbCrLf &amp;  _
"Kurumsal müşteriler için 3")
End Sub
</pre>

	<p>	Görüldüğü gibi, kod şimdi bizim için de daha okunaklı hale geldi.</p>
	<p> 	Bunu yapmanın bir yolu daha var, o da metni parçalara ayırmak.</p>

<pre class="brush:vb">
Sub satırgeçiş4()
mesaj = "Müşteri segmenti için bir değer giriniz. " &amp; vbCrLf
mesaj = mesaj + "Bireysel müşteriler için 1," &amp; vbCrLf
mesaj = mesaj + "Ticari müşteriler için 2," &amp; vbCrLf
mesaj = mesaj + "Kurumsal müşteriler için 3"

a = InputBox(mesaj)

End Sub 
</pre>

	<p> 	NOT:mesaj = mesaj + ..... şeklinde sağduyuya aykırı gibi görünen 
	kısım kafanızı karıştırdıysa <a href="Temeller_Operatorler.aspx#selfcombine">
	buradan</a> detaylı bilgi edinebilirsiniz.</p>
	<p> 	Bu iki yöntemi sadece interaktivite sağlayan yerlerde değil başka 
	yerlerde de kullanacağız.</p>
	<p> 	Bu arada hemen iki yöntem arasındaki küçük farka da değinelim. İlk 
	yöntem yani _ yöntemi ile sadece metin birleştirme değil, içinde metin 
	bile olmayan tam bir VB kodunu da parçalara ayırabiliriz, amaç yine 
	aynı: Sağa doğru uzayan kodu&nbsp;tek bir ekranda tutmak. Aşağıdaki gibi.</p>
	
<pre class="brush:vb">
Sub blabla()
Cells.Find(What:="Volkan", After:=ActiveCell, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False).Activate
End Sub
</pre>

<p>İkinci yöntemin ise tek amacı uzun metinleri parçalara ayırmaktır. </p>
	</div>




</asp:Content>
