<%@ Page Title='KosulifadeleriveDonguler ForDonguleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Koşul ifadeleri ve Döngüler'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>For...Next... Döngüleri</h1>
	<p>VBA dünyasının en temel döngü tipi For döngüleridir. Bir önceki koşullu 
	yapılar konusunda söylediğimiz gibi döngüler de programlamanın olmazsa olmazlarıdır. 
	Onlar olmasaydı belki 3 satırlık kodu onlarca hatta yüzlerce satır şeklinde 
	yazmamız gerekirdi.</p>
	<p>For döngülerinin 2 tipi vardır:</p>
	<ul>
		<li>Basit For Next döngüsü</li>
		<li>For Each döngüleri</li>
	</ul>
	<p>İlk olarak basit For döngüsü ile başlayalım</p>
	<h2>For...Next...</h2>
	<p>Genel kullanım şekli şöyledir:</p>
	<pre class="brush:vb">For i=başlangıçsayısı to bitişsayısı
    'i'ye bağlı veya i'den bağımsız olarak şunu yap
Next i</pre>
	<p>Gördüğünüz gibi, sayaç olarak kullandığımız i değişkenini döngü 
	içinde de parametrik olarak kullanabiliriz veya kullanmayabiliriz. İkisine de 
	örnek verelim. Önce i'yi döngü içinde parametrik olarak kullanacağımız örnek 
	olsun. Bu örnekte, A1'den A10'a kadar olan hücrelere sırayla i'nin kendisini yazmış oluruz.</p>
	<pre class="brush:vb">For i=1 to 10
    Cells(i,1)=i
Next i</pre>
	<p>Şimdi de parametrik olmayan örneğe bakalım, bunda ise Excele 10 kez bip 
	sesi çıkarttırıyoruz.</p>
	<pre class="brush:vb">For i=1 to 10
    Beep
Next i</pre>
	<p>i sayacı genelde 0 veya pozitif(ki bu da genelde 1dir) bir sayı olmakla 
	birlikte negatif bir sayı da olabilir, ama bunun pratikte pek kullanıldığını 
	görmedim.</p>
	<p>Sayacımız normalde 1er 1er artar, ancak istersek <span class="keywordler">
	Step</span> ifadesi ise 2'şer, 3'er de arttırabiliriz. Hatta sayacı geriye 
	doğru da saydırabiliriz, ki bunun için Step ifadesinden sonra negatif bir 
	sayı gelir.</p>
	<pre class="brush:vb">For i = 10 To 1 Step -2
   Cells(i / 2, 1) = i
Next i</pre>
	<p>Tüm sayfalarda gezindiğimiz bir kod:</p>
	<pre class="brush:vb">For i = 1 To Sheets.Count
   Worksheets(i).Select
   'diğer kodlar
Next i</pre>
	<p>Şimdi bir başka örnekte de ilk hücreden en alt hücreye doğru ilerleyelim.</p>
	<pre class="brush:vb">For i = 2 To Cells(1, 1).End(xlDown).Row
   Cells(i,1).Select
   'Diğer kodlar buraya
Next i</pre>
<h3>İçiçe For Döngüleri</h3>
	<p>For döngülerini de tıpkı koşullu yapılarda olduğu gibi iç içe geçmiş 
	şekilde kullanabiliriz. Aşağıdaki örnekte boş olan hücreleri sarıya 
	boyayalım ve sayısını bulalım:</p>
	<p><img src="/images/vbaforicice.jpg"></p>
	<p>Öncelikle şunu belirtmekte fayda var. Tek olsun içiçe olsun tüm döngülerde 
	For satırını yazdıktan sonra hemen onunu bitişi olan Next satırını da yazın, 
	aradaki kodları da girintili yazın. İçiçe For Next olacaksa bu ikinci For'u 
	da yine girintli yazın. Şimdi kodumuza bakalım.</p>
	<pre class="brush:vb">
Sub bosluksay()
Dim adet As Integer
For i = 1 To Range("A1").End(xlToRight).Column
    For k = 1 To Range("A1").End(xlDown).Row
        If IsEmpty(Cells(k, i)) Then
            Cells(k, i).Interior.Color = vbYellow
            adet = adet + 1
        End If
    Next k
Next i

MsgBox "toplamda " & adet & " adet boş hücre var"
End Sub	</pre>
	<p>ve sonuç:</p>
	<p><img src="/images/vbaforicice2.jpg"></p>

	<h2>For Each...Next...</h2>
	<p>For Each yapısı bir obje grubu ve bir dizi(veya dizimsi) içindeki elemanlar içinde gezinmek için kullanılır.</p>
	
<pre class="brush:vb">
For Each nesne in nesnegrubu
  'nesneyle ilgili bir işlem 
Next nesne </pre>
</br>
<pre class="brush:vb">
For Each elaman in dizi
  'dizi elemanıyla ilgili işlem
Next eleman </pre>

	<p>Diziler hakkında ön bilgi almak için
	<a href="DizilerveDizimsiYapilar_DizilerArray.aspx">buraya</a> tıklayın.</p>
	<p>Şimdi yukardaki boşluk saydırma örneğini For Each ile nasıl yapacağımıza 
	bir bakalım.</p>
	<pre class="brush:vb">
Sub bosluksay2()
Dim adet As Integer
Dim hucre As Range, alan As Range

Set alan = Range("A1").CurrentRegion

For Each hucre In alan
    If IsEmpty(hucre) Then
        hucre.Interior.Color = vbYellow
        adet = adet + 1
    End If
Next hucre

MsgBox "toplamda " & adet & " adet boş hücre var"
End Sub
</pre>
	<p>Bence For Each ile sanki daha kolay gibi görünüyor. O yüzden böyle bir 
	görevde ben ForEach'i kullanırdım. Ancak eğer döngü içinde i veya k parametrelerinden 
	birini kullanarak da bir işlem yapmam gerekseydi, o zaman basit For döngüsü 
	kullanırdım, gerçi yine For Each kullandığımızda hücrenin row ve column 
	özellikleri ile i ve k'yı dolaylı olarak elde edebilirim ama basit For'da 
	bunlar zaten elimde olacaklardı. Kısacası duruma göre hangisini 
	kullanacağınıza siz karar vereceksiniz. Basit For'un esnek başlangıç ve bitiş 
	değerleri ile hareket yönündeki esnekliği avantaj sayılabilirlken, For Each 
	için daha hızlı olmasını avantaj olarak vurgulayabiliriz.</p>
	<p>For Each yapısını, Range'teki hücrelere ek olarak, Workbooktaki sayfalar 
	ve tüm workbooklar arasında dolaşma şeklinde de bol miktarda kullanıyoruz.</p>
<pre class="brush:vb">
'workbooklarda dolaşma
Dim wb As Workbook
For Each wb In Workbooks
    MsgBox wb.FullName
Next wb
 
'worksheetlerde dolaşma
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    MsgBox ws.Name
Next ws
</pre>
	<p>Yine dizi veya dizimsilerin elemanları içinde gezinirken de her iki yapı 
	kullanılabilir. Basit For'da <strong>1 to elemansayısı</strong> yapısı 
	kullanılırken ForEach'te <strong>each eleman in dizi </strong>şeklinde 
	kullanırız. Şimdi kısa bir örnek de dizilerle ilgili yapalım, daha detaylı 
	örnekleri <a href="DizilerveDizimsiYapilar_Konular.aspx">Diziler </a>bölümüne bırakalım.</p>
	<p>Mesela sabit bir bölge kodu listeniz olsun ve bunların her birini bir 
	diziye atayalım, sonra da bu bölge kodunu kullanarak bir işlem yapalım.</p>
<pre class="brush:vb">
Sub forarray()
Dim blg() As Variant

blg = Array(5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009, 5010)
For Each b In blg
   Workbooks.Open ("C:\bölge raporları\" &amp; b &amp; "-netice raporu.xlsx")
   'diğer işlemler
Next b
End Sub</pre>
	<p>Akılda bulundurulması gereken önemli bir husus, ForeEach kullanıldığında read-only bir özellik gösterir. Yani bu yöntemle dizi 
	elemanlarını değiştiremezsiniz. Elamanları değiştirmek istiyorsanız basit For 
	döngüsü kullanmanız lazım.
	<a href="http://excelmacromastery.com/vba-for-loop/">Exlcemacromastery</a> 
	sitesinden aldığım bu iki örnekte fark çok açıkça anlatılmış.</p>
	<pre class="brush:vb">
Sub UseForEach()

    ' Diziyi yaratıp 3 değer ekliyoruz
    Dim arr() As Variant
    arr = Array("A", "B", "C")

    Dim s As Variant
    For Each s In arr
        ' s'nin atadığı değeri değiştirmeye çalışıyoruz
        s = "Z"
    Next s

    ' Ama değişmemiş olduğuğnu  görüyoruz
    For Each s In arr
        Debug.Print s
    Next s

End Sub	</pre>
	<p>Basit for ile değişimi görelim</p>
	<pre class="brush:vb">
Sub UsingForWithArray()

    Dim arr() As Variant
    arr = Array("A", "B", "C")

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = "Z"
    Next

    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next

End Sub</pre>
	<p>
	Son bir örnek daha yapalım. Bu sefer, toplu bir goalseek(hedef ara) işlemi 
	yapalım.</p>
	<pre class="brush:vb">
Sub toplugoalseek()
Dim h As Range
For Each h In Range("E8:E2294")
   h.Select
   h.goalseek Goal:=h.Offset(0, 4).Value2, ChangingCell:=h.Offset(0, 2)
Next h

End Sub	</pre>
	<h2>Döngüden Çıkış</h2>
	<p>Yukarıdaki kod örneklerinden de gördüğünüz üzere, For döngüleri 
	genellikle kaç kez çalıştırılacağını bildiğiniz durumlarda kullanılır. 
	Ör: <strong>For i=1 to 10</strong> dersek, 10 kez çalışacağını biliriz veya 
	<strong>For Each item in 
	dizi </strong>dersek dizideki tüm elemanlar için çalışacağını biliriz. Tabiki üst 
	limitin değişken olduğu bazı durumlarda, o limite başka VBA kodları ile de 
	ulaşabilir ve yine For döngülerini kullanabiliriz, mesela en alt satırın kaç 
	olduğunu bilmiyoruzdur ama Range("A1").End(xlDown).Row ile bunu bilip 
	döngüyü buraya kadar götürebiliriz, gerçi bu tür durumlarda diğer döngü 
	türlerini kullanmak daha kullanışlıdır.</p>
	<p>Ancak bazen, döngüyü erken terketmeniz gerekebilir. Yani döngüye girdiniz diye 
	ille sonuna kadar gitmeniz gerekmiyor. Bunun için <span class="keywordler">GoTo</span> ifadesi 
	kullanılabileceği gibi, <span class="keywordler">Exit For</span> yapısı da kullanılabilir. GoTo kullanımında doğrudan 
	belli bir etikete yönledirilirken Exit For ile döngünün hemen arkasından devam edilir.</p>
	<p>Mesela yukardaki örneğimizi, boşluğa ilk rastlanılan hücrenin 
	adresini verecek şekilide değiştrelim.</p>
	<pre class="brush:vb">Sub bosluksay3()
Dim a As String
Dim hucre As Range, alan As Range

Set alan = Range("A1").CurrentRegion

For Each hucre In alan
    If IsEmpty(hucre) Then
        a = hucre.Address
        Exit For
    End If
Next hucre

If a <> vbNullString Then
    MsgBox "ilk olarak " & a & " adresinde boşluğa rastlanmıştır"
Else
    MsgBox "herhangi bir boş hücre bulunmamaktadır"
End If
End Sub</pre>
	<p>&nbsp;</p>

</asp:Content>
