<%@ Page Title='VBA için UDF' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='6'></asp:Label></td></tr></table></div><h1>VBA için UDF</h1>
	<h2 class='baslik'>Giriş</h2>
	<div class='konu'>
<p>
	Excel'de kullanım için hazırladığımız UDF'lerden başka VBA prosedürleri 
	içinde çalışıp bir sonuç döndüren fonksiyonlar da vardır. Bunlar da tıpkı 
	VBA'in yerel fonksiyonları gibidirler. Genelde belirli bir(bazen birkaç) 
	sonuç döndürmek üzere hazırlanırlar. Ender olarak bazı kaynaklarda sonuç 
	dödürmeyen versiyonların kullanıldığını da görebilirsiniz ama bence bu 
	tamamen yanış kullanımdır. Sonuç döndürmeyen bir iş istiyorsak bunu Function 
	olarak değil Sub olarak hazırlamalıyız. </p>
	<p>Aşağıda bu yanlış kullanıma bir örnek bulunmaktadır. Bu, doğru çalışan 
	tamamen düzgün bir koddur, ancak dediğim gibi Functionların amacı bu değildir, 
	olmamalıdır.</p>
	<pre class="brush:vb">Sub YanlışFuncOrnek()
Call YanlışFunc
End Sub

Function YanlışFunc()
[A1] = 12
End Function</pre>
	<p>Fonksiyonların amacı nedir diyecek olursak, iki temel amacı vardır.</p>
	<ul>
		<li>Ana Sub prosedürünüzün çok uzaması durumunda bunu belli 
		yerlerde kesip fonksiyon olarak yazmak ve ana koddan bu fonksiyonu 
		çağırmak</li>
		<li>Bir diğeri de, belirli bir işi farklı zamanlarda, farklı kodlar 
		içinde sürekli yapmak durumunda kalıyorsanız, bunu bir 
		kez fonksiyon olarak yazarsınız, sonra her yerden bu fonksiyonu 
		çağırırsınız. Böylece gereksiz tekrardan kurtulmuş olursunuz. Üstelik 
		fonksiyonda küçük bir değişklik gerekse bile sadece bir kere yapılması 
		yeterli olacaktır.</li>
	</ul>
	<p>Son olarak belirtmek istediğim bir husus var, VBA kullanımı için 
	yazdığımız functionlara genelde kimse UDF demez, UDF denince akla genelde 
	Excel için yazdığımız UDF'ler gelir, ama aslında teknik olarak bakıldığında 
	VBA içi kullanım amacıyla yazdığımız functionlar da UDF'tir. Zira VBA'de yerel olarak gelmediği 
	için biz bunları kendimiz tanımlarız.</p>

<h3>Tanımlama</h3>
		<p>Bunların tanımlanması da tıpkı Excel UDF tanımı gibidir.</p>
		<pre class="brush:vb">Function fonksiyonadı(Pamaratre1 As datatipi,...) As DönüşTipi
.......
End Function</pre>
		<p>Sadece kullanırken Excel içinde değil de VBA içinde 
		kullanıyoruz. Bu arada Excel'den böyle bir fonksiyonu yazmaya 
		çalıştığınızda yazabilirsiniz, buna bi engel yoktur. O yüzden bunların 
		görünmesini istemiyorsanız <strong>Private</strong> olarak tanımlamanız gerekir. 
		Aşağıdaki örneklerden bunu görebilirsiniz.</p>
		<p><img src="/images/vbaudfforvba1.jpg"></p>
		<p>Aşağıda görüldüğü üzere Excelde bir hücreye <strong>=UDF</strong> yazınca sadece 
		UDFExcel çıktı.</p>
		<p><img src="/images/vbaudfforvba2.jpg"></p>
		<p>Tabi ilgili fonksiyonu Private tanımlamanın da bir dezavanatajı var, onu
		<a href="Temeller_Terminoloji.aspx#prosedurerisim">prosedürlerin erişim seviyesi bölümünde</a> 
		görmüştük. Özetle bu fonksiyonlara sadece bulunduğu modülden erişilebilir. 
		Eğer bu bir sıkıntı olmayacaksa private tanımlayın, sıkıntı olacaksa 
		public tanımlayın. (Tabi bu arada bunları bir Add-in içine yazdığınız 
		senaryoya göre söylüyorum. Add-in dışında başka bir yere mesela Personal.xlsb içine 
		yazılanlarda zaten fonksiyon ismini Excele yazınca hemen çıkmıyordu, 
		başında <strong>Personal.xlsb!</strong> olması gerekiyordu veya Insert 
		function diyip seçmek gerekiyordu) </p>
		<p>Zaten genel tavsiyem şudur: Özellikle 
		Excel UDF'lerinizi başkalarına da dağıtıyorsanız UDF.xlam dosyası içine 
		sadece Excel UDFlerini koyun, kendi kullanımınız için oluşturduğunuz 
		VBA UDF'lerini ise Personal.xlsb içine koyun.</p>
		<h3>Kullanım</h3>
		<p>Şimdi bir örnekle konuyu pekiştirelim.</p>
		<p>Diyelimki ana makronuzda(veya birçok makronuzda) öyle bir yer 
		geliyor ki, o anda ilgili alanda filtre uygulanmış mı uygulanmamış mı 
		bunu kontrol etmek ve sonrasında duruma göre de bir işlem yaparak 
		ilerlemek istiyorsunuz. Bunun için aşağıdaki gibi bir Function yazarız. 
		Bunu 
		Personal.xlsb içine yazdıyorum.</p>
		<pre class="brush:vb">Function filter_kontrol(ws As Worksheet) As Byte
If ws.AutoFilterMode = True Then
    If ws.FilterMode = False Then
        filter_kontrol = 1 'filtre açık ama criter yok
    Else
        filter_kontrol = 2 'filtre açık ve criter var
    End If
Else
        filter_kontrol = 0
End If

End Function


Sub filtrekullan()
'örnek kullanım şekli
    If Application.Run("PERSONAL.xlsb!filter_kontrol", ActiveSheet) = 2 Then 'filtre açık ve criter var
        ActiveSheet.ShowAllData
    ElseIf Application.Run("PERSONAL.xlsb!filter_kontrol", ActiveSheet) = 0 Then 'filtre uygulanmamış
        Selection.AutoFilter
    'diğer duurmlarda yani 1, yani filtre açık ama criter yok, bişey yapmaya gerek yok
    End If
End Sub</pre>
		<h3>Dönüş Değeri</h3>
		<p>Genelde fonksiyonların sadece bir adet dönüş değeri olur. Ancak bazen 
		bir fonksiyonu çağırdığımızda birden fazla dönüş değeri isteyebiliriz. 
		Bunun için çeşitli alternatifler olmakla birlikte ben ikisinden 
		bahsedceğim. Aslında ikisinden de farklı yerlede bahsetmiştiK. O yüzden 
		sadece link veriyor olacağım.</p>
		<ul>
			<li><a href="Temeller_Birazdahaterminoloji.aspx#byvalref">ByRef </a>
			ile çoklu değer döndürme</li>
			<li><a href="DizilerveDizimsiYapilar_DizilerArray.aspx#dizisonuc">Diziler</a> ve
		<a href="DizilerveDizimsiYapilar_Collectionlar.aspx#colsonuc">Collectionlar</a> aracılığı ile çoklu değer döndürme</li>
		</ul>
	<p>Konuyu tamamlamak adına İleriTerminoloji sayfasındaki
	<a href="Temeller_Birazdahaterminoloji.aspx#paramarg">Argüman ve Parametre</a> 
	ile 
	<a href="Temeller_Birazdahaterminoloji.aspx#prosedurerisim">Prosedürlere 
	Erişim</a> maddelerini gözden geçirmenizi tavsiye ederim.&nbsp;</p>
		<h3>Yinelemeli(Recursive) fonksiyonlar</h3>
		<p>Bir fonksiyonun içinde kendisine başvuru yapmamız da mümkündür. Buna 
		recursive başvuru denir. Tabi bunu sonsuza kadar değil de belirli bir 
		şart sağlanana(mesela bir limite ulaşmak gibi) kadar kurgulamak gerekir, 
		yoksa kodumuz kısır döngüye girer.</p>
		<p>Recursive fonksiyonlara klasik örnek matematikteki faktoriyel 
		hesabıdır. Aşağıdaki örnekte bu hesabı bulabilirsiniz. Aldığı parametre 1 
		olana kadar kendisini -1 değeriyle çağırıp parametre ile çarpıyor. 
		Parametrenin değeri 1'e ulaştığında(indiğinde) yineleme sonlanmış 
		oluyor.</p>
		<pre class="brush:vb">
Function faktoriyel(ByVal n As Integer) As Integer
  If n &lt;= 1 Then
    faktoriyel= 1
  Else
    faktoriyel= faktoriyel(n - 1) * n
  End If
End Function
'---------------
Sub faktoriyel_yaz()
   Debug.Print faktoriyel(5) '120
End Sub		</pre>
		<p>
		Bu fonksiyonun ele alınışı aşama aşama şöyledir.</p>
		<ul>
			<li>Önce 5 parametresini verdik, sonuç=faktoriyel(4)*5</li>
			<li>Şimdi parametre olarak 4 gitmiş oldu, sonuç=<strong>faktoriyel(3)*4</strong>*5(Koyu 
			kısım aslında bir üst satırın açılmış hali</li>
			<li>Sonra 3 gider, sonuç=<strong>faktoriyel(2)*3</strong>*4*5(Koyu 
			kısım yine bir üsttekinin açılımı)</li>
			<li>Sonra 2, sonuç=<strong>faktoriyel(1)*2</strong>*3*4*5</li>
			<li>Son olarak 1 gider, sonuç<strong>=faktoriyel(1)*2</strong>*3*4*5</li>
			<li>Nihai sonuç=1*2*3*4*5=120</li>
		</ul>
		<p>Bunun dışında parent-child tarzı oluşumlarda(Ör:klasör-dosya veya 
		HTML tag'i ve alt tag'i)&nbsp; da çok sık kullanılır. Bu kullanım 
		şekline bir örnek de aşağıda bulunmaktadır. (Dosya ve klasörlerle 
		çalışmak için <a href="Dosyaislemleri_DosyaveKlasorerisimi.aspx">buraya</a> 
		tıklayınız)</p>
		<pre class="brush:vb">
'ana prosedür
Sub recursive_fulldosya()
    Dim fso As New Scripting.FileSystemObject
    Dim anaklasorStr As String
    anaklasorStr = "C:\windows"
    Recursiveİlerle fso.GetFolder(anaklasorStr)
End Sub
 
'recursive prosedür
Sub Recursiveİlerle(kls As Variant) 'variant çünkü ilk girereken Folder sonra Folders olacak
    Dim altKlasorler As Variant
    Dim dosya As file
    Dim i As Integer
    
    On Error Resume Next 'erişim izni olmayan yerlerde hata almasın diye
    For Each altKlasorler In kls.SubFolders
        Recursiveİlerle altKlasorler 'burada recursive olarak başvuru var
    Next
    i = 1
     
    For Each dosya In kls.Files
        ActiveCell(i, 1).Value = dosya.ParentFolder
        ActiveCell(i, 1).Offset(0, 1).Value = dosya.Name
        ActiveCell(i, 1).Offset(0, 2).Value = dosya.Size
        i = i + 1
    Next
End Sub	</pre>
	</div>
		<h2 class="baslik">Çeşitli Örnekler</h2>
		<div class="konu">
		<h4 class="baslik">Excel hücre grubunu mail body'sine koymak</h4>
		<div class="konu">
		<p>Bir başka örnek de meşhur RonDeBruin'in belirli bir Excel hücre 
		grubunu mail bodysi haline getiren fonksiyondur, efsanedir, hayat 
		kurtarıcıdır. Kod aşağıdaki gibi olup orjinaline
		<a href="https://www.rondebruin.nl/win/s1/outlook/bmail2.htm">buradan</a> 
		ulaşabilirsiniz.</p>
		<pre class="brush:vb">
Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2013
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
 
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    'Close TempWB
    TempWB.Close savechanges:=False
 
    'Delete the htm file we used in this function
    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function</pre>
		<p>
		Örnek kullanımı ise şöyledir. Uzun olmaması adına tüm kodu buraya 
		koymadım, tam örnek koda
		<a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx">outlook 
		programlama</a> bölümünde değineceğiz.</p>
		<pre class="brush:vb">
....
Set rng = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
....
....
.htmlbody = "Değerli MİY'imiz," & Chr(14) & Chr(14)
.....
.....
.htmlbody = .htmlbody + Application.Run("PERSONAL.xlsb!RangetoHTML", rng) & vbCrLf & Chr(14)
.......	</pre>
</div>
		<h4 class="baslik">İlk visible alan ve sonrasını seçmek</h4>
		<div class="konu">
		<pre class="brush:vb">
Function ilkvisiblesec(erim As Range) As Range
    Set ilkvisiblesec = erim.Offset(1, 0).SpecialCells(xlCellTypeVisible).Cells(1)
End Function

Function ilkvisiblesonrasıalansec(erim As Range) As Range
    Dim ilk As Range
    Dim son As Range
    Dim n As Integer, r As Integer
    
    n = erim.Columns.Count
    r = erim.SpecialCells(xlCellTypeVisible).Cells.Count / n - 1 'tek satırlık bir alan olup olmadığını kontrol etmek için
    
    Set ilk = erim.Offset(1, 0).Resize(erim.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Cells(1, 1) 'bu kısım ilk görünen hücreyi verir
    Set son = ilk.Offset(0, n - 1)
    Set ilktoright = Range(ilk, son)
    
    If r > 2 Then
        Set ilkvisiblesonrasıalansec = Range(ilktoright, ilktoright.End(xlDown))
    Else
        Set ilkvisiblesonrasıalansec = ilktoright
    End If
End Function
</pre>
</div>
		</div>
</asp:Content>
