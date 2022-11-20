<%@ Page Title='Fonksiyonlar WorksheetFunction' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='7'></asp:Label></td></tr></table></div>

<h1>WorksheetFunction</h1>
<p> VBA'in en önemli fonksiyonlarından bir grup da WorksheetFunction class'ı 
içinde bulunan metodlardır. Bunlar, adı üzerinde, Excel içinde kullanılan 
fonksiyonlardır. Burada tüm fonksiyonlar bulunmamakla birlikte önemli bir kısmı 
dahil edilmiştir. 
Nelerin olduğunu, aşağıdaki gibi Worksheetfunction yazıp nokta koyduktan 
sonra intellisense aracılığıyla görebilirsiniz(veya F2'ye basıp object browserdan 
da arayabilirsiniz). Genelleme yapacak olursak VBA eşleniği olan fonksiyonlar 
burada yoktur. Now, Left, Mid gibi.</p>
	<p> 
	<img src="/images/vbawsf1.jpg"></p>
	<p> 
	Bunlar aslında UDF'in tersi gibi düşünülebilirler. Nasılki VBA aracılığıyla 
	Excel'e yeni fonksiyonlar kazandırabiliyorsak Excel aracılığyla da VBA'e ek 
	fonksiyonlar kazandırmış oluyoruz.</p>
	<p> 
	Mesela bir lookup işlemi yapmak istediğinizde tıpkı Excel içinde VLOOKUP 
	yazarmış gibi VBA içinden bu lookup işlemini yapabilirsiniz. <strong>Bunun sonucunu 
	bir hücrede görmek yerine bir değişkenin içine depolamış olursunuz,</strong> o kadar.</p>
	<p> 
	<strong>ÖNEMLİ NOT</strong>:Bazen WorksheetFunction yerine direkt 
	Application nesnesinin kullanıldığını görebilirsiniz. Bu bağlamda ikisi 
	özdeştir diyebiliriz, aşağıdaki örnekteki gibi. Ancak siz yine de 
	WorksheetFunction'ı kullanın, zira yapılan testler bunun %20 civarında daha hızlı 
	olduğunu söylüyor. Üstelik Application'lı versiyonda Intellisense de 
	çıkmamaktadır. 
	Hata yakalama bağlamında da farkları var ancak o detaya girmeyeceğim, arzu 
	eden
	<a href="https://www.mrexcel.com/forum/excel-questions/584913-application-vs-application-worksheetfunction.html">
	buradan</a> bakabilir.</p>
	<pre class="brush:vb">x=Application.Sum(Range("A1:A10"))
x=WorksheetFunction.Sum(Range("A1:A10"))</pre>
	<h2> 
	Örnekler</h2>
	<h3> 
	Örnek 1</h3>
	<p> 
	İlk örneğimizde sistemden çektiğimde 12 haneden küçük portföy kodlarını 12 
	haneye tamalayan bir UDF var. Bu örnekte en başa gerekn miktarda 0 
	konulmaktadır. Bunun Excelde yapmak için kendisine çok benzeyen şu formülü 
	girmem gerekirdi. </p>
	<pre class="formul">=REPT("0",12-LEN(A2)) &amp; A2</pre>
	<pre class="brush:vb">Function portföy12(pk As Range)
  portföy12 = WorksheetFunction.Rept("0", 12 - Len(pk)) &amp; pk
End Function</pre>
	<p>Hangisini yazmak daha kolay, Excel formülünü mü, UDF'i mi? Böylece bir 
	önceki bölüme ithafen UDF'lerin gücünü de anmış olalım.</p>
	<p>Bu arada bunu alternatif olarak şöyle de yapabilirdik ancak 
	WorksheetFunction'ı örneklemek adına böyle yaptık.</p>
	<pre class="brush:vb">Function portföy12(pk As Range)
  portföy12 = String(12 - Len(pk),"0") &amp; pk
End Function</pre>
	<p>Ama diyelim String yöntemi VBA'de yok, siz de WorksheetFunction'ı 
	bilmiyorsunuz. Bu durumda bu UDF'i şöyle 
	yazardık.</p>
	<pre class="brush:vb">Function portfoy12(pk As Range)
ilave = 12 - Len(pk)

For i = 1 To ilave
    ek = ek & 0
Next i

portfoy12 = ek & pk
End Function	</pre>
	<p>Gördüğünüz gibi WorksheetFunction bizi gereksiz bir For döngüsü yazmaktan 
	kurtarmış olur(Tabi bu örnekte String alternatifimizin olduğunu birkez daha 
	altını çizelim, ama her zaman böyle alternatifler olmaz)</p>
	<h3> 
	Örnek 2</h3>
	<p>İkinci örneğimzde vlookup kullanımı var. Bu örnekte de 'aranan' isimli 
	bir değeri 'table' isimli bir range içinde arıyoruz. Bunu normal VBA ile 
	yapmak için <a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">Dictionary</a> 
	tanımlamak gerekirdi, ki bu da kodlarımızı oldukça uzatırdı.</p>
	<pre class="brush:vb">.....
sonuc = WorksheetFunction.VLookup(aranan, table, 1, False)</pre>
	<p>Gördüğünüz gibi Excel'in içinde herhangi bir hücreye Vlookup formülü 
	yazıdrmış olmuyoruz, tamamen VBA tarafındayız ve <strong>sanki </strong>
	Excel'de Vlookup yapmış gibiyiz ve sonucunu da bir değişkene atıyoruz. 
	Excel'de bir hücreye formül yazdırmak için <strong>Range</strong> nesnesinin
	<strong>Formula</strong> property'si kullanılır. Bu ayrım, özellikle ilk 
	başlarda kafa karıştırıcı olabilir. Birkaç kez kullandıkça farkı 
	anlayacaksınız. </p>
	<h3> 
	Örnek 3&nbsp;</h3>
	<p> 
	Aşağıda ise çifte vlookup yapmayı sağlayan bir UDF var, burada 
	
	<strong>COUNTIF</strong> fonksiyonunun kullanımı görüyorsunuz.</p>
	<pre class="brush:vb">Function Çiftevlookup(alan As Range, sütun As Long, İlk_kriter, İkinci_kriter)

Dim rCheck As Range, bFound As Boolean, lLoop As Long
On Error Resume Next

Set rCheck = alan.Columns(1).Cells(1, 1)
For lLoop = 1 To WorksheetFunction.CountIf(alan.Columns(1), İlk_kriter)

  Set rCheck = alan.Columns(1).Find(İlk_kriter, rCheck, xlValues, xlWhole, xlNext, xlRows, False)
  If UCase(rCheck(1, 2)) = UCase(İkinci_kriter) Then

  bFound = True
  Exit For
  End If

Next lLoop

End With

If bFound = True Then
  Çiftevlookup= rCheck(1, sütun)
Else
  Çiftevlookup= "#N/A"
End If

End Function</pre>
	<h3> 
	Örnek 4</h3>
	<p>Yine bir önceki UDF bölümünde gördüğümüz bir fonksiyonda <strong>LARGE</strong> 
	ve <strong>SMALL</strong> 
	fonksiyonlarının kullanımına şahit oluyoruz.</p>
	<pre class="brush:vb">
Function uçhariçort(alan As Range, Uç As Variant)
Dim aratoplam As Double
Dim enbüyükler As Double
Dim enküçükler As Double

For i = 1 To Uç
    enbüyükler = enbüyükler + WorksheetFunction.Large(alan, i)
Next i
    
For i = 1 To Uç
    enküçkler = enküçkler + WorksheetFunction.Small(alan, i)
Next i
     
aratoplam = WorksheetFunction.Sum(alan) - enbüyükler - enküçkler
uçhariçort = aratoplam / (alan.Count - Uç * 2)

End Function	</pre>
	<h2>
	Evaluate</h2>
	<p>
	WorksheetFunction'a benzer bir de Evaluate metodu vardır. Bazı durumlarda 
	WorksheetFunction'ın&nbsp;özdeşi olup buna göre yazımı daha kolay olduğu 
	için tercih edilebilir. Mesela aşağıdaki iki ifade özdeştir.</p>
	<pre class="brush:vb">WorksheetFunction.Sum(Range("A1:A10"))
Evaluate("SUM(A1:A10)")</pre>
	<p>
	Farkettiyseniz bunda Range yok, aslında tırnak içindeki formül tamamen Excel 
	ortamında yazdığımız formül gibi, içerde hiçbir VBA terimi yok.</p>
	<p>
	Ama Evaluate sadece bu kadar iş yapmıyor, onun üstün basan yönleri de var. 
	Bence bunların en önemlisi içeriğindeki formülü dizi formülü olarak 
	üretebilmesidir. Ve bu da Evaluate'i dizi formülü üreten UDF'ler yazmada 
	oldukça kullanışlı kılar.
	<a href="http://www.decisionmodels.com/calcsecretsh.htm">Burada</a> bu metod 
	ile ilgili çok daha detaylı bilgi bulabilirsiniz.</p>
	<p>
	Aşağıdaki UDF ile belirli bir alandaki En Büyük/Küçük X değerin toplamını 
	alırız. X=True ise En büyük, False ise En küçük x değere bakılır.</p>
	<pre class="brush:vb">
Function EnXTopla(alan As Range, N As Long, Optional x As Boolean) As Single
Dim strAddress As String


On Error Resume Next
strAddress = alan.Address
    If x = False Then
        EnXTopla = Evaluate("=SUMPRODUCT((" _
            & strAddress & ">=LARGE(" & strAddress & "," & N & "))*(" & strAddress & "))")
    Else
       EnXTopla = Evaluate("=SUMPRODUCT((" _
            & strAddress & "<=SMALL(" & strAddress & "," & N & "))*(" & strAddress & "))")
    End If
End Function	</pre>
	</asp:Content>
