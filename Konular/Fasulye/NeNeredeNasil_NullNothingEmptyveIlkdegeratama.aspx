<%@ Page Title='NeNeredeNasil NullNothingEmptyveIlkdegeratama' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Fasulye'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Ne Nerede Nasıl'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushCSharp.js" ) %>"> </script>
<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushSql.js" ) %>"> </script>
<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushJScript.js" ) %>"> </script>

<h1>Null, Nothing, Empty</h1>
<p>Bu konu üzerine söylenecek şey miktarı oldukça fazla olduğu için burada platformları farklı kutulara koymak durumunuda kaldım. Her bir platform için ayrı ayrı tıklayıp bakabilirsiniz.
</p>
<p>Farklı platformlarda <span class=" keywordler">Empty, Null, Nothing, ZLS)</span> gibi birbirine çok benzeyen kavramlar vardır. Tahmin edeceğiniz üzere bunları sorgulama şekli de farklılık göstermektedir. Bunlar arasındaki farkı anlamak önemlidir, aksi halde hatalarla karşılaşmak kaçınılmaz olacaktır. </p>

<h2 class='baslik'>Excel</h2>
<div class='konu'>
<p>
Excel'de bu kavramlardan sadece Blank(Boş) ve ZLS mevcuttur. Boş değerler <span class=" keywordler">ISBLANK(A2)</span> ile, ZLS'ler ise 
	<span class=" keywordler">IF(A2="";TRUE;FALSE)</span>  veya 
	<span class=" keywordler">IF(LEN(A2)=0;TRUE;FALSE)</span> formülüyle sorgulanabilir. Örnek bir gösterim ve sorgulama şekli aşağıdaki gibidir. Görüldüğü gibi sadece hiçbirşey girilmeyen yani boş olan hücre için ISBLANK formülü TRUE döndürmüştür, ancak ZLS sorgusu hem boş hücre için hem de ="" değeri için TRUE dönmüştür.</p>

<img src="/images/NNNNullEmpty.jpg" class="zoomla" alt="Excell blank ZLS"/>
</div>

<h2 class='baslik'>VBA</h2>
<div class='konu'>
<h2>Empty</h2>
<p>Konu programlamaya geldiğinde karşımıza <span class=" keywordler">null</span> ve <span class=" keywordler">Nothing</span> kavramları da çıkar. VBA'da değer atanmamış tüm değişkenlere otomatik olarak default değerler atanır. Bu, 
numerik değişkenler için 0, Stringler (metinler) için ZLS, Object(nesne) tipli değişkenler için Nothing ve Variant değişkenler içinse Empty'dir. </p>

<p><span class=" keywordler">Empty,</span> başlangıç değeri atanmamış demektir. <span class=" keywordler">Null</span> ise geçerli bir data içermeyen değişken demektir. <span class=" keywordler">Empty</span>, Variant alttipi olarak karşımıza çıkar. İstenirse bir değişkene başlangıç anında da <span class=" keywordler">Empty</span> değeri verilebilir. </p>

<p>VBA'de yerel değişken(metodların elemanları) ve field(class seviyesindeki global eleman) ayrımı olmadığı için makro çalıştığında tüm değişkenler default değerlerine otomatik atanır. (Vb.Net ve C# gibi dillerde ise yerel değişkenlere mutlaka değer atanması lazım, fieldlar ise otomatik default değerlerine atanır.)
</p> 

<p>Başlangıç değeri atanmamış veya bilinçli şekilde Empty değeri atanmış bir değişkeni yakalamak için <span class=" keywordler">IsEmpty()</span> fonksiyonunu kullanırız. Bununla birlikte bir değişkenin 
Empty olması çok karşılaşılan bir durum olmayacak, daha çok bir hücre değerinin içinin boş olup olmadığını sorgulayacağız. 
Aslında bir hücrenin içinin boş olup olmadığını sorgularken Range objesinin 
default değeri olan Value'yu sorgulamış oluyoruz. Yani <span class="keywordler">
If IsEmpty(Range("A1")</span> ile <span class="keywordler">If 
IsEmpty(Range("A1").Value)</span> aynı şeydir, ve Value da Variant değer döndürdüğü 
için <span class="keywordler">IsEmpty</span> ile sorguladığımızda gerçekten 
boşsa True döndürür. Bu sorgulamayı <span class=" keywordler">Len </span>metodu ile de yapabiliriz.</p>

<pre class="brush:vb">
If Len(degisken)=0 Then 
'veya excel hücresi için
If Len(Range("A1"))=0 Then
</pre>

<p><span class=" keywordler">ZLS</span> tıpkı Excelde olduğu gibi "" şeklinde ifade edilir. 
Bu bir stringtir ve sıfır uzunluktadır. Çoğu durumda <span class=" keywordler">vbNullString </span> ile ZLS, VBA tarafından aynı şekilde yorumlanır. vbNullString tam bir string olmamakla birlikte birçok durumda ZLS yerine kullanılabilir ve hatta kullanılmalıdır da çünkü ZLS'ye göre performans açısından daha verimlidir, özellikle büyük bir döngü içinde sürekli bir "" ataması olacaksa. 
Çünkü ZLS, bellekte 6 byte yer kaplarken <span class=" keywordler">vbnullstring</span> 
ise ilave yer kaplamaz, zira <span class=" keywordler">vbnullstring</span> bir constant olup zaten VBA tarafından baştan yaratılmıştır ve yeniden yaratımına gerek 
yoktur.</p> 
<p>Bu arada Excelde olduğunun aksine VBA'de <span class=" keywordler">Blank</span> diye 
bir kavram dolayıysla <span class=" keywordler">IsBlank </span>diye bir sorgulama şekli de yoktur.</p>
	<p>Son olarak, bir şekilde başlangıç değeri olmayan Variant tipli bir değişkene
	<span class="keywordler">IsNumeric</span> sorgulaması yapıldığında True 
	değeri döndürür.</p>

<h2>Null</h2>
<p>Bir değişken veri içermiyorsa bu değişken <span class=" keywordler">Null</span> değere sahiptir diyebiliriz. Bir değişkenin değerinin Null olabilmesi için ya bilinçli bir şekilde Null ataması ya da Null içeren başka birşeyle etkileşime girmesi gerekir. Empty gibi Null da sadece Variant tipinin bir özelliğidir, yani saece Variant tipteki bir değişlen Null değer alabilir. Başka tipteki bir değişkene Null atanmak istendiğinde hata alınır. 
Ancak Stringlere <span class=" keywordler">vbNullString</span> ile null atama 
yapılabilir. Bir değişkenin içeriğinin Null olma olasılığı varsa ve bunu başka bir değişeknele işleme sokacaksanız, işleme sokmadan önce o anda Null içerip içermediğini <span class=" keywordler">IsNull()</span> ile kontrol etmeniz gerekir, aksi halde yine hatayla karşılaşırsınız.</p>

<h2>Nothing</h2>
<p>Ve son olarak bir de <span class=" keywordler">Nothing</span> var. Tanımlanmış ancak henüz yaratılmamış 
<strong>objelerin</strong> değeri <span class=" keywordler">Nothing</span>dir. Bir objeye bu değer atandığında ise, objenin kendisiyle obje değişkeni arasındaki bağı koparmış oluruz.</p> 


<p><span class=" keywordler">Nothing</span> sadece object tipindeki değişkenlere atanan bir özelliktir ve Set ifadesi ile kullanılır. Bir nesnenin <span class=" keywordler">Nothing</span> olup olmadığını anlamak için "<span class="keywordler">Is Nothing</span>"(iki 
kelme ayrı) sorgulaması yapılır. Burada "is" kullanımı önemlidir, eşitlik ("=") yerine "olmak" ile sorguluyoruz.</p>
	
<pre class="brush:vb">If Degisken Is Nothing Then</pre>

<p><a name="DebugOrnekliKod"></a>Şimdi de VBA'de tüm bunlar nasıl kullanılıyor, sonuçları ne oluyor, ona bi bakalım.</p>

<pre class="brush:vb">
Sub null_empty_nothing_zls()
'Vartype, TypeName
'0:empty, 1:null, 2:int, 3:long,....7:Date, 8:string, 9:object, 11:boolena, 12:Variant (sadece variant arraylerde), 8192:Array(normal değer + 8192)
'Vartype'ı sorgularken rakamları bilmen gerekmez, vbConst ifadelerini de kullanabilrsin, aşağıdaki örnekte ='den sonra intellisense ile seçebilirsinizDebug.Print "henüz tanımlanmamış bir p değişkeni için VarType(p) = vbNull mı sonucu:" &amp; IIf(VarType(p) = vbNull, True, False) 'intellisense de çalışır ve sana seçtirir
Debug.Print vbNewLine

'Default değerler
'bi makro çalışıtğında tüm değişkenler ilk olarak default değerlere atanır
'numerikse:0 stringse:""(ZLS) objectse:Nothing variantsa:Empty

Dim v 'variant olduğu için empty başlar
Dim v2 'az sonra değer atıycam
Dim i As Integer 'int olduğu için 0 başlar
Dim s As String 'string olduğu için zls bşlar, empty değil
Dim r As Range 'object olduğu için başlangıcı nothing
Dim z As String 'zls olacak, s ile aynı sonuçları verir
Dim n 'variant olduğu için empty başlar
Dim n2 'variant olduğu için empty başlar
Dim u 'null string
Dim o As Object 'obje olduğu için nothing başlar
Dim d(10) As String ' dizi olduğu için defaultdeğerlerle başlar
Dim vd() As Variant

'atamalar
v2 = Array("armut", "elma", "kiraz") 'dizi
z = ""
n = Null 'artık empty değil null
u = vbNullString
n2 = n * 10 'null ile etkileşime girdiği için null


'TM=Type Mismatch hatası almasınlar diye bu ifadenin kısaltaması olan TM yazdım
Debug.Print "Variable açıklaması" &amp; vbTab &amp; "Değer" &amp; vbTab &amp; "VarType" &amp; vbTab &amp; "TypeName" &amp; vbTab &amp; "IsEmpty" &amp; vbTab &amp; vbTab &amp; "IsNull" &amp; vbTab &amp; "ZLS mi" &amp; vbTab &amp; "0 mı" &amp; vbTab &amp; "vbNullStrmi" &amp; vbTab &amp; "LenB" &amp; vbTab &amp; "Nothing mi"
Debug.Print "---------------------------------------------------------------------------------------------------------------------------"
Debug.Print "v:Variant data yok" &amp; vbTab &amp; "| " &amp; v &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(v) &amp; vbTab &amp; vbTab &amp; TypeName(v) &amp; vbTab &amp; vbTab &amp; IsEmpty(v) &amp; vbTab &amp; vbTab &amp; IsNull(v) &amp; vbTab &amp; IIf(v = "", True, False) &amp; vbTab &amp; IIf(v = 0, True, False) &amp; vbTab &amp; IIf(v = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(v) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "v2:Variant dizi " &amp; vbTab &amp; "| " &amp; "N/A" &amp; vbTab &amp; "| " &amp; VarType(v2) &amp; vbTab &amp; TypeName(v2) &amp; vbTab &amp; IsEmpty(v2) &amp; vbTab &amp; vbTab &amp; IsNull(v2) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; "N/A"
Debug.Print "v2(0):V Dizi Eleman" &amp; vbTab &amp; "| " &amp; v2(0) &amp; vbTab &amp; "| " &amp; VarType(v2(0)) &amp; vbTab &amp; vbTab &amp; TypeName(v2(0)) &amp; vbTab &amp; vbTab &amp; IsEmpty(v2(0)) &amp; vbTab &amp; vbTab &amp; IsNull(v2(0)) &amp; vbTab &amp; IIf(v2(0) = "", True, False) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(v2(0) = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(v2(0)) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "i:Integer data yok" &amp; vbTab &amp; "| " &amp; i &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(i) &amp; vbTab &amp; vbTab &amp; TypeName(i) &amp; vbTab &amp; vbTab &amp; IsEmpty(i) &amp; vbTab &amp; vbTab &amp; IsNull(i) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(i = 0, True, False) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; vbTab &amp; LenB(i) &amp; vbTab &amp; vbTab &amp; "*TM*"
Debug.Print "s:String data yok" &amp; vbTab &amp; "| " &amp; s &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(s) &amp; vbTab &amp; vbTab &amp; TypeName(s) &amp; vbTab &amp; vbTab &amp; IsEmpty(s) &amp; vbTab &amp; vbTab &amp; IsNull(s) &amp; vbTab &amp; IIf(s = "", True, False) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(s = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(s) &amp; vbTab &amp; vbTab &amp; "*TM*"
Debug.Print "z:ZLS çift tırnak" &amp; vbTab &amp; "| " &amp; z &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(z) &amp; vbTab &amp; vbTab &amp; TypeName(z) &amp; vbTab &amp; vbTab &amp; IsEmpty(z) &amp; vbTab &amp; vbTab &amp; IsNull(z) &amp; vbTab &amp; IIf(z = "", True, False) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(z = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(z) &amp; vbTab &amp; vbTab &amp; "*TM*"
Debug.Print "n:Null atanmış " &amp; vbTab &amp; vbTab &amp; "| " &amp; n &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(n) &amp; vbTab &amp; vbTab &amp; TypeName(n) &amp; vbTab &amp; vbTab &amp; IsEmpty(n) &amp; vbTab &amp; vbTab &amp; IsNull(n) &amp; vbTab &amp; IIf(n = "", True, False) &amp; vbTab &amp; IIf(n = 0, True, False) &amp; vbTab &amp; IIf(n = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(n) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "n2:Null'la işlem " &amp; vbTab &amp; "| " &amp; n &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(n2) &amp; vbTab &amp; vbTab &amp; TypeName(n2) &amp; vbTab &amp; vbTab &amp; IsEmpty(n2) &amp; vbTab &amp; vbTab &amp; IsNull(n2) &amp; vbTab &amp; IIf(n2 = "", True, False) &amp; vbTab &amp; IIf(n2 = 0, True, False) &amp; vbTab &amp; IIf(n2 = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(n2) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "u:vbNullString " &amp; vbTab &amp; vbTab &amp; "| " &amp; u &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(u) &amp; vbTab &amp; vbTab &amp; TypeName(u) &amp; vbTab &amp; vbTab &amp; IsEmpty(u) &amp; vbTab &amp; vbTab &amp; IsNull(u) &amp; vbTab &amp; IIf(u = "", True, False) &amp; vbTab &amp; IIf(u = 0, True, False) &amp; vbTab &amp; IIf(u = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(u) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "r:Range set'siz " &amp; vbTab &amp; "| " &amp; "N/A" &amp; vbTab &amp; "| " &amp; VarType(r) &amp; vbTab &amp; vbTab &amp; TypeName(r) &amp; vbTab &amp; vbTab &amp; "N/A " &amp; vbTab &amp; vbTab &amp; IsNull(r) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(r Is Nothing, True, False)
Debug.Print "o:Obje nesne yok " &amp; vbTab &amp; "| " &amp; "N/A" &amp; vbTab &amp; "| " &amp; VarType(o) &amp; vbTab &amp; vbTab &amp; TypeName(o) &amp; vbTab &amp; vbTab &amp; "N/A " &amp; vbTab &amp; vbTab &amp; IsNull(o) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(o Is Nothing, True, False)
Debug.Print "d:Dizi elemansız " &amp; vbTab &amp; "| " &amp; "N/A" &amp; vbTab &amp; "| " &amp; VarType(d) &amp; vbTab &amp; TypeName(d) &amp; vbTab &amp; IsEmpty(d) &amp; vbTab &amp; vbTab &amp; IsNull(d) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; "*TM*"
Debug.Print "d(0):Dizi Eleman" &amp; vbTab &amp; "| " &amp; d(0) &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(d(0)) &amp; vbTab &amp; vbTab &amp; TypeName(d(0)) &amp; vbTab &amp; vbTab &amp; IsEmpty(d(0)) &amp; vbTab &amp; vbTab &amp; IsNull(d(0)) &amp; vbTab &amp; IIf(d(0) = "", True, False) &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; IIf(d(0) = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(d(0)) &amp; vbTab &amp; vbTab &amp; "N/A"
Debug.Print "vd:Variant dizi " &amp; vbTab &amp; "| " &amp; "N/A" &amp; vbTab &amp; "| " &amp; VarType(vd) &amp; vbTab &amp; TypeName(vd) &amp; vbTab &amp; IsEmpty(vd) &amp; vbTab &amp; vbTab &amp; IsNull(vd) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "*TM*" &amp; vbTab &amp; "N/A"


Debug.Print vbNewLine
Debug.Print "Yeniden atamalar yapıyoruz"

'Empty değer ataması
v = Empty 'hala emtpydir
Debug.Print "v:Empty Atanıyor " &amp; vbTab &amp; "| " &amp; v &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(v) &amp; vbTab &amp; vbTab &amp; TypeName(v) &amp; vbTab &amp; vbTab &amp; IsEmpty(v) &amp; vbTab &amp; vbTab &amp; IsNull(v) &amp; vbTab &amp; IIf(v = "", True, False) &amp; vbTab &amp; IIf(v = 0, True, False) &amp; vbTab &amp; IIf(v = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(v) &amp; vbTab &amp; vbTab &amp; "N/A "

i = Empty 'yine de 0'dır, çünkü empty sadece variantın bir tipi
i = 456
Debug.Print "i:Değerli Integer" &amp; vbTab &amp; "| " &amp; i &amp; vbTab &amp; "| " &amp; VarType(i) &amp; vbTab &amp; vbTab &amp; TypeName(i) &amp; vbTab &amp; vbTab &amp; IsEmpty(i) &amp; vbTab &amp; vbTab &amp; IsNull(i) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; IIf(i = 0, True, False) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; LenB(i) &amp; vbTab &amp; vbTab &amp; "*TM*"

v = 23
Debug.Print "v:Integer variant" &amp; vbTab &amp; "| " &amp; v &amp; vbTab &amp; "| " &amp; VarType(v) &amp; vbTab &amp; vbTab &amp; TypeName(v) &amp; vbTab &amp; vbTab &amp; IsEmpty(v) &amp; vbTab &amp; vbTab &amp; IsNull(v) &amp; vbTab &amp; IIf(v = "", True, False) &amp; vbTab &amp; IIf(v = 0, True, False) &amp; vbTab &amp; IIf(v = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(v) &amp; vbTab &amp; vbTab &amp; "N/A "
v = "acb"
Debug.Print "v:String variant" &amp; vbTab &amp; "| " &amp; v &amp; vbTab &amp; "| " &amp; VarType(v) &amp; vbTab &amp; vbTab &amp; TypeName(v) &amp; vbTab &amp; vbTab &amp; IsEmpty(v) &amp; vbTab &amp; vbTab &amp; IsNull(v) &amp; vbTab &amp; IIf(v = "", True, False) &amp; vbTab &amp; IIf(v = 0, True, False) &amp; vbTab &amp; IIf(v = vbNullString, True, False) &amp; vbTab &amp; vbTab &amp; LenB(v) &amp; vbTab &amp; vbTab &amp; "N/A "

Set r = Range("a1")
Debug.Print "r:Range set'li " &amp; vbTab &amp; vbTab &amp; "| " &amp; "" &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(r) &amp; vbTab &amp; vbTab &amp; TypeName(r) &amp; vbTab &amp; vbTab &amp; "N/A " &amp; vbTab &amp; vbTab &amp; IsNull(r) &amp; vbTab &amp; IIf(r = "", True, False) &amp; vbTab &amp; IIf(r = 0, True, False) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; IIf(r Is Nothing, True, False) 'obj olduğu için değer yazdırılamaz

Set o = CreateObject("Outlook.Application")
Debug.Print "o:Obje(outlook) " &amp; vbTab &amp; "| " &amp; "" &amp; vbTab &amp; vbTab &amp; "| " &amp; VarType(o) &amp; vbTab &amp; vbTab &amp; TypeName(o) &amp; " N/A " &amp; vbTab &amp; vbTab &amp; IsNull(o) &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; vbTab &amp; "N/A" &amp; vbTab &amp; vbTab &amp; IIf(o Is Nothing, True, False) 'obj olduğu için değer yazdırılamaz


'Çıktısı aşağıdaki gibidir
'henüz tanımlanmamış bir p değişkeni için VarType(p) = vbNull mı sonucu:False


'Variable açıklaması Değer VarType TypeName IsEmpty IsNull ZLS mi 0 mı vbNullStrmi LenB Nothing mi
'---------------------------------------------------------------------------------------------------------------------------
'v:Variant data yok | | 0 Empty True False True True True 0 N/A
'v2:Variant dizi | N/A | 8204 Variant() False False N/A N/A N/A *TM* N/A
'v2(0):V Dizi Eleman | armut | 8 String False False False *TM* False 10 N/A
'i:Integer data yok | 0 | 2 Integer False False *TM* True *TM* 2 *TM*
's:String data yok | | 8 String False False True *TM* True 0 *TM*
'z:ZLS çift tırnak | | 8 String False False True *TM* True 0 *TM*
'n:Null atanmış | | 1 Null False True False False False N/A
'n2:Null'la işlem | | 1 Null False True False False False N/A
'u:vbNullString | | 8 String False False True False True 0 N/A
'r:Range set'siz | N/A | 9 Nothing N/A False N/A N/A N/A *TM* True
'o:Obje nesne yok | N/A | 9 Nothing N/A False N/A N/A N/A *TM* True
'd:Dizi elemansız | N/A | 8200 String() False False N/A N/A N/A *TM* *TM*
'd(0):Dizi Eleman | | 8 String False False True *TM* True 0 N/A
'vd:Variant dizi | N/A | 8204 Variant() False False N/A N/A N/A *TM* N/A


'Yeniden atamalar yapıyoruz
'v:Empty Atanıyor | | 0 Empty True False True True True 0 N/A
'i:Değerli Integer | 456 | 2 Integer False False N/A False N/A 2 *TM*
'v:Integer variant | 23 | 2 Integer False False False False False 4 N/A
'v:String variant | acb | 8 String False False False False False 6 N/A
'r:Range set'li | | 0 Range N/A False True True N/A N/A False
'o:Obje(outlook) | | 8 Application N/A False N/A N/A N/A N/A False
End Sub
</pre>

</div>

<h2 class="baslik">VB.Net</h2>
<div class="konu">
<p>VB.Net, .Net çatısı altındaki dillerden biri olup, sadece VBA'den değil onun 
atası olan VB6'dan da bazı noktalarda ayrılmaktadır. Mesela VBA'de tüm değişkenler local değişken iken VB.Net'de 
local değişken ve global değişken(field) ayrımı vardır. VBA'de tüm değişkenler makro çalışır çalışmaz default 
değerlerine atanırken, VB.Net'te sadece global değişkenler default değerlerine atanır, local değişkenlerin ise mutlaka bir ilk değerinin olması beklenir, aksi halde hata alınır.</p>

<h2>Null/Nothing</h2>
<p>Bir diğer fark da şudur: Vb.Net'te değişkenler değer tipli ve referans tipli olmak üzere ikiye ayrılır. Mesela stringler referans tiplidir ve tüm diğer referans tipler gibi null değer alabilir, (VBA'da sadece variant tipi değişkenler null değer alabiliyordu). Bu arada konuşurken null değerden bahsediyoruz ancak Vb.Net'te bir değişkenin null değer alması <span class=" keywordler">Nothing</span> keywordu ile sağlanır, <span class=" keywordler">null</span> keywordu ile değil.</p>

<pre class="brush:vb">Dim str As String = Nothing</pre>

<p>Değişkenin null mı olduğu <span class="keywordler">IsNothing(degisken)</span> 
metodu ile veya referans tiplerde <span class="keywordler">if degisken 
<span style="text-decoration: underline"><strong>is</strong></span> Nothing</span>, 
nullable olmayan değer tipli değişkenler <span class="keywordler">if degisken <span style="text-decoration: underline">
<strong>=</strong></span> Nothing</span> şeklinde gibi sorgulanabilir. Burada detayına girmeyeceğim ancak 
<span class="keywordler">IsNothing</span> metodunun birçok yerde kullanılmaması öğütleniyor, ben de bu öğüte uyuyorum, aşağıdaki karşılaştırma tablosuna dahi almıyorum. 
Bunun yerine ikinci yöntemi kullanacağız. </p>

<pre class="brush:vb">
'değer tipli sorgulama şekli. Her ne kadar stringler referans tipli de olsa daha önceden belirtitğim gibi değer tipli sorgulamsı yapılır.
Dim str As String = Nothing
If str = Nothing Then 
   MsgBox("String Null'dır")
End If

'referans tipli sorgulama
Dim obj As Object = Nothing
If obj is Nothing Then 
   MsgBox("Nesnemiz Null'dır")
End If

'IsNothing ile her iki tür (değer veya referans)de sorgulanabilr
If IsNothing(s) Then
    Console.WriteLine(True)
End If
</pre>


<p>Nullable olmayan değer tipli değişkenler için <span class="keywordler">Nothing'</span>in özel bir anlamı vardır; böyle bir değişkene 
<span class="keywordler">Nothing</span> değeri atadığımızda ona default değerini atamış oluruz. 
<em>(Bu, obje elemanlardan oluşan bir dizi veya dizimsi yapılardaki tüm elemanları bir döngü içinde default değere döndürmek için kullanışlı bir yöntemdir, onun dışında bir değişkene default değeri atamak için
</em><span class="keywordler"><em>Nothing</em></span><em> yerine kendi default değeri neyse onu atamak daha mantıklı bir yol olacaktır.)</em> C#'ta ise nullable olmayan bir değişkene null değer atamak hata verecektir.</p>

<p>Aşağıda MSDN sitesinden alınan ve bu durumu açıklayan bir kod bloğu bulunmaktadır.</p>

<pre class="brush:vb">
Module Module1

    Sub Main()
        Dim ts As TestStruct 'değer tipli bir değişken, çünkü bir structure
        Dim i As Integer 'nullable olmayan değer tipli bir integer
        Dim b As Boolean 'nullable olmayan değer tipli bir boolean

        ' Aşağıdaki ifade ts.Name'in değerini Nothing'e ve ts.Number'ı da 0'a çevirir.
        ts = Nothing

        ' Aşağıdaki ifade i'yi 0 ve b'yi False yapar.
        i = Nothing
        b = Nothing

        Console.WriteLine("ts.Name: " & ts.Name)
        Console.WriteLine("ts.Number: " & ts.Number)
        Console.WriteLine("i: " & i)
        Console.WriteLine("b: " & b)

        Console.ReadKey()
    End Sub

    Public Structure TestStruct
        Public Name As String
        Public Number As Integer
    End Structure
End Module
</pre>

<p>Bir değişken referans tipliyse ona <span class="keywordler">Nothing</span> değerini atamak onun hiçbir nesneyle bağının kalmaması demektir. Aşağıdaki örnek bunu gösterir.</p>


<pre class="brush:vb">
Module Module1

    Sub Main()

        Dim testObject As Object
        ' Aşağıdaki ifade testObject'i null yapar
        testObject = Nothing

        Dim tc As New TestClass
        tc = Nothing
        ' yukardaki ifade tc'yi null yapar, böylece tc'nn hiçbr elamnına erişlemez, çünkü bu bir referans tipli classtır. bir önceki örnekte ise bir strcuturın elenalrıan erişeibliyorduk. Dolayıyısla bir alt satırda hata alınır.
        Console.WriteLine(tc.Field1)

    End Sub

    Class TestClass
        Public Field1 As Integer
        ' . . .
    End Class
End Module</pre>


<h2>Empty</h2>
<p>Bir değişkene ilk değer olarak boş değer atanmak isteniyorsa <span class=" keywordler">String.Empty</span> ifadesi kullanılmalıdır. 
Bu, VBA'deki vbNullstringe denk düşmektedir. ZLS burda da geçerlidir ancak VBA'daki aynı performans sebeplerinden dolayı kullanılmaması daha iyi olacaktır.</p>

<pre class="brush:vb">Dim str As String = String.Empty</pre>


<p>Değişkenin empty mi olduğu aşağıdaki gibi sorgulanabilir:</p>

<pre class="brush:vb">
Dim str As String = String.Empty
If str = String.Empty Then
   MsgBox("String Empty'dir")
End If</pre>

<p>veya</p>
<pre class="brush:vb">If degisken.Length=0</pre>

<p>String sınıfının <span class=" keywordler">IsNullOrEmpty </span>diye güzel bir metodu var, bununla bir değişkenin aynı anda null veya empty mi olduğu tek seferde sorgulanabilir. Ancak sadece Empty mi veya sadece Null mı sorgulaması yapmak için bunu değil yukardaki fonskiyonları kullanın.</p>

<pre class="brush:vb">
  Dim str As String = Nothing
  If String.IsNullOrEmpty(str) Then
	  MsgBox("String null veya empty'dir")
  End If
</pre>


</div>

<h2 class="baslik">C#</h2>
<div class="konu">
<p>C# da Vb.Net gibi .Net çatısının altındaki dillerden biri olup, VB.Net konusu altında bahsettiğim birçok şey 
C# için de geçerlidir. Farklılıklar ise şöyledir.</p>

<p>Vb.Net null değerler için <span class=" keywordler">Nothing </span>ifadesini kullanırken c# <span class=" keywordler">null</span> 
ifadesini kullanır ve sorgulaması <span class="keywordler">if (degiksen==null)</span> 
ifadesi ile yapılır.</p>

<pre class="brush:csharp">class Program
{
    static string globaldegisken; //class seviyesindeki global değişken, diğer adı field

    static void Main()
    {
	// global string değişkenimize hemen defualt değeri olan null atanır
	if (globaldegisken == null) //string her ne kadar referans tipli olsa da sorgulama yapılırken değer tipli değişkenler gibi == ile sorgulanır. eğer bu başka bir obje olsaydı "Is" ile sorgulardık
	{
	    Console.WriteLine("Global Değişkenimiz null'dır");
	}

	string localdegisken1; //local değişken

	// // burayı çalışıtırısak hata laırız, çünkü lokcal değişkene henüz bir değer atanmadı
	// if (localdegisken1== null)
	// {
        // ......
	// }

	string localdegisken2=null; //local değişken

	 // burayı çalışıtırısak hata laırız, çünkü lokcal değişkene henüz bir değer atanmadı
	 if (localdegisken2== null)
	 {
	    Console.WriteLine("Local Değişkenimiz null'dır");
	 }
    }
}</pre>


<p>Ayrıca yine Vb.Net bölümünde gördüğümüz üzere, C#'ta referans tipli olmayan hiçbir değişkene null değer atanamaz(nullable değişkenler hariç)</p>
</div>

<h2 class="baslik">SQL</h2>
<div class="konu">
<p><span class=" keywordler">Coalesce:</span>İlk null olmayan değeri elde edersiniz. Coalesce(x,y,z). 
Ansi standartı olup tüm SQL platformlarında geçerlidir. Alternatifi
<span class="keywordler">Case When</span></p>
<p><span class=" keywordler">Nvl(Oracle)</span>:ifade sonucu nullsa ikinci parametre 
döner. Nvl(ifade,x)</p>
	<p><span class="keywordler">IsNull(SQL Server)</span>:İfade sonucu nullsa ikinci parametre. 
IsNull(ifade,x)&nbsp;&nbsp;&nbsp; </p>
<p><span class=" keywordler">Nvl2(Oracle):</span>ifade sonucu <strong>null 
değilse</strong> ikinci parametre yoksa üçüncü parametre döner. Nvl2(ifade,x,y)<br>
<span class="keywordler">IIF(SQL Server)</span>:ifade sonucu <strong>nulsa
</strong>ikinci parametre, yoksa üçüncü parametre döner. IIF(ifade is null,x,y)</p>
<p><span class=" keywordler">Nullif</span>:iki değer de eşitse null, eşit değilse ilk değer. Ansi 
standartı olup tüm SQL platformlarında geçerlidir.</p>
	<p>Where kısmında null kayıtları sorgulamak ise çok daha basit.</p>
<pre class="brush: vb">
Select * 
from tabloadı
where
alan is null --null olmayanlar için "alan is not null"
</pre>

</div>

<h2 class="baslik">Nihai Karşılaştırma tablosu</h2>
<div class="konu">


<p>İlk değer ataması, boş değer atama ve boş değer sorgulaması aşağıdaki gibi özetlenebilir</p>

<table class="alterantelitable">
<th>Karşılaştırma Konusu</th>
<th>Değer Atanmamışsa</th>
<th>Boş Atama</th>
<th>Boşmu</th>



<tr>
<td>Excel</td>
<td>N/A</td>
<td>Boş bırakılır veya ="" yazılabilir, özellikle bir formül içindeyse ="" tek yoldur. Ör: IF(A1=0;"";A1)</td>
<td>ISBLANK():Boşsa true <br>
If(A1="";True;False):Hem boşsa hem "" ise true</td>
</tr>

<tr>
<td>VBA</td>
<td>Tüm değişkenlere default değerler atanır.<br>Stringse&gt;ZLS,<br>Numerikse&gt;0,<br>Variantsa&gt;Empty,<br>
Arrayse&gt;Emtpy,<br>Objcetse&gt;Nothing</td>
<td>Stringse&gt;="" veya =vbNullString><br>
Variantsa&gt;=Empty</td>
<td>Stringse&gt;if len(degisken)=0 veya<br>if degisken=""<br>
Variantsa&gt;IsEmpty()<br></td>
</tr>


<tr>
<td>Vb.Net</td>
<td>Sadece gloabal değişkenlere default değerler atanır.<br>Stringse&gt;Nothing<br>
Numerikse&gt;0<br>Objectse&gt;Nothing<br>Arrayse&gt;Nothing</td>
<td>String dahil tüm referans tiplerde &gt;=String.Empty<br>
<br>Stringte ayrıca =""</td>
<td>If degisken= String.Empty veya if degisken.Length=0<br>
veya<br>String.IsNullOrEmpty(degisken)</td>
</tr>

<tr>
<td>C#</td>
<td>Sadece gloabal değişkenlere default değerler atanır.<br>Stringse&gt;null<br>
Numerikse&gt;0<br>Objectse&gt;null<br>Arrayse&gt;null</td>
<td>String dahil tüm referans tiplerde &gt;=String.Empty<br>
Stringte ayrıca =""</td>
<td>If (degisken== String.Empty) veya if (degisken.Length==0)<br>
veya if (String.IsNullOrEmpty(degisken))</td>
</tr>

</table>


<p>Null/Nothing ataması ve sorgulaması ise aşağıdaki gibi özetlenebilir</p>

<table class="alterantelitable">
<th>Karşılaştırma Konusu</th>
<th>Null Atama</th>
<th>Null mı</th>


<tr>
<td>VBA</td>
<td>Stringse&gt; =vbNullString<br>
Variantsa&gt;=Null<br>
Objectse&gt;Set obj=Nothing</td>
<td>String:if degisken=vbNullString<br>
Variant:if IsNull(degisken)<br>
Object:if obj is Nothing<br>
<strong>Olumsuzu</strong>:if not obj is Nothing
</td>
</tr>


<tr>
<td>Vb.Net</td>
<td>String dahil tüm referans tiplerde&gt;=Nothing.<br>
Değer tiplerde kullanıldığında onların default değerlerine döndürür.</td>
<td>Referans tiplerde If degisken is Nothing<br>
Nullable olmayan değer tiplerde If degisken = Nothing<br>
veya if String.IsNullOrEmpty(degisken)
<strong>Olumsuzu:</strong><br>
Referans tipler için> if degisken IsNot Nothing(veya eski VB'den kalma yöntem: If not degisken is Nothing)<br>
Nullable olmayan değer tipler için>if degisken <> Nothing(veya eski VB'den kalma yöntem:If not degisken = Nothing)<br>

</td>
</tr>

<tr>
<td>C#</td>
<td>String dahil tüm referans tiplerde&gt;=null</td>
<td>If (degisken==null)<br>veya if String.IsNullOrEmpty(degisken)<br>
<strong>Olumsuzu:</strong>if (degisken !=null)</td>
</tr>

</table>
</p>
</div>


</asp:Content>
