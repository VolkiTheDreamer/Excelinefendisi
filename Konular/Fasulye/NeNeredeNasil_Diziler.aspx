<%@ Page Title='NeNeredeNasil Diziler' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Fasulye'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Ne Nerede Nasıl'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushCSharp.js" ) %>"> </script>
<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushSql.js" ) %>"> </script>
<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushJScript.js" ) %>"> </script>

<h1>Diziler</h1>
<p>Heralde kodlama yaptığım sürede u diziler(ve dizimsiler) kadar kafamı karıştıran birşey olmamıştır. Şimdi bunlara bi bakalım,. Örneklerde basitlik olması adına tek boyutlu dizilerden bahsedilecektir.</p>
<p><strong>Ön bilgi:</strong>VBA/VB.Net küçük/büyük harf ayrımı gözetmez, o yüzden property ve metod isimlerinde harfin büyüklüğü önemli değildir. c# ve javascript ise case-sensitive'dir. C#'ta property ve metodlar büyük harfle başlarken, javascriptte küçük harfle başlar. Aşağıda da bu gösterime sadık kalınmıştır. Ör:C#'ta boyut özellilği <span class="keywordler">Length</span> ile elde edilirken javascriptte <span class=" keywordler">length</span> ile elde edilir.</p>

<h2 class='baslik'>Genel Bilgiler</h2>
<div class='konu'>
<p>Syntax'a baktıgımızda aşağıdaki dillerden VB kökenli dillerde dizi parantezi normal parantez yani () iken, diğer dillerde köşeli parantez yani [] şeklinde ifade edilir. Ve yine VB kökenli dillerde bu parentez dizi değişkeninin yanına yazılırken, diğerlerinde veri tipinin yanına yazılır.</p>

<h2>Base(taban)</h2>
<p><strong>VBA:</strong>Varsayılan değer 0'dır, ancak <span class=" keywordler">Option Base 1</span> yapılarak baz 1'den başlatılabilir. (<span class=" keywordler">ParamArray</span>'ler hariç). Benim tavsiyem, varsayılan değeri bozmayın ve 0'dan başlatın ve bunu alışkanlık haline getirin, yoksa diğer dillerle çalışmaya başladığınızda karmaşa yaşayabilirsiniz. Baz index değeri VBA'de <span class=" keywordler">Lbound</span> 
fonksiyonu ile bulunabilir ve bu metod çoğu durumda 0 döndürecektir.</p>

<p>VBA'de cells, worksheets gibi bazı collection üyelerinin indexi 1'den başlar 
ve bu değiştirilemez.</p>




	<p><strong>VB.Net:</strong>VBA 'de yazılı olanlar burda da geçerlidir. Ancak 
	baz index değeri VBA'de <span class=" keywordler">Lbound</span> fonksiyonu ile 
	bulunurken VB:Net'te <span class=" keywordler">GetLowerBound</span> bulunabilir.</p>

<p><strong>C#:</strong>Varsayılan değer 0'dır. Optiona Base gibi bir seçenek yoktur, ancak wrapper class yazarak istenilen indexten başlatılabilir(bence gereksizdir)</p>

<p><strong>Javascript:</strong>Varsayılan değer 0'dır.</p>

<p>Bu arada, tüm dillerde, baz değer istenirse 0'dan küçüğe de ayarlanabilir, bunun çok kullanım yeri olduğunu düşünmüyorum ancak teorik olarak imkan bulunmaktadır.</p>




<h2>Boyut/Uzunluk/Eleman Sayısı ve Üst Değer(Son eleman indexi)</h2>
<p><strong>VBA</strong>:Dizi boyutu, dizi tanımlanırken parantez içinde verilen değer+1'dir, 
y<strong>ani parantez içine yazılan değer boyut değil son indextir </strong>ve 
son index <span class=" keywordler">Ubound(dizi)</span> ile boyut ise
<span class=" keywordler">Ubound(dizi)+1</span> fonksiyonu ile elde edilir.&nbsp; 
Dolu eleman sayısı ise loop'a sokularak bulunur.</p>
	<p><strong>VB.Net:</strong>Dizi boyutu, dizi tanımlanırken parantez içinde verilen değer+1'dir,
	<strong>yani parantez içine yazılan değer boyut değil son indextir</strong> 
	ve son index <span class=" keywordler">GetUpperBound</span> 
	ile boyut ise <span class="keywordler">Length</span> özelliği veya
	<span class="keywordler">GetLength(d)</span> metodu ile elde edilir. 
	GetLength ile belirli bir boyuttaki elemans sayısı alınırken Length ile tüm 
	boyutlardaki eleman toplamı gelir. Dolu eleman sayısı ise loop'a sokularak 
	bulunur. </p>

<p>Dizilerin bir boyutu olması gerektiğini biliyoruz, ancak VBA ve VB.Net'te ilk tanımlarken boyut vermek zorunda değiliz, boyutsuz tanımlayıp, daha sonra <span class=" keywordler">ReDim</span> ifadesi ile de boyut verebiliriz. 
Aşağıdaki konu başlığında bunun detaylarını görebilirsiniz.</p>

<p><strong>VBA/VB.Net(List(sadece VB.Net)/Collection/Dictionary):</strong>Dizimsi yapıların baştan belirlenmiş bir boyutu yoktur, belli bir andaki boyutu <span class=" keywordler">Count</span> property'si ile bulunur. Boyutsuz tanımlandıkları için dolu eleman sayısı diye bir kavram yoktur. Son eleman indexi, <span class=" keywordler">Count-1</span> şeklinde bulunur. </p>

<p><strong>C#:</strong>C#'ta dizi boyutu parantez içindeki rakamdır, dolayısıyla 
son index numarası da bundan 1 eksik olan rakamdır. Kullanılan özellik ve 
metodlar VB.Net ile aynıdır.</p>

<p><strong>C#(List/Dictionary/Collection):</strong>VB.Net'teki gibi C#'ta da dizimsi elemanların eleman sayısı <span class=" keywordler">Count</span> propertysi ile tespit edilir.</p>

<p><strong>Javascript(Array):</strong>Boyut <span class=" keywordler">length</span> ile alınır. </p>




<h2>Initialization(Diziye ilk değer atama)</h2>
<p>The default values of numeric array elements are set to zero, and reference elements are set to null.</p>

<p><strong>VBA:</strong>VB'de dizi tanımlaması aşağıdaki şekillerde yapılabilir</p>

<pre class="brush: vb">
Dim dizi1() as String	'eleman sayısı sonraya bırakılabilir. VBA'da veritipi belirtilmezse Variant olur, VB.Nette veritipi belirtilmelidir
Dim dizi2(5) as String 'son indexi 5 olan 6 elemanlı dizi
Dim dizi3(1 to 5) as String 'son indexi 5 olan 5 elemanlı dizi
Dim dizi4 = Array("Elma","Armut","Karpuz") 'Elemanlar baştan belirtilerek. Veri tipi belritilemdiği için Varianttır.
</pre>

<p><strong>VB.Net:</strong>VB.Net'te de VBA ile aynı ifadeler geçerli iken ayrıca aşağıdaki gibi de initialization yapılabilir.</p>
<pre class="brush: vb">Dim nums() As Integer = {1, 2, 3}</pre>

<p><strong>C#:</strong>Dizi üyeleri, eğer deklarasyon anında belirtilmedilerse default değerleri atanır. Ayrıca boyut veya elemanlardan biri mutlaka belirtilmelidir </p>
<p>Aşağıda, MSDN sitesinden alınan örnek başlatım şekilleri bulunuyor.</p>
<pre class="brush: csharp">
// Tek boyutlu nümerik diziler
int[] n0 = new int[4];//sadece boyut belirtilip, değer olarak default değerler atanır, bu örnekte 0 atanır. gerektiğinde bunların dğeerleri program akışı içinde atanır
int[] n1 = new int[4] {2, 4, 6, 8};//boyut ve elemanlar birlikte belirtilebilir
int[] n2 = new int[] {2, 4, 6, 8};//sadece elemanlar new ifadesi ile belirtilebilir
int[] n3 = new[] {2,4,6,8}//Sistem kendisi anlar ne tipte olduğunu
int[] n4 = {2, 4, 6, 8};//sadece elemanlar new ifadesi olmadan belirtilebilir

// Tek boyultu string diziler
string[] s0 = new string[3];//defualt değer olarak null atanır, çünkü string bir referans tiptir
string[] s1 = new string[3] {"John", "Paul", "Mary"};
string[] s2 = new string[] {"John", "Paul", "Mary"};
string[] s3 = new[] {"John", "Paul", "Mary"};
string[] s4 = {"John", "Paul", "Mary"};

//boş ve null diziler
int[] bosArray = new int[0]; // Boş dizi
int[] nullArray; // Tüm dizilerin default değeri Null'dır, çünkü diziler referans tipli nesnelerdir. İçeriği belirtilmediği için bu dizinin de şuanki değeri null'dır.
//int[] nullArray; ifadesinin default değerinin null olması ile, int[] dizi=new int[4]; ifadesinin elemanları int olup elemanlar belirtilemdiği için bunları default değeri 0dır. yani ilkinde dizniin kendisi default değernine, ikncisinde ise elemanlar default değerlerine atanır. 
</pre>

<p>Bunlardan ilk 4ü aynı zamanda <span class=" keywordler">var</span> ifadesi ile de yapılabilrdi, çünkü eşitliğin sağındaki bilgi diziyi başlatmak için yeterlidir, ancak 5. satır mutlaka belirtildiği gibi yazılmalıdır. Ör:

<pre class="brush: csharp">var s1 = new string[3] {"John", "Paul", "Mary"};</pre>



<p><strong>Javascript:</strong>Initialization kavramı yoktur. Var ifadesi ile 
diziyi yaratmak yeterlidir. Ayrıca javascript elemanları objedir, bu nedenle 
belirli bir veritipi belirtilmez, yani aynı dizi içinde hem string hem nümerik 
elemanlar bulunabilir.</p>

<pre class="brush:js">
var cars = ["Saab", "Volvo", "BMW"];
var cars = new Array("Saab", "Volvo", "BMW"); //ancak bunun kullanılması önerilmiyor, bi üstteki hem verimlilik açısından hemn okunurluk açısından daha iyidir
</pre>


<h2>Veri ekleme</h2>
	<p><strong>VBA</strong>:Normal dizilerde eleman ekleme, ilgili indeksteki 
	elemana değer atama şeklinde gerçekleşir. Dizinin sonuna eleman ekleme, veya 
	iki eleman arsına eleman yerleştirme diye birşey yoktur. Collectionlarda, 
	listenin sonuna Add metodu ile ekleme yapılır. Dictionary'lerde de gerek Add 
	metodu ile veya dizilerdeki gibi doğruda atama yöntemiyle eleman 
	eklenebilir. Ne Colletionda ne Dictionary'de araya eleman ekleme diye birşey 
	bulunmamaktadır.</p>
	<p><strong>VB.Net ve C#:</strong>VBA için söylenenler geçerli olup dizimsi 
	yapıların her biri için farklı bir eleman eklem yöntemi vardır. Mesela List 
	için Add, Queue için Enqueue, Stack için Push gibi.</p>
	<p><strong>Javascript</strong>:Push metodu ile yeni eleman eklenebileceği gibi index numarası 
	verilerek de ilgili indexe eleman eklenebilir.</p>

<pre class="brush:js">
fruits.push("Lemon"); 
fruits[fruits.length] = "Lemon"; //bir üstteki ile aynı sonucu döndürür
</pre>



<!--<h2>Veri silme/çıkarma</h2>
<p></p>
--></div>


<h2 class="baslik">Nihai Karşılaştırma</h2>
<div class="konu">
<p>Yakında...</p>
<!--
<table class="alterantelitable">
<th>Dil</th>
<th>Syntax</th>
<th>Baz</th>
<th>Boyut/Uzunluk</th>
<th>Initialization</th>



<tr>
<td>VBA/VB.Net</td>
<td>Dim dizi(SonIndex) as String</td>
<td>0. <br> İstenirse 1 to n <br>İstenirse Option Base 1 <br>Lbound ile bulunur</td>
<td>.</td>
<td>.</td>
</tr>


<tr>
<td>C#</td>
<td>String[Boyut] dizi</td>
<td>0</td>
<td>.</td>
<td>.</td>
</tr>



</table>
-->
</div>

</asp:Content>
