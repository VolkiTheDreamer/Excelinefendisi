<%@ Page Title='Temeller Operatorler' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' 
runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Temeller'>
</asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Operatörler</h1>
<h2>Aritmetik Operatörler</h2>
<p>Matematikte bildiğimiz 4 işlem operatörü VBA'de de aynen geçerlidir. Bunlara ek olarak;</p>
	<ul>
		<li><span class=" keywordler">"\"</span> tam sayı bölme işareti olup, bir tamsayılı bölme sonunda böleni verir. Ör: 9 \ 6=1, başka bir örnek:14 \ 3 =4. </li>
        <li><span class="keywordler">Mod</span> operatrörü ise tamsayılı bölme işleminde kalanı verir. Ör: 9 Mod 6=3, 14 Mod 3= 2. Küçük bir sayının büyük sayıyla Mod'u küçük 
		sayının kendisidir.</li>
		<li><span class=" keywordler">^</span> işareti: Üs aldırır. 3^2=9</li>
	</ul>
	<h2>Karşılaştırma Operatörleri</h2>
	<table class="alterantelitable">
		<th>Amaç</th>
		<th>Operatör</th>
		<th>Örnek</th>		
		<tr>
			<td>Eşit mi</td>
			<td>=</td>
			<td>If A=B then ....</td>
		</tr>
		<tr>
			<td>Büyük mü</td>
			<td>&gt;</td>
			<td>If A&gt;B then ....</td>
		</tr>
		<tr>
			<td>Küçük mü</td>
			<td>&lt;</td>
			<td>If A&lt;B then ....</td>
		</tr>
		<tr>
			<td>Büyük eşit mi</td>
			<td>&gt;=</td>
			<td>If A&gt;=B then ....</td>
		</tr>
		<tr>
			<td>Küçük eşit mi</td>
			<td>&lt;=</td>
			<td>If A&lt;=B then ....</td>
		</tr>
<tr>
			<td>Eşitsizlik</td>
			<td>&lt;&gt;</td>
			<td>If A&lt;&gt;B then ....</td>
		</tr>
	</table>
	<p>Özel yazımı olan bir kontrol şekli var, o da True/False kontrolü. Bu 
	kontrolü yaparken direkt boolean tipli değişkenin kendisini yazarak 
	sorgulayabiliriz. Ör:</p>
	<pre class="brush:vb">
Sub bool_andor()
Dim a As Boolean
a = True

If a And (x = 0 Or y = 1) Then 'if a=True demek yerine
    MsgBox "Doğru"
Else
    MsgBox "yanlış"
End If

End Sub	</pre>
<h2>Mantıksal Operatörler</h2>
	<table class="alterantelitable">
		<th>Amaç</th>
		<th>Operatör</th>
		<th>Örnek</th>		
		<tr>
			<td>Ve</td>
			<td>And</td>
			<td>If A=B and A&gt;0 then ....</td>
		</tr>
		<tr>
			<td>Veya</td>
			<td>Or</td>
			<td>If A&gt;B or A=0 then ....</td>
		</tr>
		<tr>
			<td>Değil</td>
			<td>Not</td>
			<td>If Not obj Is Nothing then ....</td>
		</tr>
		</table>
		<p><strong>Not</strong> operatörünün ilginç bir kullanımı da boolean tipli değişkenleri tersine döndürmek içindir. Özellikle toggle işlemlerinde(Ör:Bi düğmeye defalarca basıldığında True/False döngüsüne girme durumu) çok kullanılır.</p>

	<pre class="brush:vb">
Sub bool_not()
Dim a As Boolean
a = True 

a= not a ' a şimdi False oldu

End Sub	</pre>
<h2>Birleştirme operatörleri</h2>		
<p>İki tür birleştirme operatörü var.</p>
	<ul>
		<li><span class=" keywordler">+</span>: Bu operatör iki numerik ifadeyi toplarken iki string ifadeyi 
		birleştirir.</li>
		<li><span class=" keywordler">&amp;</span>:Bu hem numerik hem string değişkenleri birleştirir</li>
	</ul>
	<p>
		+ işareti kullanıldığında değişkenlerden biri string tipte olsa bile 
		eğer içeriği sayı ise birleşme yerine toplama olur. Aşağıda örnekler 
		mevcut.
<pre class="brush:vb">
Sub birlestirme()
	Dim a As String
	Dim b As String
	Dim c As Integer
	Dim d As Integer
	Dim e As String
	a = "10"
	b = "20"
	c = 300
	d = 5000
	e = "volkan"
	
	Debug.Print "merhaba " + e 'iki string + ile birleşir
	Debug.Print "merhaba " &amp; e 'iki string &amp; ile birleşir
	Debug.Print a + b 'iki sayısal içerikli string + ile birleşir&gt;1020
	Debug.Print a &amp; b 'iki sayısal içerikli string &amp; ile birleşir&gt;1020
	Debug.Print a + c 'bir sayısal içerikli string ve bir numerik + ile toplanır&gt;310
	Debug.Print a &amp; c 'bir sayısal içerikli string ve bir numerik &amp; ile birleşir&gt;10300
	Debug.Print c + d 'iki numerik + ile toplanır&gt;5300
	Debug.Print c &amp; d 'iki numerik &amp; ile birleşir&gt;3005000
	'Debug.Print c + e 'hata verir, numerik ve sayısal içerikli olmayan string toplanamaz
End Sub</pre>

	<h3>Değişkenleri<a name="selfcombine"></a> kendisiyle toplama/birleştirme</h3>
	<p>Bir InputBox/MsgBox içindeki veya otomatik mailingdeki Body metni çok 
	uzun ise bu metni parçalar halinde yazıp bunları sürekli kendisiyle 
	birleştirerek ilerlemek yaygın bir yöntemdir. </p>

<pre class="brush:vb">
Sub satırgeçiş()

mesaj = "Müşteri segmenti için bir değer giriniz. " &amp; vbCrLf
mesaj = mesaj + "Bireysel müşteriler için 1," &amp; vbCrLf
mesaj = mesaj + "Ticari müşteriler için 2," &amp; vbCrLf
mesaj = mesaj + "Kurumsal müşteriler için 3"

a = InputBox(mesaj)

End Sub </pre>

	<p>Bir başka örnek de şöyle olabilir</p>

<pre class="brush:vb">
Sub mailbodyornek()

bodymsj="Değeri arkadaşlarımız" &amp; Chr(10) &amp; Chr(10)
bodymsj=bodymsj+"........."
bodymsj=bodymsj+"........"
bodymsj=bodymsj+"........"

'Diğer kodlar
End Sub</pre>
	<p>
	Bir de sayısal değişkenlerin kendisiyle toplanması vardır. Bu yöntemi de 
	özellike döngüsel yapılar içinde kullanırız. Değişkenin kendisini 1 ile 
	toplayarak, değerini artırmış oluruz. </p>
	<pre class="brush:vb">Sub sayıartır()
Dim i As Integer

i=0

Do
  'Diğer kodlar
   i=i+1  'burada i'yi her defasında bir artırmış oluyoruz. Gelişmiş dillerdeki i++ ifadesinin aynısıdır
Loop Until i=100
End Sub</pre>

	</asp:Content>
