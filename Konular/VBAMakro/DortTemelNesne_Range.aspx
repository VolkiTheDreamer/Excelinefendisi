<%@ Page Title='DortTemelNesne Range' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik">
	<div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'>
</asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Dört Temel Nesne'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Range Nesnesi </h1>

<p>Range nesnesi, VBA dünyasındaki en temel nesnemizdir. Kodlarımızın nerdeyse hepsinde bu 
nesneyi kullanacağız ve dahası, kodlarımızın önemli bir kısmını bu nesne ve 
türevleri oluşturacak. Bu 
nedenle bu nesnenin üyelerini yani metod ve özelliklerini(property) çok iyi 
bilmemiz gerekiyor.</p>

	<h2 class="baslik">Temeller</h2>
	<div class="konu">
	<h3>Hücrelere erişmek</h3>
	<p>Range diyince aklımıza neler gelir: Tek bir hücre, bitişik 
	olup olmaması farketmeyen hücre grubu, bir satır, birkaç satır, veya 
	sütunlar. Hemen örneklere bakalım:</p>
	<table class="alterantelitable">
	<th>Gösterim</th>
	<th>Tür</th>
	<th>Anlamı</th>	
		<tr>
			<td>Range("A1")</td>
			<td>Tek hücre</td>
			<td>A1 hücresi</td>
		</tr>
		<tr>
			<td>Range("A1:B2")</td>
			<td>Bitişik çoklu hücre</td>
			<td>A1 ile B2 arası</td>
		</tr>
		<tr>
			<td>Range("A1:B2,C3:D4")</td>
			<td>Bitişik olmayan çoklu hücre</td>
			<td>A1-B2 ve C3-D4 arası</td>
		</tr>
		
		<tr>
			<td>Range("Ozelalan")</td>
			<td>İsim verilmiş alan</td>
			<td>Name Managerde "Ozelalan" olarak belirlenen yer</td>
		</tr>
		
		<tr>
			<td>Range(Range("A1"),range("C5"))</td>
			<td>başlangıç ve bitişi ayrı belirtilmiş rangeler</td>
			<td>A1-C5 arası</td>
		</tr>
		
	</table>
	
		<p>Bir hücreyi işaret etmenin başka yolları da var. Özellikle 
	döngülü kodlarda <span class="keywordler">Cells(satırno, sütunno)</span> 
		veya <span class="keywordler">Cells(satırno, sütunharfi)</span> 
		ifadelerini çok kullanacağız. Bunun dışında bir de <span class="keywordler">
	[]</span>'ler içinde direkt hücre adresini vermek şeklinde de ulaşacağız. Ör:</p>
	<pre class="brush:vb">Cells(3,2).Select 'B3 hücresini seçer
Cells(3,"B").Select 'Bu da B3 hücresini seçer
[A1:B5].Select 'A1-B5 arasını seçer
[ozelalan].Select</pre>
	<p>Bu arada gördüğünüz üzere bir Range'i seçmek için <span class="keywordler">
	Select</span> metodunu kullanıyouruz.</p>
		<pre class="brush:vb">'Range ile döngü
For i = 1 To 10
   Range("A" & i).Value = i & ".satır"
Next&nbsp;

'Cells ile döngü
For i = 1 to 10
   Cells(i,1).Value=&nbsp;i & ".satır"
Next</pre>
	<p>Şimde de tüm sütun seçme işlemleri var</p>
	<table class="alterantelitable">
	<th>Gösterim</th>
	<th>Tür</th>
	<th>Anlamı</th>	
		<tr>
			<td>Range("A:A")</td>
			<td>Range içinde harf</td>
			<td>A kolonu</td>
		</tr>
		<tr>
			<td>Columns("A")</td>
			<td>Columns içinde harf</td>
			<td>A kolonu</td>
		</tr>
		<tr>
			<td>Columns(1)</td>
			<td>Columns içinde index</td>
			<td>A kolonu</td>
		</tr>
		<tr>
			<td>Range("A:C")</td>
			<td>Bitişik çoklu kolon</td>
			<td>A-C arası kolonlar</td>
		</tr>
		<tr>
			<td>Columns("A:C")</td>
			<td>Bitişik çoklu kolon</td>
			<td>A-C arası kolonlar</td>
		</tr>
				
		<tr>
			<td><span>Range("A:C,E:E,H:K")</span></td>
			<td>Bitişik olmayan çoklu kolon sadece Range ile</td>
			<td>A-C, E ve H-K kolonları</td>
		</tr>
		<tr>
			<td>Columns</td>
			<td>Tüm kolonlar</td>
			<td>Tüm kolonlar</td>
		</tr>
		
				
	</table>
	



	<p>Bir de tüm satır seçme işlemleri var</p>
	<table class="alterantelitable">
	<th>Gösterim</th>
	<th>Tür</th>
	<th>Anlamı</th>	
		<tr>
			<td>Range("1:1")</td>
			<td>Range içinde satır no</td>
			<td>1.satır</td>
		</tr>
		<tr>
			<td>Rows(1)</td>
			<td>Rows içinde index</td>
			<td>1.satır</td>
		</tr>
		<tr>
			<td>Range("1:5")</td>
			<td>Bitişik çoklu satır</td>
			<td>1-5 arası satırlar</td>
		</tr>
		<tr>
			<td>Rows("1:5")</td>
			<td>Bitişik çoklu satır</td>
			<td>1-5 arası satırlar</td>
		</tr>
		
		<tr>
			<td>Range("1:5,8:10")</td>
			<td>Bitişik olmayan çoklu satır sadece Range ile</td>
			<td>1-5 ve 8-10 arası satırlar</td>
		</tr>
		<tr>
			<td>Rows</td>
			<td>Tüm satırlar</td>
			<td>Tüm satırlar</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		
	</table>
	
	<p>Çoklu seçimlerde bir de <span class="keywordler">Union</span> metodu 
	kullanılır. Gerçi bu metod Range nesnesinin değil Application nesnesinin 
	metodudur ama olsun yeri geldiği için burada değinmiş olduk. (Kesişim 
	kümesini bulan 
	Intersect metodunu ise <a href="DortTemelNesne_Application.aspx">Applicaton</a> 
	sayfasında inceleyeceğiz.)</p>
		<pre class="brush:vb">Dim r1, r2, myMultipleRange As Range 
 Set r1 = Sheets("Sheet1").Range("A1:B2") 
 Set r2 = Sheets("Sheet1").Range("C3:D4") 
 Set myMultipleRange = Union(r1, r2) 
 myMultipleRange.Font.Bold = True </pre>
	<p>Bir yerde çoklu seçim olup olmadığını anlamak için
	<span class="keywordler">Areas</span> özelliği kullanılır.</p>
<pre class="brush:vb">
If Selection.Areas.Count &gt; 1 Then 
   MsgBox "Çoklu seçim var, lütfen tek bir alan seçip öyle devam edin" 
   Exit Sub
End If
</pre>
<p>Bazı durumlarda Range'leri bir değişkene atayacağız. Range nesnesini, adı 
üstünde bir nesne olduğu için <span class="keywordler">Set</span> ifadesi ile atamasını yaparız.</p>

<pre class="brush:vb">Dim alan As Range
Set alan=Range("A1:B5")
alan.Formula="=Rand()"</pre>
		<h3>Hücreleri seçmek</h3>
		<p>Rangeler aslında bir Worksheetin Range propertysinin dönen değeridir. 
		O yüzden bir sayfa ismiyle kullanılırlar, ancak çoğu zaman sayfa ismi 
		olmadan kullanırız. Bunun anlamı, o anki aktif sayfanın Range'leri 
		üzerinde işlem yapıyoruz demektir. Yani aslında <strong>Range("A1")</strong> 
		yazmak, <strong>ActiveSheet.Range("A1") </strong>yazmanın kolay yoludur.
		</p>
	<p>Yukarıdaki örneklerde gördüğünüz üzere range nesnelerini seçmek için 
	Select metodunu kullanıyoruz. Range("A1").Select gibi.</p>
		<p>Bunu kullanırken dikkat edilecek nokta, seçtiğiniz Range nesnesinin 
		aktif sayfada olmasıdır. Yani 2.sayfada iken şöyle bir seçim 
		yapamazsınız.</p>
		<pre class="brush:vb">Sheets(1).Range("A1").Select</pre>
		<p>Bunu tek satırda yapmanın farklı bir yolu var: (Biraz garip ama doğrusu bu)</p>
		<pre class="brush:vb">Application.Goto Sheets(1).Range("A1")
'veya bir diğer yol da önce ilgili sayfayı seçip sonra hücreyi seçmek olabilir
Sheets(1).Select
Range("A1").Select</pre>
		<p>Başka dosyadaki bir hücreyi ve daha bir sürü farklı seçme türünü 
		görmek için <a href="https://support.microsoft.com/en-us/kb/291308">buraya</a> göz atmak isteyebilirsiniz.</p>
		<p><strong>NOT</strong>:Excel, aynı anda birden çok sayfadaki Range'i seçmemize izin 
		vermez.</p>
		<h4>Seçme işleminin performansa etkisi</h4>
		<p>Select işlemi oldukça maliyetli bir işlemdir. Bu nedenle mecbur 
		kalınmadığı sürece seçme işlemi yapılmamalıdır. Örneğin bir hücreye 
		değer atanacaksa seçmeden de atanabilir. Aşağıdaki örnek ne demek 
		istediğimi net şekilde anlatmaktadır.</p>
		<pre class="brush:vb">
Sub secim_maliyeti()
Dim bas As Single, bitis As Single
bas = Timer
For i = 1 To 10000
    Cells(i, 1).Select 'bu varken 43 sn, yokken 1 sn
    Cells(i, 1) = i
Next i
bitis = Timer
Debug.Print bitis - bas
End Sub</pre>
		<h4>Select vs Activate</h4>
		<p>Select'e benzer bir görevi olan bir de <span class="keywordler">Activate</span> metodu vardır, ki 
		<strong>Activate 
		sadece tek bir hücreyi aktive ederken Select ile bir hücre grubunu da 
		seçebiliriz.</strong> Üstelik hali hazırda seçili bir yer varken, seçim 
		değiştirilmeden o seçim içinde bir hücre bile aktive edilebilir. 
		Aşağıda bu konuyla ilgili küçük bir örneğimiz olacak, ancak yine bu 
		konuyla alakalı olarak ActiveCell ve Selection farkına da değinelim.</p>
		<h4>ActiveCell vs Selection</h4>
		<p>O anda bulunduğumuz hücre üzerinde bir işlem yapmak istersek bir 
		Range türü olan <strong>ActiveCell</strong> nesnesini kullanırız. </p>
		<p>Bulunduğunuz yer tek bir hücre değil de hücre grubu ise <strong>
		Selection</strong> nesnesini kullanırız.</p>
		<p>Bir hücre grubu seçiliyken ActiveCell özelliği kullanılırsa o hücre 
		grubunun ilk hücresi(sol üstteki) dikate alınır.</p>
		<p>Bunların ikisi de <strong>Application</strong> nesnesinin bir propertysi olup, Application 
		ifadesi olmadan da kullanılabilirler(Application default nesne olduğu 
		için). Bu arada bu iki özellik de Range 
		<a href="Giris_ExcelNesneModeli.aspx#donendeger">objesi döndürdükleri</a> için 
		bunlardan nesne diye bahsederiz.</p>
<pre class="brush:vb">
Sub select_activate()

Range("A1:B10").Select
Debug.Print ActiveCell.Address 'sol üstteki ilk hücre, yani A1
Debug.Print Selection.Address 'tüm alan, yani A1:b10

Range("B8").Activate
Debug.Print ActiveCell.Address 'B8
Debug.Print Selection.Address 'seçili hücre B8 olmakla birlikte seçili alan hala aynıdır, değişmemiştir, o yüzden Immediate Window'da a1:b10 yazar

Range("C8").Activate 'ilk seçim alanının dışında bir hücre seçiliyor, artık selection da değişmiştir
Debug.Print ActiveCell.Address 'C8
Debug.Print Selection.Address 'C8

End Sub</pre>
		<h4>Areas</h4>
		<p>Ctrl tuşuyla seçilen ve birbirinden farklı bölgelerde 
		bulunan hücrelere <span class="keywordler">Areas</span> collection'ı ile ulaşırız.</p>
<pre class="brush:vb">
Sub areasornek()
  x = 10
  For Each alan In Selection.Areas
    alan.Interior.ColorIndex = x
    x = x + 1
  Next alan
End Sub
</pre>
		<h4>CurrentRegion ve UsedRange</h4>
		<p>Bulunduğunuz hücrenin veya o anki aktif seçili bölgenin etrafındaki tüm 
		bitişik hücrelerden oluşan bölgeyi seçmek, o bölgede işlem yapmak için
		<span class="keywordler">CurrentRegion</span> 
		özelliği kullanılır. Aynı işlemi Excelde <strong>Home&gt;Editing&gt;&gt;Find&amp;Select&gt;Go To Speacial 
		</strong>seçeneğinden aşağıdaki gibi de 
		yapabilirsiniz.</p>
		<p><img src="/images/vbarangecurrentregion.jpg"></p>
		<p><span class="keywordler">Syntax: RangeNesnesi.CurrentRegion</span> 
		şeklinde olup örnek kullanımı şöyle olabilir.</p>
		<pre class="brush:vb">ActiveCell.CurrentRegion.Select</pre>
		<p>Peki bu örnek yeterli değil. O zaman size bir ipucu. Sıralama ve 
		filtreleme işlemlerinizi Makro Kaydedici ile yaptığınızda , VBA'in size 
		ürettiği kodda sabit bir alan görebilirsiniz. Ancak sizin makronuz her 
		zaman bu sabit alan üzerinde çalışmayacaktır. O yüzden o sabit alanı 
		CurrentRegion özelliğinden faydalanarak değiştirebilirsiniz. Ama bunu 
		yapmak o kadar da basit değil. Offset ve Resize özelliklerinden de 
		faydalanmamız gerekecek. Bu örneği biraz aşağıda ilgili yere gelince 
		düzelteceğiz.</p>
		<pre class="brush:vb">Range("A2").Select
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range( _
"A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
xlSortNormal
With ActiveWorkbook.Worksheets("Sheet1").Sort
.SetRange Range("A2:F175") 'işte burayı değiştireceğiz, Range("A2:F175") yerine Range("A1").CurrentRegion yazıp deneyin, işe yaramadığını görün
.Header = xlNo
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With</pre>
		<p><span class="keywordler">UsedRange,</span> Range nesnesinin değil Worksheet nesnesinin bir özeliğidir 
		ama CurrentRegiona benzerliği nedeniyle burada bahsetmenin uygun 
		olacağını düşündüm. Bununla birbirine komşu olmayan alanları bile 
		kapsayacak şekilde, tüm <span style="text-decoration: underline">işlenmiş
		</span>alanları seçmek için kullanırlır. Burada 
		işlenmişten kastım, bir veri girişi yapılmış veya default formatı 
		değiştirilmiş bir hücredir. Mesela aşağıdaki örnekte, UsedRange 
		seçimi yapıldığında B2:G10 seçilirken</p>
		<p>
		<img src="/images/vbarangeusedrange.jpg"></p>
		<p>H11 hücresinin font büyüklüğü 1 birim artırılırsa veya Tarih formatı 
		uygulanırsa, içi boş bile olsa UsedRange bunu da kapsayacak şekilde 
		genişler ve B2:H11 olur.</p>
		<p><span class="keywordler">Syntax: WorkSheet.UsedRange</span> şeklinde olup örnek kullanımı da 
		şöyledir.</p>
		<pre class="brush:vb">ActiveSheet.UsedRange.Interior.Color = vbYellow</pre>
		<h4>Özel seçimler(SpecialCells)</h4>
		<p>Bazen de boş hücreler, formül içeren hücreler, sadece görünen 
		hücreler gibi özel seçimler yapmak isteriz. Bunlar için de
		<span class="keywordler">SpecialCells</span> özelliği kullanılır. İki 
		parametre alır ve syntaxı şöyledir.</p>
		<p><span class="keywordler">Syntax:SpecialCells(Type, Value)<br></span></p>
		<p>Type parametresinin alacağı değerler şöyledir:</p>
		<table class="alterantelitable">
			<th>Constant</th>
			<th>Anlamı</th>
			<th style="text-align: center">Numerik Değeri</th>

			<tr><td>xlCellTypeAllFormatConditions. </td>
				<td>Herhangibir formatı olan</td>
				<td style="text-align: center">-4172</td>
			<tr><td>xlCellTypeAllValidation. </td>
				<td>Validation kuralı eklenmiş hücreler</td>
				<td style="text-align: center">-4174</td>
			<tr><td>xlCellTypeBlanks.</td>
				<td>Boş hücreler</td>
				<td style="text-align: center">4</td>
			<tr><td>xlCellTypeComments. </td>
				<td>Yorumlu hücreler</td>
				<td style="text-align: center">-4144</td>
			<tr><td>xlCellTypeConstants.</td>
				<td>Sabit değer içeren hücreler(formülsüzler yani)</td>
				<td style="text-align: center">2</td>
			<tr><td>xlCellTypeFormulas.</td>
				<td>Formüllü hücreler</td>
				<td style="text-align: center">-4123</td>
			<tr><td>xlCellTypeLastCell.</td>
				<td>Ctrl+End etkisi(Son hücre)</td>
				<td style="text-align: center">11</td>
			<tr><td>xlCellTypeSameFormatConditions. </td>
				<td>Aynı conditional formatı olan hücreler</td>
				<td style="text-align: center">-4173</td>
			<tr><td>xlCellTypeSameValidation.</td>
				<td>aynı validation kuralı olan hücreler</td>
				<td style="text-align: center">-4175</td>
			<tr><td>xlCellTypeVisible.</td>
				<td>Görünür hücreler</td>
				<td style="text-align: center">12</td>
		</table>
		<p>Eğer tip olarak constant veya formül seçildiyse ikinci paremetre 
		olarak şunlar seçilebilir.</p>
		<table class="alterantelitable">
			<th>Constant</th>
			<th style="text-align: center">Numerik Değeri</th>

			<tr><td>xlErrors</td><td style="text-align: center">16</td>
			<tr><td>xlLogical</td><td style="text-align: center">4</td>
			<tr><td>xlNumbers</td><td style="text-align: center">1</td>
			<tr><td>xlTextValues</td><td style="text-align: center">2</td>
		</table>
		<p>Birkaç örnek vermek gerekirse;</p>
		<pre class="brush:vb">
'nümerik değer içeren hücrelerin arkaplan rengini sarı yapalım
Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeConstants, xlNumbers).Interior.Color = vbYellow
'formül içeren hücrelerin yazı rengini kırmızı yapalım
Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeConstants, xlTextValues).Font.Color = vbRed
</pre>
		<h4>Uç noktalar</h4>
		<p>Bir hücrenin veya hücre grubunun sanki Excel'de Ctrl+Home, Ctrl+End, 
		Ctrl+Ok tuşlarına basılmış gibi bir etki göstermesi için de çeşitli 
		propertyler var. Şimdi onlara bakalım.</p>
		<table class="alterantelitable">
			<th>Eylem</th>
			<th style="text-align: center">Yöntem</th>
			<tr><td>Ctrl+Home</td><td>Cells.SpecialCells(xlCellTypeVisible)(1).Select
				<span style="font-size: small">*</span></td>
			<tr><td>Ctrl+End</td><td>Cells.SpecialCells(xlCellTypeLastCell).Select</td>
			<tr><td>Ctrl+↑</td><td>ActiveCell.End(xlUp).Select</td>
			<tr><td>CTrl+↓</td><td>ActiveCell.End(xlDown).Select</td>
			<tr><td>CTrl+→</td><td>ActiveCell.End(xlToRight).Select</td>
			<tr><td>CTrl+←</td><td>ActiveCell.End(xlToLeft).Select</td>			
		</table>
		<p style="font-size: small">*Gizlenmiş/Filtrelenmiş bir satır/sütun 
		yoksa Range("A1").Select de bu işi görecektir</p>
		<h3>Bir hücrenin değerini okumak, ona değer atamak</h3>
	<p>Range nesnesinin <span class="keywordler">Value</span>,
	<span class="keywordler">Value2 </span>ve <span class="keywordler">Text
	</span>propertyleri vardır. Bunlardan 
	Value, bu nesnenin default özelliğidir. Default özellik şu demek, onu 
	yazmadan da kullanabilirsiniz. Mesela <strong>Range("A1")</strong> 
	ile <strong>Range("A1").Value</strong> tamamen özdeştir. 
	Ancak ben şahsen sizlere iyi bir kodlamacı özelliği olarak default özellikleri es geçmeden 
	açıkça yazmanızı öneririm.</p>
	<p><strong>Value</strong> özelliği hem okunabilir hem değer atanabilir bir özelliktir, dönüş 
	değeri Varianttır. Yani herşeyi döndürebilir, string, integer, tarih, boş.</p>
	<pre class="brush:vb">Dim v As Variant
v=Range("A1").Value 'Değeri okudum, Value özelliğini belirttim
Range("A1")="Volkan" 'Değer atadım, Value özelliğini belirtmedim, çünkü default özellik</pre>


	<p><strong>Value2</strong> özelliği Value'ya benzer, ancak Value hücrenin içindeki gerçek 
	değeri aynen verirken Value2 ise Tarih veya parabirimi formatındaki veriyi de 
	düz sayıya çevirir. </p>
		<p><strong>Text</strong> ise Excelde gözümüzle ne görüyorsak bize onu veirir, yani 
		formatlanmış halini verir.</p>
		<p>Eğer düz bir metin veya sayı ile çalışıyorsanız Value2'yi kullanmanızı 
		tavsiye ederim. İçlerinde en hızlısı ve sorunsuzu budur.</p>
		<p><strong>Ör</strong>:A1 hücresinde 21.01.1979 değeri varken hücre formatını Long Date 
		olarak değiştirdikten sonra aşğıdaki kodları çalıştıralım ve farkı 
		görelim.</p>
		<pre class="brush:vb">
Sub valuetext()
    With Range("A1")
        Debug.Print .Text '21 Ocak 1979 Pazar
        Debug.Print .Value '21 Ocak 1979
        Debug.Print .Value2 '28876
    End With
End Sub		</pre>
		<h3>Hücrenin rengi</h3>
		<p>Hücrenin arkaplanı ve yazı rengini belirlemek için sırasıyla 
		aşağıdaki kodları kullanırız.</p>
		<pre class="brush:vb">Range("a1").Interior.Color = vbRed
Range("a1").Font.Color = vbYellow
'Interior ve Font propertylerinde başka neler var, şöyle bir gözatın
'işinize yarayacak neler var aklınızda tutun, ilerde lazım olabilir</pre>
		<h3 id="pastespecial">Cut,Copy,Paste,Insert işlemleri</h3>
		<h4>Cut,Copy,Paste</h4>
		<p>Excelde yaptığımız Cut, Copy işlemleri için Range nesnesinin aynı 
		isimli metodlarını kullanıyoruz. Paste işlemi içinse Range nesnesinin 
		doğrudan bir Paste metodu yok, PasteSpecial metodu vardır. Exceldeki "Sağ 
		tık&gt;Yapıştır" veya Ctrl+V kombinasyonlarıyla yaptığımız yapıştırma 
		işleminin VBA karşılığı ActiveSheet metodunun Paste metodu olup detayına
		<a href="DortTemelNesne_Worksheet.aspx#paste">buradan</a> 
		ulaşabilirsiniz, aşağıda da küçük bir örnek bulunuyor.</p>
		<pre class="brush:vb">
Sub cutcopypaste()
    Sheets(1).Select
    Range("A5:B5").Select
    Selection.Cut 'veya Selection.Copy
    Sheets(2).Select
    ActiveSheet.Paste
End Sub	</pre>
		<h4>
		PasteSpecial</h4>
		<p>
		A5:B5 hücre grubunda formüller varsa ve bunların sadece değerlerini 
		yapıştırmak istersek işte o zaman PasteSpecial kullanırız. Aşağıda Macro 
		Recorder ile kaydedilmiş bir kodu görüyoruz.</p>
		<pre class="brush:vb">
Sub cutcopypaste()
    Sheets(1).Select
    Range("A5:B5").Select
    Selection.Copy
    Sheets(2).Select
    ActiveCell.PasteSpecial xlPasteValues
End Sub	
</pre>
		<p>
		Aslında bu işlemi yapmanın basit bir yolu daha var, Copy metodunu 
		destination parametresi ile kullanmak:</p>
<pre class="brush:vb">
Sub cutcopypaste()
  Range("A5:B5").Copy Destination:=Sheets(2).Range("A15") 'Destination tek argüman olduğ için yazmaya gerek yok
End Sub</pre>
		<p>
		PasteSpecial'in parametrelerinin alabileceği değerleri görmek için PasteSpecial yazdıktan 
		sonra boşluk veya "(" tuşuna basar basmaz Intellisense bize enumeration 
		listesini çıkarır.</p>
		<p>
		<img src="/images/vbarangefind2.jpg"></p>
		<p>
		<strong>NOT:</strong>Tek hücre için PasteSpecial'ı xlPasteValues 
		parametresiyle yapmak yerine kısayol olarak doğrudan hücrelerin 
		içeriğini eşitlenebilir.</p>
		<pre class="brush:vb">Range("A2").Value=Range("A1").Value</pre>
		<h4>Insert</h4>
		<p>Excelde yaptığımız satırları/sütunları kesip/kopyalayıp başka bir 
		satırların/sütunların arasına sokma işlemini yine aynı isimli Insert 
		metodu ile yapıyoruz.</p>
		<pre class="brush:vb">
Rows("2:2").Select
Selection.Cut
Rows("9:9").Select
Selection.Insert Shift:=xlDown 'parametre belirtilmezse Excel kendi karar verir		</pre>
		<h3>
		Formül işlemleri</h3>
		<p>
		Biz genel olarak kodlarımızda hücrelere formül yazdırmayacağız, ancak 
		ender de olsa bu işlemi yaptırmamız gerekebilecektir. Bunun için önce 
		Macro Recorder ile hücrelere birşeyler yazalım ve nasıl formül 
		ürettiğine bakalım. Gördüğünüz gibi tüm formüller, hatta metin ifadeler 
		bile <strong>FormulaR1C1</strong> özelliği ile ele alındı.</p>
		<pre class="brush:vb"> Range("H6").Select
ActiveCell.FormulaR1C1 = "volkan"
ActiveCell.FormulaR1C1 = "=TODAY()"
ActiveCell.FormulaR1C1 = "=RC[-2]*2" 'solundaki 2 hücreye referans</pre>
		<p>Biz Excel'de formüllerimizi klasik stilde(A1 stili) yazmaya alışık 
		olduğmuz için bu FormlaR1C1 yerine <strong>Formula</strong> propertysini 
		kullanırız.</p>
		<pre class="brush:vb"> Selection.Formula = "F6*2"</pre>
		<p>Son olarak dizi formülü yazmaya yarayan <strong>FormulaArray</strong> 
		var, onun kullanımı da aşağıdaki gibi olup, Exceldeki görünümü 
		{=MAX(IF(C:C&lt;100;C:C))} şeklindedir.</p>
		<pre class="brush:vb"> Selection.FormulaArray = "=MAX(IF(C[-5]&lt;100,C[-5]))"</pre>
		</div>
	<h2 class="baslik">İleri seviye Range işlemleri</h2>
		<div class="konu">

		<h3>Range'in adresi</h3>
		<p><span class="keywordler">Address</span> propertysi ile alınır. String 
		değer döndürür.</p>
		<pre class="brush:vb">Debug.Print ActiveCell.Address '$ işaretli mutlak adres
Debug.Print ActiveCell.Address(0,0) '$ işaretsiz göreceli adres</pre>
		<p>Address özelliğini <a href="Olaylar_WorksheetOlaylarievent.aspx">
		worksheet</a> olaylarında <span class="keywordler">Target.Address</span> 
		şeklinde çok kullanacağız.</p>
			<p>Keza, aktif hücrenin belirli bir adreste olup olmadığını kontrol 
			etmek için de kullanılabilir.</p>
		<pre class="brush:vb">If ActiveCell.Adress="B$1$" Then .....</pre>
		<h3>Range Property'si</h3>
		<p>Şimdiye kadar Range nesnesini, Worksheet nesnesinin bir özelliği 
		olarak kullanmış olduk. Ancak bunu bir Range nesnesinin özelliği olarak 
		da kullanabiliriz.</p>
			<p>İster tek bir hücre seçiliyken ister bir hücre grubu seçiliyken 
			olsun, Range özelliğini kullanırsak, tek hücre için kendisini, <strong>hücre 
			grubu için sol üstteki ilk hücreyi</strong> referans alarak yeni konum 
			belirlemiş oluruz ve bu referansı da A1 hücresine olan göreli farkla 
			elde ederiz. Ör: C3 hücresindeyken, Selection.Range("B1").Select 
			dediğimizde ne olur tahmin edelim:A1'e göre B1 nerdedir, bir sütun 
			sağdadır, o yüzden C3'teyken Range("B1") dersek bir sağdaki D3 
			seçilir. Keza C4:F6 hücre grubu seçiliyken 
			Selection.Range("C2").Select dersek ne olur;C2, A1'e göre 2 sağda 1 
			aşağıdadır. C4:F6'nın ilk hücresi de C4 olup 2 sağ ve 1 alttaki 
			hücresi nedir, E5.</p>
			<p>Ben şahsen Range nesnesinin Range özelliğini çok kullanmam, onun yerine 
			Item ve Offset özellikleri daha kullanışlıdır. Hemen onlara bakalım.</p>


		<h3>Item, Cells, Offset, Resize Propertyleri ve göreceli başvurular</h3>
			<p><strong>Item</strong>, Belirtilen range'in hangi satır 
			sütunundaki hücresi olduğunu döndürür. Ör: Range("B3:D6") alanının 
			1.satır, 2.sütundaki hücresini şu şekilde elde ederiz.</p>
			<pre class="brush:vb">Range("B3:D6").Item(1,2).Select 'C3 hücresi seçilir</pre>
		<p>İkinci parametre opsiyonel olup, tek parametre verilirse, ilgili 
		alanın soldan sağa kaçıncı hücresinin seçileceği belirtilmiş olur.</p>
			<pre class="brush:vb">Range("B3:D6").Item(1).Select 'sol üstteki ilk hücre yani B3 seçilir
Range("B3:D6").Item(2).Select 'C3 hücresi
Range("B3:D6").Item(5).Select 'C4 hücresi</pre>
			<p>Item özelliği Range nesnesinin default özelliği olup bunu yazmadan da 
			kullanabiliriz.</p>
			<pre class="brush:vb">Range("B3:D6")(1).Select 'Range("B3:D6").Item(1).Select demektir
Cells(2,3).Select 'Cells.Item(2,3).Select demektir. </pre>
			<p>Item özelliğinde ilgili Range'in dışına çıkabiliriz. Mesela 0 veya 
			negatif rakamlarla sola ve üste, Range'in toplam alanından daha büyük 
			rakamlarla da sağa ve aşağı doğru ilerleyebiliriz.</p>
			<pre class="brush:vb">Range("D3:F6").Item(0,-1).Select 'B2 hücresi seçilir
Range("D3:F6")(13).Select 'D7 seçilir</pre>
			<p><strong>Cells</strong> ile de Item'a benzer bir kullanımımız 
			mevcuttur. Cells, hem Range nesnesinin, hem de Worksheet nesnesinin bir propertysidir. Range ile kullanıldığında o 
			Range'in hangi 
			hücresini seçeceğimiz belirtirken worksheet ile kullanırken(veya 
			aktif sayfa için sheet adı yazmadan) tüm sayfanın kaçıncı hücresi 
			olduğunu belirtiriz.</p>
			<p>Bunların birarada kullanımına ait örnekleri aşağıda 
			bulabilirsiniz.</p>
			<pre class="=brush:vb">Sub görelisecim()
Dim alan As Range
Set alan = Range("C5:E8")

alan.Select ' tamamı
alan.Range("A1").Select
alan.Item(0, 0).Select ' bir satır sol üst
alan.Item(1, 1).Select 'sol üstteki ilk hücre
alan.Item(1).Select 'sol üstteki ilk hücre
alan.Item(0).Select 'sol üstteki ilk hücrenin bir solu
alan.Cells(1, 1).Select 'sol üstteki ilk hücre
alan.Cells.Select 'tamamı
alan.Offset(1, 1).Select 'toplam alan büyüklüğü aynı kalack şekilde bir aşağı bir sağa kaydır
alan.Offset(0, 0).Select 'tamamı
alan.Offset().Select 'tamamı
alan(1, 1).Select 'item gibi davranır
alan(0, 0).Select 'item gibi davarnır
alan(1).Select 'item gibi davranır
alan(0).Select 'item gibi davranır

End Sub</pre>
			<p>Gördüğünüz gibi Item ile Cell hep aynı sonucu veriyor. O halde 
			neden iki ayrı property var diye düşünüyor olabilirsiniz. Cevap:Item aslında 
			Range'e ait specific bir özellik değildir. Item, 
			collection tarzı tüm&nbsp; nesnelerin genel bir özelliği olup, 
			collectionlardaki elemanların her birini ifade eder. Range de bir 
			hücreler koleksiyonu olduğu için bu özelliği devralmıştır.</p>
			<p>Bu arada farkettiniz mi bilmiyorum ama bu tek indeks verme olayı, 
			bir sütunda aşağı doru ilerlemek için güzel bir fırsat sunuyor. M<span>esela&nbsp;</span>Range("A1")(1)<span>&nbsp;</span>A1 
			hücresini ifade ederken,<span>&nbsp;</span>Range("A1")(5)<span> A</span>5,<span>&nbsp;</span>Range("A4")(11)<span>&nbsp;de 
			A</span>14 hücresini. Bu yöntem az sonra göreceğimiz Offsetin güzel 
			bir alternatifi olmaktadır.</p>
			<p><strong>Offset(x,y)</strong> ile referans verilen bir range'in x satır sağına(x 
			negatifse soluna) ve y sütun altına(y negatifse üstüne) gideceğimize karar 
			veririz.</p>
			<p><span class="keywordler">Syntax:RangeObject.Offset(satır,sütun)</span> 
			şeklindedir.</p>
			<p>Burda Range tek bir hücre olabildiği gibi, hücre grubu veya bir 
			satır/sütun da olabilir</p>
			<pre class="brush:vb">Sub offsetornek()
Range("C2").Offset(1, 0).Select 'C3
Range("C2").Offset(-1, 2).Select 'E1
Range("C2").Offset(0, -2).Select 'A2
Range("C2").Offset(0, 0).Select 'C2

Range("C2").EntireRow.Offset(1).Select '3.satır. Range("C2").EntireRow.Offset(1,0).Select ile aynıdır. İkinci parametre yoksa 0 anlamındadır
Range("C2").EntireRow.Offset(-1).Select '1.satır
Range("C2").EntireColumn.Offset(, -1).Select '2.sütun. Range("C2").EntireColumn.Offset(0, -1).Select ile aynıdır. ilk paramterde 0 yerine boş da geçilebilir

Range("C2:F6").Offset(1, 1).Select 'D3:G7 seçilir
End Sub</pre>
			<p id="resize"><strong>Resize</strong> ile bir Range nesnesi yeniden 
			boyutlandırılır. </p>
			<p><span class="keywordler">Syntax:Resize(satırboyutu, sütunboyutu)</span>. 
			İki 
			parametre de opsiyonel olup belirtilmezlerse olduğu yerde kalır.</p>
			<pre class="brush:vb">Sub resizeornek1()
Range("C3:G7").Select
Selection.Resize(Selection.Rows.Count - 1, Selection.Columns.Count + 2).Select

Range("C3:G7").Resize.Select ' aynen kalır
Range("C3:G7").Resize().Select 'aynen kalır
Range("C3:G7").Resize(1).Select 'kolon parametresi yok, o yüzden aynı kalır, satır ise 1 satır olacak şekilde daralır
Range("C3:G7").Resize(, 2).Select 'satır parametresi yok, o yüzden aynı kalır, sütun ise 2 sütun olacak şekilde daralır

End Sub</pre>
			<p>Resize'ın güzel kullanımlarından biri de bir range'in olduğu gibi 
			bir diziye atanıp başka bir yere kopyalanmasındadır. Aşağıdaki 
			örneğe bakalım. A1:C16 arasındaki herşey E1:G16 alanına aktarılacak. 
			İşlemin hangi kodla yapıldığını <a href="#rangedizi">buradan</a> 
			görebilirsiniz.</p>
			<p><img src="/images/vbarangeresize.jpg"></p>
			<p>İşte şimdi bir de offset ve resize'ın birlikte kullanımına bir 
			örnek. (Globaliconnect.com sitesinden alınmıştır)</p>
			<pre class="brush:vb">Sub RangeOffsetResize1()
'form a pyramid of numbers, using Offset &amp; Resize properties of the Range object - refer Image 7a.

Dim rng As Range, i As Integer, count As Integer

'set range from where to offset:
Set rng = Range("A1")
count = 1

For i = 1 To 5
'range will offset by one row here to enter the incremented number represented by i - see below comment:
Set rng = rng.Offset(count - i, 0).Resize(, i)
rng.Value = WorksheetFunction.RandBetween(10 ^ (i - 1), 10 ^ i)
'note that 2 is added to i here and i is incremented by 1 after the loop, thus ensuring that range will offset by one row and the incremented number represented by i will be entered in the succeeding row:
count = i + 2
Next

End Sub</pre>
			<p>Şimdi de, biraz yukarda CurrentRegion özelliğinde bir örneğimiz 
			vardı, yeri gelince düzelteceğiz demiştik. Yapacağımız şey, bir 
			listeyi başlığı hariç seçmek olacak. Burda yukardaki örnekten farklı 
			olarak önce Resize sonra Offset yapacağız.
</p>
			<pre class="brush:vb">enalt = Range("A1").End(xlDown).Row
Set alan = Range("A1").CurrentRegion.Resize(enalt - 1).Offset(1)
Range("A2").Select
ActiveWorkbook.Worksheets("calculatedlar").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("calculatedlar").Sort.SortFields.Add Key:=Range( _
"A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
xlSortNormal
With ActiveWorkbook.Worksheets("calculatedlar").Sort
.SetRange alan 'Makro recordardan sabit gelen bu kısmı değiştirdik
.Header = xlNo
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With
</pre>
			<h3>Column, Columns, Row, Rows</h3>
			<p>
			<strong>Columns</strong>:Bir Range nesnesinin kolonlarını ifade 
			eder. Genelde count özelliği ile kullanılır ve ilgili alanda kaç kolon seçili olduğu 
			elde edilir. Ayrıca ilgili alandaki belirli kolonlar üzerinde 
			formatting işlemleri yapılabilir. Tabi burda kolondan kastımız tüm bir 
			kolonunun seçimi değil, ilgili alandaki dikey blokların seçimidir. Örneğin 
			A2:B10 alanı içinde Columns(2) seçildiğinde sadece B2:B10 seçilir, tüm 
			B kolonu değil.</p>
			<p>
			<strong>Column</strong>:Tek bir hücre sözkonusuysa onun bulunduğu 
			kolonun index numarası, birden çok hücre grubu sözkonusuya ilk 
			hücresinin bulunduğu kolonun index numarası döner.</p>
			<p>
			<strong>Rows</strong> ve <strong>Row</strong> de Columns ve Column'un 
			satır versiyonudur.</p>
			<p>
			Mesela şu tablo;</p>
			<p><img src="/images/vbarangecolumns1.jpg"></p>
			<pre class="brush:vb">Sub ZebraYap()
Dim alan As Range

Set alan = Range("B3:D11")
For i = 1 To alan.Rows.Count
  If i Mod 2 = 1 Then
    alan.Rows(i).Interior.Color = vbBlue
  Else
    alan.Rows(i).Interior.Color = vbWhite
  End If
Next i

alan.Columns(1).Font.Bold = True

End Sub</pre>
	<p>kodu çalıştıktan sonra böyle görünür</p>
			<p><img src="/images/vbarangecolumns2.jpg"></p>
		<p>Peki bir hücrenin bulunduğu tüm satır veya sütunu seçmek isteseydik? 
		O zaman da <span class=" keywordler">EntireRow ve EntireColumn</span> özelliklerini 
		kullanırız.</p>
		<pre class="brush:vb">ActiveCell.EntireColumn.Select 'aktif hücrenin bulunduğu tüm kolonu seçer
ActiveCell.EntireRow.Select 'aktif hücrenin bulunduğu tüm satırı seçemer
</pre>
			<p>Item özelliğinde olduğu gibi Columns için de range'in sınırları 
			dışındaki index numarları verilebilir. Bu da bize kolonlarda tek tek 
			ilerlemenin kolay yöntemlerinden birini sunmuş olur. 
			Ör:Range("A1").Columns(2), B2 hücresine ilerlemenin bir yoludur ve 
			döngüsel işlemlerde bi yerden bi yere programatik olarak 
			ilerlememizi sağlar.</p>
			<h3>Satır/Sütun ekleme, silme, gizleme</h3>
			<p>Makro kaydedicisi ile birkaç işlem yaptım ve kodları aşağı aldım.</p>
			<pre class="brush:vb"> Columns("E:E").Select
'Şimdi kolon eklenecek
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("F:F").Select
'Şimdi kolon silinecek
Selection.Delete Shift:=xlToLeft
Rows("4:4").Select
'Satır eklenecek
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'Satır silinecek
Selection.Delete Shift:=xlUp
Columns("E:E").Select
'Gizlemek için bir eylem yerine Gizli özelliğine True değeri veriliyor
Selection.EntireColumn.Hidden = True
Columns("D:F").Select
'Gizli olan kolonu açmak için de Gizli özelliğine False değeri atanıyor
Selection.EntireColumn.Hidden = False</pre>
			<h3><a name="rangedizi"></a>Range'i bir diziye atama</h3>
			<p>Belli bir rangedeki değerleri almanın en hızlı yolu burayı
			<strong>Variant</strong> türünde bir diziye atamaktır. Ancak tek 
			boyutlu bir dizi değil, satır ve sütunu ifade eden iki boyutlu bir 
			dizi.</p>
			<p><a href="DizilerveDizimsiYapilar_DizilerArray.aspx">Diziler bölümünde</a> detaylı göreceğiz, bir dizinin eleman sayısı 
			<strong>Ubound</strong> metodu ile elde edilir. n. boyutundaki eleman sayısı da 
			Ubound(dizi, n) ile. Bu case'de 1.boyutumuz satırları 2.boyutumuz 
			ise kolonları ifade ediyor.</p>
			<pre class="brush:vb">Sub range_diziata()

Dim siciller As Variant
siciller = Range(Range("a1"), Range("a1").End(xlDown).End(xlToRight)).Value
'colon sayısı:ubound(siciller,2)
'row sayısı:Ubound(siciller,1) veya kısaca ubound(siciller), 1 defaulat value
Debug.Print UBound(siciller, 2)
Debug.Print UBound(siciller, 1)

'bunu aynen bi yere yapıştırmak için, aynı boyutta olması lazım, bunu da resize ile hallederiz
Range("h1").Resize(UBound(siciller), UBound(siciller, 2)).Value = siciller

End Sub</pre>
		<h3>Veri depolama aracı olarak Range</h3>
			<p>Masaüstü veya Web tabanlı programlamayla uğraşanlar bilirler, 
			aynı oturumdayken programı birkaç kez çalıştırdığımızda önceki 
			çalıştırmalarda elde ettiğimiz bir değeri daha sonra kullanmak için 
			Visible özelliği False olan yani gizli olan bir Textbox'a(veya 
			Labela) bu değeri atarız. Mesela bir tuşa kaç kez basıldığını böyle 
			bir kutu içinde depolayabilir ve 10.kez basıldığında kullanıcıya bir 
			mesaj gösterebilir veya Programdan çıkışı sağlayabilir veya başka 
			bir kodun çalışmasını sağlayabilirsiniz.</p>
			<p>İşte Excel'de de Range nesnesini bu amaçla kullanabilirsiniz. 
			Mesela, bir düğmeye ilk kez basıldığında uzun bir kodun çalışmasını, 
			sonraki basışlarda ise daha kısa kodların çalışmasını 
			sağlayabiliriz. Örnek kod aşağıdaki gibi olacaktır. Geçici depolama 
			işini Z1 hücresinde yapacağız(ilk açılışta görünmeyen bir hücre 
			olması nedeniyle. Mümküse font rengini de beyaz yaparız, veya Z 
			kolonunu gizleriz.)</p>
<p>NOT:Bu yöntem <a href="Temeller_DegiskenlerveVeriTipleri.aspx#static">static değişken</a> tanımlamanın bir alternatifidir.</p>
			<pre class="brush:vb"> Sub dugme_click()
'Z1 ilk başta boştur, yani 0'dır.

If Range("Z1").Value2&gt;1 Then
'Sadece son sayfadaki Table'ı refresh et. 10 sn sürer
Else 'Z1=0 ise yani düğmeye ilk kez basılıyorsa
'Tüm sayfalardaki connectionları refresh et. 5 dk sürer.
End If

Range("Z1").Value2=Range("Z1").Value2+1 'Z1dei değeri bir artırıyoruz
End Sub</pre>
			<p> Bir başka örnek de Workbook_Close ve Workbook_Deactivate olay 
			metodlarını bir seferde ele almak olabilir. Bu örneğe
			<a href="Olaylar_WorkbookOlaylarievent.aspx">bu sayfada</a> yer 
			verilmiştir.</p>
			<h3> Find ve Replace</h3>
			<p>Excel içindeyken Ctrl+F(veya Home&gt;Find) yaptığımızda 
			karşımıza çıkan <strong>Find</strong> dialog kutusunu ve sonrasında 
			yaptığımız bulma işlemini Range nesnesinin <strong>Find </strong>
			metodu ile yapıyoruz. Tahmin ettiğiniz üzere bu metod da yine bir 
			range nesnesi döndürür, dönen nesne de aranan şeyin ilk bulunduğu 
			hücredir. Eğer aranan değer bulunmazsa <strong>Nothing</strong> 
			döner.</p>
			<p> En basit haliyle aşağıdaki gibi kullanılır.</p>
			<pre class="brush:vb">
Dim arabul As Range
Set arabul = Range("A1").CurrentRegion.Find("Volkan")</pre>
			<p>
			<img src="/images/vbarangefind1.jpg"></p>
			<p>Gördüğünüz gibi hiç parametre kullanmadık. Böyle bir durumda 
			hangi değerler baz alınır? Şimdi bi genel kültür bilgisi verelim. 
			Siz Excel'de bu araçla çalışırken en son hangi değerleri 
			kullanıdysanız bir sonraki aramanızda da bu değerler kullanılır, 
			çünkü arama ayaralarınız kaydedilir. VBA'de de olan budur. Yani 
			parametresiz Find kullanırsanız Excelde en son kullandığınız arama 
			kriterleri kullanılır. O yüzden tavsiyem bu kriterleri açıkça 
			belirterek yazmanızdır. Aşağıdaki gibi:</p>
			<pre class="brush:vb">
Dim arabul As Range
Set arabul = Range("A1").CurrentRegion.Find( _
		What:="volkan", _
		After:=ActiveCell, _
		LookIn:=xlFormulas, _
		LookAt:=xlPart, _
		SearchOrder:=xlByRows, _
		SearchDirection:=xlNext, _ 
		MatchCase:=False, _
		SearchFormat:=False)
</pre>
			<p>Parametrelerin açıklamalarını aşağı yukarı tahmin ediyorsunuzudur, 
			zira bunları zaten Excel içinde sık sık kullanıyorsunuz. Yine de 
			açıklamlara ve alabileceği değerlere
			<a href="https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-find-method-excel">
			buradan</a> bakabilirsiniz. (Bunlardan bir tek L<span lang="TR">ook</span>A<span lang="TR">t 
			param</span>e<span lang="TR">tresi kafa kar</span>ı<span lang="TR">şt</span>ı<span lang="TR">rıc</span>ı 
			olabilir<span lang="TR">, çünkü </span>Find <span lang="TR">dialog</span> 
			kutusun<span lang="TR">daki </span>M<span lang="TR">atch</span> C<span lang="TR">ase
			</span>seçeneği <span lang="TR">için aynen </span>M<span lang="TR">atch</span>C<span lang="TR">ase 
			parametresi varklen, </span>M<span lang="TR">atch entire </span>cell 
			cotent için M<span lang="TR">atch</span>E<span lang="TR">ntire</span>Content 
			diye bir parametre beklerken<span lang="TR"> </span>bunun <span lang="TR">yerine
			</span>L<span lang="TR">ook</span>A<span lang="TR">t parametresi var</span>.)</p>
			<h4>Nothing</h4>
			<p>Dedik ki aradığımız değer bulumazsa Nothing döndürür, bunu da 
			aşağıdaki gibi sorgulayabiliriz.</p>
			<pre class="brush:vb">
If Not arabul Is Nothing Then
    arabul.Select    
End If			</pre>
			<h4>
			Diğer Find detayları</h4>
			<p>
			Tüm sayfalarda arama yapmak için For Each döngüsü ile sayfalarda 
			dolaşmak gerekebilir.</p>
			<pre class="brush:vb">
Dim ws As Worksheet
Dim arabul As Range
For Each ws In ActiveWorkbook.Sheets
    ws.Activate 'veya select
    Set arabul = Cells.Find( _
        What:="volkan", _
        After:=ActiveCell, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    If Not arabul Is Nothing Then
        arabul.Select
        Exit For
    End If
Next ws</pre>
			<p>
			Aradığımız değerin istediğimiz sonucu getirmediğini görür ve aramaya devam 
			etmek istersek <strong>FindNext</strong> metodu kullanılır, bu metod 
			tek bir After parametresi alır.</p>
			<pre class="brush:vb">Cells.FindNext(After:=ActiveCell).Activate</pre>
			<p><span>Formatlı arama için <strong>Application.FindFormat</strong> 
			propertyleri kullanılır ve SearchFormat=True yapılır. Record makro 
			ile detaylara bakabilirsiniz.</span></p>
			<h4>Replace</h4>
			<p>Replace metodunun kullanımı Find'a benzer. Find'dan farklı olarak
			<span>Range nesnesi değil Booelan döndürür. Sonuç True ise işlemi 
			yapar, False ise birşey yapmaz. O anda bulunulan hücrenin yeri 
			değişmez.</span></p>
			<pre class="brush:vb">
 Cells.Replace _
 	What:="ali", _
 	Replacement:="veli", _
 	LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False						</pre>
			<h3><a name="Calculation">Belirli bir hücre grubunun değerini hesaplatma</h3>
			<p> Application konusunu anlatırken göreceğimiz ve üzerinde epeyce 
			duracağımız konulardan biri da Calculation işlemleridir. Calculate 
			metodu Application nesnesinde var, Worksheet nesnesinde var, 
			Workbook nesnesinde yok ama bir döngü ile sağlanabiliyor. Range'te 
			yok mu? Tabiki var. Diyelim ki, Calculation durumu Manuel set 
			edilmiş durumda ve sayfanızın münferit yerlerinde tonla formül var, 
			siz sadece belirli bir grup hücrede formüllerin hesaplanmasını 
			isteyebilirsiniz. Bunun için küçük bir kod yazıp bunu 
			QuickAccessToolbarınıza koyabilirsiniz. İşte kodumuz.</p>
			<pre class="brush:vb">Sub rangecalc()
   Selection.Calculate
End Sub</pre>
			<h3>Parent özelliği</h3>
			<p>Bazen bir hücrenin hangi sayfada olduğunu elde etmek isteriz. 
			Bunun için hiyerarşide bir üst basamağa çıkmamızı sağlayan Parent 
			özelliğini kullanırız.</p>
			<pre class="brush:vb">Debug.Print TypeName(Activecell.Parent) 'Worksheet
Debug.Print Activecell.Parent.Name 'ilgili Worksheetin adı</pre>
</div>

<h2 class="baslik"><a name="Ornekler"></a>Çeşitli Örnekler</h2>
<div class="konu">
	<h4 class="baslik">Bir alanı başlık hariç seçmek</h4>
	<div>
	<p>Sıklıkla bir alanı başlık hariç seçmeniz gerekebilecektir. Arkasından bu alanla ne yapmak istiyorsanız yapabilirsiniz.
	Bunun için aşağıdaki gibi bir fonksiyon tanımlayıp seçili alanı bu fonksiyona parametre olarak gönderebiliriz.</p>
	
	<pre class="brush:vb">
	Function başlıkhariçalan(alan As Range) As Range
	    Dim enalt As Long
	    enalt = alan.Rows.Count
	    Set başlıkhariçalan = alan.Resize(enalt - 1).Offset(1)
	End Function
	</pre>
	
	<img src="/images/vbarangeornek1.jpg">
	
	<p>Fonksiyonun kullanımı oldukça basit.</p>
	<pre class="brush:vb">		
	'diyelim ki o anda ilgili alanı seçmişiz
	başlıkhariçalan(Selection).Select
	'veya bulunduğumuz hücrenin CurrentRegionunı da gönderebiliriz
	başlıkhariçalan(Activecell.CurrentRegion).Select
	</pre>
	
	<img src="/images/vbarangeornek2.jpg">
	
	</div>
	
	<h4 class="baslik">Filtrelenmiş bir alandaki ilk visible(görünür) hücreyi/satırı seçmek</h4>
	<div>
	<p>Bazen elinizdeki listeler/tablolar üzerinde filtreleme işlemi yapıp hemen arkasından da filtrelenmiş bu kısmı
	başlık hariç olarak kullanmanız gerekecektir. Böyle bir durumunda, az önceki gibi başlığın hemen bir altındaki
	hücreye ilerlememiz yeterli olmayacaktır. Zira görünen ilk hücre aşağıdaki gibi 33.satırda olabilir.</p>
	
	<img src="/images/vbarangeornek3.jpg">
	
	<p>Bu örnekte bizim önce 33.satıra gelebilmemiz lazım. Ondan sonrası kolay: Önce en sağa, ve ordan da en aşağıya kadar seçeriz.</p>
	
	<pre class="brush:vb">
Function ilkvisiblesonrasıalan(alan As Range) As Range
    Dim ilk As Range
    Dim son As Range
    Dim n As Integer
    n = alan.Columns.Count
    Set ilk = alan.Offset(1, 0).Resize(alan.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Cells(1, 1) 'bu kısım _
    ilk görünen hücreyi verir
    Set son = ilk.Offset(0, n - 1)
    Set ilktoright = Range(ilk, son)
    Set ilkvisiblesonrasıalan = Range(ilktoright, ilktoright.End(xlDown))
End Function

'kullanım yine oldukça basit
ilkvisiblesonrasıalan(Selection) 'eğer ki hali hazırda alan seçiliyse
ilkvisiblesonrasıalan(Activecell.CurrentRegion) 'alan henüz seçili değilse
	</pre>

<p>ve sonrasında durum şöyledir:</p>

<img src="/images/vbarangeornek4.jpg">

	</div>

	<h4 class="baslik">Filtre uygulanmış bir alandaki son hücreyi elde etmek</h4>
	<div>
	<p>Aşağıdaki gibi filtre uygulanmış bir alandaki görünen son hücreyi elde 
	etmek kolaydır. Ya A1'den aşağı indirip([A1].End(xlDown)) veya sayfanın en 
	altındaki hücreden (1048576. hücreden) yukarı çıkarak. Böylece 13'üncü 
	satırı veya A13 hücresini elde ederiz.</p>
		<p>Ancak bazen biz bu örnekten konuşacak olursak 34.satır veya A34 
		hücresi gerekir. Bunun için bir fonksiyon yazacağız.</p>
<img src="../../images/vbarangefiltreornek1.jpg">
		<pre class="brush:vb">
Function FiltredekiSonHucre(ws As Worksheet)

On Error GoTo hata
FiltredekiSonHucre = ws.Range(Split(ws.AutoFilter.Range.Address, ":")(1)).Row

Exit Function

hata:
End Function			</pre>
		<p>Fonksiyonun yaptığı iş basit. F8 ile ilerlerken baktığımızda filtreli 
		alanın adresini aşağıdaki gibi görüyoruz: $A$1:$P$34. Bunu bir metin 
		olarak el alıp ayracı ":" olacak şekilde Split ile bir diziye 
		dönüştürüyoruz. Sonrasında elde edilen dizinin 1.indeksindeki yani 
		2.elemanını alıyoruz. 1.eleman $A$1 olup ikinci eleman $P$34'tür. Sonra 
		bu stringi Range içine koyup bunu hücre olarak elde edip onun da Row'unu 
		döndürüyoruz.</p>
		<p>
		<img src="../../images/vbarangefiltreornek2.jpg"></p>
	</div>

	<h4 class="baslik">Otomatik mail gönderiminde body'ye Excel hücreleri yapıştırma</h4>
	<div>
	<p>Bu olağanüstü faydalı makroyu üstatlardan Ron de Bruin'in 
	<a href="http://www.rondebruin.nl/win/s1/outlook/bmail2.htm">sayfasından</a> aldım. Hiç değiştirmeden aynen alıntılıyorum.
	Siz de aynen bu şekilde kullanabilirsiniz.
	</p>
	
	<pre class="brush:vb">
Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

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
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
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
End Function	
	</pre>
	
	<p>Fonksiyonun kullanımı oldukça basit. Yine aynı sitede örnek kullanım da var. Ancak isterseniz sitemin
	Outlook otomasyonuyla ilgli <a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx">sayfalara</a> da bakabilirsiniz.
	</p>
	</div>

</div>
</asp:Content>
