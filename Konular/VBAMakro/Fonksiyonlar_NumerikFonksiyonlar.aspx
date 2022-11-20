<%@ Page Title='Fonksiyonlar NumerikFonksiyonlar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Nümerik Fonksiyonlar</h1>
<div class='konu'>
<p> Excel VBA'de kullanılmak üzere çok fazla nümerik fonksiyon bulun<span style="text-decoration: underline">ma</span>maktadır.</p>
	<p> Bununla birlikte Excel'in kendisiyle gelen nümerik fonksiyonların 
	birçoğunu <strong>WorksheetFunction</strong> fonksiyonu aracılığı ile kullanabiliriz. Bununla 
	ilgili detaylı bilgiye ilgili <a href="Fonksiyonlar_WorksheetFunction.aspx">
	sayfada</a> değineceğiz.</p>
	<p> Şimdi VBA içindeki birkaç faydalı numerik fonksiyona bakalım.</p>
	<h3> Dönüştürme fonksiyonları</h3>
	<p> <strong>Ön bilgi</strong>:Türkiyede biz "," işaretini ondalık ayraç olarak kullanırken 
	ABD ve dolayısıyla bir ABD firmasnın ürünü olan Excel ve VBA "." işaretini ondalık ayraç 
	olarak kullanır. Binlik ayraç olarak ise tersi. Gerçi regional settings(bölgesel 
	ayarlar) 
	ayarlaması yapıldığında Excelin kendisi de(Sadece Excel, VBA değil) ondalık ayraç olarak "," 
	kullanabilmektedir. Ve nitekim Türkiye'deki bilgilsayarların çoğunda bu ayar 
	yapılı haldedir. Ancak VBA "." karakterini her zaman ondalık ayraç için 
	kulanırken, binlik ayraç diye birşey kullanmaz. VBA dünyasında "," karakteri 
	parametre ayracı olarak kullanılır. Dolayısıyla VBA'deki 2.22 ifadesi Türkçe 2,22 
	algılanırken 2,22 ifadesi ise 2 ve 22 şeklinde iki ayrı parametre olarak algılanır. 
	Bu 2,22 ifadesi tırnak içinde değilse birçok durumda tek parametre alan 
	fonksiyonlarda hataya neden olurken, tırnak içindeyse metinsel bir ifade 
	olarak algılanır.</p>
	<p> <span class="keywordler">Val(String)</span>: Parametre olarak aldığı 
	stringi rakama çevirir. Dönüş tipi Double'dır.</p>
	<p> Val fonksiyonu, dönüştüreceği değer içindeki ilk sayısal olmayan kısıma kadar olan kısmı sayıya çevirir. 
	Ör:123asr45 değerini 123'e çevirir.</p>
	<pre class="brush:vb">Debug.Print Val("00123") '123
Debug.Print Val("123asr45") '123
Debug.Print Val("2,22") '2 döndürür. çünkü virgülü non-numerik algılar ve ilk non-numerik karakterden önceki kısmı döndürür
Debug.Print Val("2.22") '2,22 döndüdür
'Debug.Print Val(2, 22) 'hata verir, sanki iki paramter var gibi alıgılar
Debug.Print Val(2.22) '2 döndürür, noktayı nonnumerik algılar. Burada 2,22 dönmesini beklediyseniz yukarıda yazıanları tekrara okuyun lütfen</pre>
	<p>Val'in tarihleri nasıl değiştirdiğini de görelim.<br></p>
	<pre class="brush:vb">Debug.Print Val("21.01.1979") '21,01 döndürür, çünkü ikinci noktayı saysal olmayan karakter olarak algılar</pre>
	<p> <span class="keywordler">CInt, CDbl v.s(İfade)</span>: Bunlar da 
	dönüştürme görevi görürler. <strong>C</strong>'den sonraki tipe dönüşüm yaparlar. Ör: C<span style="color: red">Int</span>, 
	içindeki değeri <span style="color: red">Integer'</span>e dönüştürken, C<span style="color: red">Dbl</span>, 
	<span style="color: red">Double</span>'a, C<span style="color: red">Lng Long</span>'a. İfade 
	olarak bir string olabileceği gibi daha küçük boyutlu bir numerik değer de 
	olabilir.</p>
	<p> Şunu merak etmiş olabilirsiniz. CDbl da Val de double türüne dönüştürüyor. 
	Ne fark var? Neden 2 tane fonksiyon var? Aşağıdaki kodları ve sonuçları ile 
	yukardaki Val'in sonuçlarını incelerseniz farkı görebilirsiniz.</p>
	<pre class="brush:vb">Debug.Print CDbl("00123") '123
'Debug.Print CDbl("123asr45") 'hata alır. numerik olmasını bekler
Debug.Print CDbl("2,22") '2,22 döndürür
Debug.Print CDbl("2.22") '222 döndürür
'Debug.Print CDbl(2, 22) 'hata verir, sanki iki paramter var giib alÄ±gÄ±lar
Debug.Print CDbl(2.22) '2,22 döndürür</pre>
	<p>Biraz karışık gelmiş olabilir. Özeti şu: Val'i sanki Str fonksiyonu ile 
	string hale gelmiş sayısal ifadeleri sayıya çevimek için kullanmanız 
	gerekirken, CDbl'i kullanıcı tarafından girilen bir sayısal metni sayıya 
	çevirmek için kullanın.</p>
	<h3>Yuvarlama Fonksiyonları</h3>
	<p> <span class="keywordler">Int(Number) ve Fix(Number)</span>: İkisi de 
	küsurlu sayıların küsuratını atar. 
	Fark şu: Int, negatif sayılarda aşağı doğru yuvarlarken, Fix sadece küsurat 
	atar. Ör: 3,85 için ikisi de 3 döndürürken, -3,85 için Int -4, Fix ise -3 
	döndürür.</p>
	<p> <span class="keywordler">Round(Number,Duyarlılık)</span>;Excel'in built-in(yerleşik) Round fonksiyonundan farklıdır. Excel fonksiyonunda negatif değer girerek 
	sayı 10'un katları şeklinde de yazılabilirken, VBA'de sadece pozitif rakamlar 
	girilebilmektedir. <strong>Round(95.458, 1)</strong> bize 95.5 değerini 
	verir. Ondalık olarak virgül yerine nokta karakteri yazıldığına dikkat edin.&nbsp; </p>
	<p> Negatif değer girip 10'un katları şeklinde yuvarlamak ve 
	RoundUp/RoundDown gibi seçenekleri ele almak için WorksheetFunction'dan 
	faydalanabilriz.</p>
	<pre class="brush:vb">Debug.Print Round(95.458, 1) '95,5
Debug.Print worksheetfunction.Round(95.498, -1) '100
Debug.Print worksheetfunction.RoundUp(95.498, 0) '96
Debug.Print worksheetfunction.RoundDown(95.498, 0) '95</pre>
	<h3> Matematsiksel fonksiyonlar</h3>
	<p> NOT:Trigonometrik fonksiyonlar gibi MIS dünyasında kullanımı az olan veya hiç olmayan 
	fonksiyonlara yer verilmemişir.</p>
	<p> <span class="keywordler">Randomize ve Rnd(Number)</span>:Rasgele sayı 
	üretmek için kullanılırlar. Rasgele sayı üretimi bilgisayarın sistem saati 
	baz alınarak üretilir. Rnd ile 0-1 arasında rasgele sayı üretilir. Sayı 
	üretildikten sonra, teknik olarak bir sonraki rasgele sayının ne olacağı 
	tahmin edilebilir, bu da sonraki sayının rasgeleliğine şüphe düşürür. Bununla birlikte her Rnd işleminden önce 
	<strong>Randomize</strong> 
	fonksiyonu tek başına kullanılırsa sistemin baz alacağı değer bir nevi 
	resetlendiği için bir sonraki rasgele sayının tahmini imkansızlaşır ve 
	gerçek bir rasgele sayı üretilmiş olur. Özellikle birden fazla rasgele sayı 
	üretmeniz gerektiği durumlarda Rnd öncesinde Randomize kullanmanız önerilir, 
	diğer durumlarda tek başında Rnd iş görecektir.</p>
	<p> Rnd, Excelin RAND fonksiyonu gibi işlemeketdir. Aşağıda örnekte 1 ile 100 
	arasında bir sayı üretilip x değişkenine atanmaktaır, bu da 
	RANDBETWEEN(1,100) gibi.</p>
	<pre class="brush:vb">Randomize
x=CInt(Rnd*100)

'ancak 10-20 gibi daha üst seviyelerde bir rasgele rakam istenirse
x=WorksheetFunction.RandBetween(10,20)</pre>
	<p> <span class="keywordler">Abs(number)</span>: Mutlak değer üretir.
	Ör.
	<strong>Abs(-100)</strong> 100 değerini verir.</p>
	<h3> Sayı formatlama</h3>
	<p> Sayıları formatlamak için iki fonksiyon bulunmakta.
	<span class="keywordler">Format </span>ve <span class="keywordler">
	FormatNumber</span>. Bunların ikisi de aslında String 
	modülünün bir fonksiyonudur. Ama konsept olarak sayılarla ilgili olduğu 
	için burada ele almayı uygun buldum.&nbsp;</p>
	<p> Ben bunlardan FormatNumber yerine Format'ı kullanmayı tercih ediyorum. 
	Zaten FormatNumber üzerine de çok fazla online bilgi de bulunmuyor. Format 
	fonksiyonuyla ilgili detay bilgilere ise
	<a href="https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/format-function-visual-basic-for-applications">
	buradan</a> ulaşabilrisiniz.</p>
	<p> Sayılarda formatlama yaparken dikkat edilecek hususlar bellidir. Binlik 
	ayraç, ondalık ayracı, yuvarlamalar ve gerekirse belirli miktarda fazladan 0 
	gösterimi.</p>
	<p> Formatlamada iki temel karakter kullanıyoruz. <strong>0 </strong>ve<strong> 
	#</strong>. Ben bunlardan en çok #'i kullanıyorum. 0'ın farkı şu:Eğer 
	sayının başında belirli miktar 0 olsun istenirse bu kullanılır. Neden böyle 
	birşey gereksin? Mesela BT ekibine bir liste hazırlıyorsunuzdur, listeyi sizden 10 
	haneli sayılar şeklinde(başta 0 olacak şekilde) isterler. Mesela sizin 
	göndereceğiniz listedeki sayılardan biri 123 ise bunu 0000000123 şeklinde 
	isterler. 0, aynı zamanda küsuralartarda da fazladan 0 gösterebilir. Mesela 
	4 haneli küsurat olsun isteniyorsa 3,18'in gösterimi 3,1800 şeklinde olur.</p>
	<p> Gerek 0 gerek #, eğer gerekenden az miktarda kullanılmışsa 
	otomatikman gereken adede tamamlanır. (Aşağıda hem 0'ın hem #'in ilk 
	örneklerinde görüldüğü gibi). Ondalık ayraçın solunda sadece 1 adet 0/# 
	olmasına rağmen dört rakamın dördü de gösterilmiştir.</p>
	<p> Aşağıdaki örneklerle anlattıklarımız pekiştirelim. Başlamadan önce "ön 
	bilgi" bölümünde yazan '.' ve ',' işaretlerinin kullanımına tekrar bakmanızı 
	öneririm.</p>
	<pre class="brush:vb">
Sub formatting()
    Debug.Print Format(8315.4, "0.000") '8315,400
    Debug.Print Format(8315.4, "0,000") '8.315
    Debug.Print Format(8315.4, "0.0") '8315,4
    Debug.Print Format(8315.4, "0,0") '8.315
    Debug.Print Format(8315.4, "0000.0") '8315,4
    Debug.Print Format(8315.4, "0000000.0") '0008315,4
     
    Debug.Print vbNewLine
    Debug.Print Format(8315.4, "#.###") '8315,4. 0 formatından farklı olarak takip eden 0lar görünmez.
    Debug.Print Format(8315.4, "#,###") '8.315
    Debug.Print Format(8315.4, "#.#") '8315,4
    Debug.Print Format(8315.4, "#,#") '8.315
    Debug.Print Format(8315.4, "####.#") '8315,4
    Debug.Print Format(8315.4, "#######.#") '8315,4. 0 formatından farklı olarak baştaki fazla 0lar görünmez.
     
    'nokta ve virgül beraber
    Debug.Print vbNewLine
    Debug.Print Format(8315.4, "0,000.00") '8.315,40
    Debug.Print Format(8315.4, "#,###.##") '8.315,4
    Debug.Print Format(8125648315.486, "#,###.##") '8.125.648.315,49
End Sub
	</pre>

</div>
	</a>
</asp:Content>
