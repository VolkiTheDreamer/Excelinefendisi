<%@ Page Title='FormulasMenusu1 BilgiVerenFormuller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>

<h1>Bilgi Veren Formüller</h1>
<p> Excelin <strong>Information</strong> kategorisinde yer alan fonksiyonların 
hepsine bu sayfada yer verilmemiştir. Her zaman olduğu gibi MIS/Raporlama 
konularında bize yardımı olabilecek fonksiyonları ineleyeceğiz.</p>
	<p> Bu fonksiyonlar <strong>TRUE/FALSE</strong> değerlerini döndürüp çoğu 
	zaman <a href="FormulasMenusuFonksiyonlar_MantiksalFonksiyonlar.aspx">IF</a> 
	fonksiyonu ile birlikte kullanılırlar. Bazen de IF olmadan kullanılırlar. 
	Mesela formül yazıldıktan sonra tüm kolonda aşağı indirilir ve TRUE/FALSE 
	filtrelemesi yapılabilir.</p>
	<p> Şimdi önemli olanlarına bir bakalım.</p>
	<p> <span class="keywordler">ISBLANK</span>: Bir hücrenin içeriğinin boş 
	olup olmadığını döndürür. Ancak hücre içeriğine özellikle bir formül sonucu 
	"" değeri geliyorsa(Ör:IF(A2&gt;0;1;"") bu fonksiyonla sorgulandığında TRUE 
	döndürmez. Bu fonksiyonun TRUE döndürmesi için hücrenin içeriğinin tam 
	anlamıyla boş olması gerekir. </p>
	<p> Aşağıdaki örnekte A2 hücresi boşsa Evet, değilse Hayır yazdıran bir 
	formül var. Muhtemelen bu formül kolon boyunca aşağı indirilerek 
	kullanılacaktır.</p>
	<pre class="formul">=IF(ISBLANK(A2);"Evet";"Hayır")</pre>
	<p> <span class="keywordler">ISERR, ISERROR ve ISNA</span>:Bunlardan 
	ISERROR, ISERR ve ISNA'nın bileşimidir. ISERR, NA dışında bir hata olup 
	olmadığını sorgularken ISNA sadece NA'ları sorgular. (NA, bildiğiniz gibi 
	aranda bir değerin bulunamaması durumunda döner.) ISERROR ise tüm hataları 
	yakalar. Ancak açık söylemek gerekirse 2007 ile gelen <strong>
	<a href="FormulasMenusuFonksiyonlar_MantiksalFonksiyonlar.aspx">IFERROR</a></strong> 
	fonksiyonundan sonra bunlara çok gerek kalmamıştır. Belki bunları, IF'siz 
	haliyle çalıştırıp, kolonda aşağı kaydırıp filtreleme yapmak istediğinizde 
	kullanılabilirsiniz.</p>
	<pre class="formul">=ISNA(A2) //Ör:Vlookup sonucunda eşleşme olmayan(NA döndüren) kayıtları filtrelemek istiyoruz
=ISERR(A2) //Ör:0'a bölmeleri(DIV/0) hatası döndüren) kayıtları filtrelemek istiyoruz
=ISERROR(A2) //Her tür hata</pre>
	<p> <span class="keywordler">ISEVEN ve ISODD</span>:Bir sayının tek mi çift 
	mi olduğunu sorgulamak istediğimizde bunları kullanırız. Mesela bir şubedeki 
	müşterileri hacim sırasına dizip şubedeki iki müşteri temsilcisine 
	atayacaksınız diyelim. Aşağıdaki fonksiyon ile kolayca çözüme 
	ulaşabilirsiniz.</p>
	<pre class="formul">=IF(ISODD(C2);"Ali";"Veli")</pre>
	<p> <img src="/images/excelformulinfoisodd.jpg"></p>
	<p> <span class="keywordler">ISFORMULA</span>: Her ne kadar Home menüsünden
	<strong><a href="Giris_PratikKisayollar.aspx#ozelbul">Find&amp;Select</a></strong> diyip 
	ilgili sayfa veya seçili yerdeki formülleri tek seferde seçebilyorsanız da 
	bazen daha sistematik bir "hücre formül içeriyor mu" kontrolüne ihtiyacınız 
	olabilir. Mesela formül içeren hücreleri filtrelemek isteyebilrisiniz.</p>
	<p> Diyelim ki <strong>şube kodu-segment-ürün adı-dönem-hedef</strong> 
	şeklinde 100bin satırlık bir datanız var. Bir yan kolona bir formül yazdınız 
	ve aşağı çektiniz, ama 100bin satır olduğu için bilgisayarınız kastı, siz de 
	tüm formülleri Value olarak yapıştırdınız ama tek tük de olsa bazı 
	satırlardaki hedefleri başka formül yazarak revize etme ihtiyacınız oldu. 
	Sonra da başka bir gruba farklı bir formül yazmanız gerekti ve bu böyle bi 
	süre daha gitti. En son durumda hangi satırlarda formül olduğunu görmek 
	isterseniz bu formülü kullanabilirsiniz.</p>
	<pre class="formul">=ISFORMULA(H2) //ve aşağı kaydırılır</pre>
	<p> <span class="keywordler">ISNONTEXT ve ISNUMBER</span>:Sırayla hücrenin 
	içeriğinin metin dışı bir değer içerip içermediğini ve sayı içerip 
	içermediğini döndürürler. ISNONTEXT, sadece metinler için FALSE döndürürken 
	boşluklar dahil herşey için TRUE döndürür. ISNUMBER ise sayı ve 
	tarihler(tarihler Excelde sayı olarak tutulur) için TRUE döndürürken diğer 
	değerler için FALSE döndürür.</p>
	<p> <span class="keywordler">N</span>:Bir değeri sayısal hale çevrir.
	<strong>VALUE</strong> fonksiyonuna benzemekle birlikte VALUE'nun sayıya 
	çeviremediği birçok değeri de sayıya&nbsp; çevirebilir(1/0 şeklinde olsa 
	da). Bu fonksiyonu
	<a href="FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx">dizi 
	formüllerinde ve SUMPRODUCT</a> içinde, değerleri sayısal hale çevirmenin 
	alternatif bir yolu olarak oldukça kullanacğaız. </p>
	<p> Aşağıda bu son 4 fonksiyonun karşılaştırmalı bir tablosu görülmektedir.</p>
	<p> <img src="/images/excelformulainfo2.jpg"></p>
	<p> &nbsp;</p>

</asp:Content>
