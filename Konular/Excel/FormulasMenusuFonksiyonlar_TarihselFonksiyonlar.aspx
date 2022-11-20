<%@ Page Title='FormulasMenusu1 TarihselFormuller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Tarihsel Formüller</h1>

<p>Tarihsel formüller, Excelde sık kullandığımız formüllerdendir. Bunları tek tek kullandığımız gibi başka fonksiyonlarla birarada da kullanımı oldukça yaygındır.</p>

<p>Öncelikle giriş mahiyetinde bir ön bilgi vermek isterim. Excel, tarihleri nümerik sayılar olarak tutar. Bunların tarih olarak gösterimi ise tamamen bir formatlama işidir. Mesela, bir hücrede 21.01.2017 tarihi varsa bu aslında 42756 sayısından başka bir şey değildir. Ancak Excel, girilen değerin tarihsel bir değer olup olmadığını anladığı için onu otomatik olarak formatlar ve 21.01.2017 olarak gösterir. Format Cell yapıp "General" tipini seçtiğinizde 42756'ya döndüğünü görürsünüz.</p>

<p>Şimdi genel olarak tarihsel formüllere bakacağız, sonrasında da çeşitli örnekler üzerinden sık ihtiyaç duyulan görevlerin nasıl yapıldığına bakacağız.</p>


<!--********************************************************************************************************************************************--><h2 class="baslik">Genel Bakış</h2>
<div class="konu">

<h3>Formüller</h3>


<p><span class=" keywordler">NOW():</span>Bugünün tarihini ve saat ve dakikasını verir. İstenirse Format Cell yapılarak saniye de gösterilebilir.</p>
<pre class="formul">=HOUR(NOW()) //Şuanın saatini verir. MINUTE ile de dakikası elde edilir.</pre> 

<p><span class=" keywordler">TODAY():</span>Bugünün tarihini verir. Format Cell ile saat gösterilmeye çalışılırsa 00:00:00 gösterilir.</p>
	<pre class="formul">=IF(A2&lt;TODAY();"Eski";"Yeni")</pre>
<p>Dünün  tarihini hesaplamak da zor değildir.</p>
	<pre class="formul">=TODAY()-1</pre>
<p>Belirli bir tarihe kaç gün kaldığını hesaplamak da aynı derecede basittir.</p>
	<pre class="formul">
=A2-TODAY()
//veya
=DAYS(A2;TODAY())
</pre>

<p><span class=" keywordler">DAY(Tarih):</span>İlgili tarihin ayın kaçıncı günü olduğunu verir.</p>
	<pre class="formul">=IF(DAY(TODAY())<=15;"Ayın ilk yarısı";"Ayın ikinci yarısı")</pre>

<p><span class=" keywordler">MONTH(Tarih):</span>İlgili tarihin kaçıncı ayda olduğunu verir.</p>
	<pre class="formul">=MONTH(TODAY())</pre>

<p><span class=" keywordler">YEAR(Tarih):</span>İlgili tarihin yılını verir. Aşağıdaki örnekte bugünün yılından "doğumgünü" isimli Named Range'in yılı çıkarılarak kişinin yaşı bulunmaktadır. </p>
<p><pre class="formul">=YEAR(TODAY())-YEAR(dogumgünü)</pre></p>
<p>Bu işlem aşağıdaki formülle de yapılabilir. Tek farkı, bunun küsuratlı bir değer vermesi ve gerektiğinde aşağı/yukarı yuvarlamaya imkan vermesidir. Mesela 38,9 yaşında çıkan birini 39 göstermek daha doğrudur.</p>
<p><pre class="formul">=YEARFRAC(dogumgünü;TODAY())</pre></p>


<p><span class=" keywordler">WEEKDAY(Tarih;[Başlangıç]):</span>İlgili tarihin haftanın kaçıncı günü olduğunu verir. İkinci parametre, haftanın ilk gününün ne alınması gerektiğini belirtir. Bilindiği gibi bazı ülkelerde haftabaşı Pazar iken bazılarında Pazartesidir. Başka seçenekler de var tabi ancak genelde 1 ve 2 kullanılacaktır. Default değer 1'dir, yani seçim yapmazsanız haftabaşı Pazar gibi alınır. Bizim ülkemiz için bu biraz kafa karıştırıcı olabilir.O yüzden bu fonksiyonu 2 ile kullanmanızı tavsiye ederim.</p>

<p>Mesela, bi kolondaki tarihlerden haftasonlarını işaretlemek veya filtrelemek istiyorsunuz, veya haftasonuysa şu, değilse bu şeklinde bir formül yazakcasınız </p>
<p><pre class="formul">=IF(WEEKDAY(A2;2)>5;"Haftaiçi";"Haftasonu")</pre></p>

<p><span class=" keywordler">WORKDAYS(Başlangıç; kaçgün; [Tatiller]):</span>Bir tarihe belirtilen adette işgünü ekler. İsterseniz bir hücre grubuna gireceğiniz tatil günleri ile bunları da eklenecek günlere dahil edebilirsiniz.</p>
<p><pre class="formul">=WORKDAY(A1;7;B1:B3)</pre></p>

<p><span class=" keywordler">NETWORKDAYS(Başlangıç; bitiş; [Tatiller]):</span>İki tarih arasındaki <u>iş günü</u> sayısını verir. İsterseniz bir hücre grubuna gireceğiniz tatil günleri ile bunları da hariç tutabilirsiniz. </p>
<p><pre class="formul">=NETWORKDAYS(A1;A2;B1:B3) //B1:B3 arasında bayram tatilleri girilmiş</pre></p>

<p>İki gün arasındaki toplam gün sayısı için ilgili tarihler birbirinden doğrudan çıkarılır ve 1 eklenir.</p>
<p><pre class="formul">=A1-A2+1</pre></p>

<p>Her iki yöntemde de başlangıç ve bitiş tarihleri günsayısına dahildir.</p>

<p><span class=" keywordler">DATEVALUE(StringTarih):</span>Metin formatında verilen tarihi Tarih formata çevirir. Böylece bu tarihi başka tarihlerle karşılaştırabilirsiniz. Ayrıca Özet tablo veya Grafik yapmak istediğinizde tarihler sıralı bir şekilde gelir ve tarih olarak kullanılır, aksi halde alfabetik sıraya göre gelir ve tarih özelliklerinden faydalanılamaz</p>

<img src="/images/metinseldatevalue.jpg">&nbsp;
	<p><span class=" keywordler">EDATE(Tarih;aysayısı):</span>İlgili tarihe belirtilen ay kadar ekleme yapar. Ör:25 Marta 2 ay eklenirse 25 Mayıs olur. Özel durum olarak 28 şubatı söyleyelim. 31 Ocak'a 1 ay eklenirse ilgili yılın artık yıl içerip içermediğien göre 28 veya 29 Şubat döndürür. Ancak 28 Şubata 1 ay eklendiğinde 31 Mart değil 28 Mart döndürür. Bu fonksyion bu bağlamda, Oracle SQL'deki Add_months ve SQL Server'daki DateAdd fonksiyonlarına  benzemektedir.</p>
	<pre class="formul">=EDATE(A1;2)</pre>

<p><span class=" keywordler">EOMONTH(Tarih;aysayısı):</span>İlgili tarihe beliritlen adet kadar ay eklendiğinde çıkan tarihin aysonunu verir.</p>
	<pre class="formul">=EOMONTH(A1;2) //2 ay sornasının ay sonunu verir</pre>

<p><span class=" keywordler">DATE(Yıl,Ay,Gün):</span>Verilen Yıl, Ay ve Gün birleştirilerek ilgili tarih elde edilir..</p>
<p><pre class="formul">
=DATE(2017;1;21) //21.01.2017
</pre></p>


<h3>Formatlama işlemleri</h3>
<p>Bazen tarih formatı olan bir hücreyi bir metin formülü içinde kullanmak isteriz. Böyle bir durumda Excel otomatikman bu tarihin temel numerik değerini kullanır, ki bu da pek hoş olmaz. Ne demek istediğimi aşağıdaki örnekte görebilirsiniz</p>

<img src="/images/excelmetinseltext1.jpg"  alt="kısatanım"  />

<p>Bu sorunu çözmek için bu tarihleri formatlamamız gerekir. Bu konuyu <a href="/Konular/Excel/FormulasMenusuFonksiyonlar_MetinselFonksiyonlar.aspx#formattext">bu sayfadaki</a> TEXT formülüyle yapıyoruz.</p>

</div>
<!--********************************************************************************************************************************************--><h2 class="baslik">Çeşitli Örnekler</h2>
<div class="konu">

<h4 class="baslik">Çeşitli günleri tespit etme</h4>
<div>
<p>Sık kullanılan özel tarihleri aşağıdaki gibi özetleyebiliriz.</p>
<p><pre class="formul">
=EOMONTH(TODAY();0) //bu ayın aysonu
=EOMONTH(TODAY();-1)+1 //bu ayın başı
=TODAY()-DAY(TODAY())+1 //bu ayın başı(2)
=EOMONTH(TODAY();0)+1 //sonraki ayın başı
=EOMONTH(TODAY();1) //sonraki ayın sonu
=EOMONTH(TODAY();-2)+1 //geçen ayın başı
=EOMONTH(TODAY();-1) //geçen ayın sonu	
=TODAY()-WEEKDAY(TODAY();2)+1 //bir önceki Pazartesi: (Bugünden, bugünün gün numarasını çıkarıp 1 ekliyoruz)
=TODAY()+(8-WEEKDAY(TODAY();2)) //bir sonraki Pazartesi: (Bugüne, "7-bugünün gün numarası" farkını ekleyip 1 daha ekliyoruz) 
=WORKDAY(EOMONTH(date)+1,-1) //bu ayın son işgünü
</pre></p>
</div>


<h4 class="baslik">Çeşitli seviyelerde artışlar yapma</h4>
<div>
<p>Çoğu kez belirli bir tarihe belirli frekanslarda eklemeler yapmamız gerekecek, aşağıda bunları örnek olarak göstermek istedim.</p>
<img src="/images/exceltarihsel0.jpg">
<p>Yıl eklemek için ayrıca aşağıdaki fomrül de yazılabilir. 5 yıl ekleyelim.</p>

<pre class="formul">=DATE(YEAR(A1)+5;MONTH(A1);DAY(A1))</pre>
</div>


<h4 class="baslik">Tatil günlerini sayma</h4>
<div>
<p>TatilRange isminde bir Named Range'imiz olduğunu düşünelim. Buraya belirli bir dönemdeki tüm tarihleri girmiş olalım. 
A1 ve A2 de sırayla başlangıç ve bitiş tarihleri bulunuyor olsun. Böyle bir durumda bu iki tarih arasındaki tatil günlerini aşağıdaki gibi hesaplarız.</p>
<pre class="formul">=SUMPRODUCT((TatilRange>=A1)*(TatilRange<=A2))</pre>
<p>Eğer buraya bir de haftasonlarını eklemek isterseniz formülümüzü şu şekilde güncelleyebiliriz.</p>
<pre class="formul">=SUMPRODUCT((TatilRange>=A1)*(TatilRange<=A2))+A2-A1+1-NETWORKDAYS(A1;A2)</pre>
</div>



</div>

</asp:Content>
