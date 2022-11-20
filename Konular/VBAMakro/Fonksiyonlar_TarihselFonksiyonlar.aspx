<%@ Page Title='Fonksiyonlar TarihselFonksiyonlar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik">
	<div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>Tarihsel Fonksiyonlar</h1>

	<h2 class="baslik">Ön Bilgiler</h2>
<div class="konu">
	<p>Excel’de tüm tarihlerin sayısal bir karşılığı vardır. Aslında karşılığı 
	vardır demekten ziyade tüm tarihler bir sayıdır demek daha doğru olur. Excel bunları gösterirken 
	tarih formatında gösterir. Tarihler tam sayı depolanırken saatler küsurlu 
	formda tutulur. Bu sayı 1.1.1900'den itibaren geçen günsayısıdır. Ör:17.05.2015 
	arka planda 42141 sayısı olarak tutulur.</p>
	<p>Ayraçları, bölgesel ayarlar belirler, ancak tarih “<strong>#</strong>” 
	işaretleri arasında yazılırsa bölgesel ayarlardan bağımsız olarak Amerikan 
	tarih formatına göre belirtilmiş olur:<strong>Ay/gün/yıl</strong>. Bunu pek 
	kullanmanızı önermem, ancak görürseniz de şaşırmayın. Biz bildiğimiz 
	formatta kullanalım, yoksa 5 Temmuz yazacağınız yerde 7 
	Mayıs gibi sonuçlar 
	elde edebilirsiniz, bu da önemli hatalara neden olabilir.</p>
	<p>Tarihlerle ilgili dikkat edilmesi gereken bir husus var. 
	Formatlı Gösterim tipi, kaynak tipi veya dönüş tipi. Formatlı Gösterim şekli 
	her zaman stringtir. Ancak kaynak tip veya dönüş tipi farklı tipler de 
	olabilir.</p>
	<p>Tarihlerin kaynak/dönüş tipi aşağıdakiler olabilir.</p>
	<ul>
		<li>String</li>
		<li>Nümerik</li>
		<li>Tarihsel<ul>
			<li># karakterleri arasında</li>
			<li>Tarihsel bir 
	fonksiyondan dönen değer olarak. (Excel hücresinden veya Date, DateSerial gibi VBA 
	fonksiyonlarından)</li>
		</ul>
		</li>
	</ul>
	<p>Konuyla ilgili daha detaylı bilgiye
	<a href="https://msdn.microsoft.com/en-us/library/aa227472(v=vs.60).aspx">bu 
	sayfadan</a> ulaşabilirsiniz.</p>
	</div>
	<h2 class="baslik">DateTime modülü</h2>
<div class="konu">
	<p>Tarihsel işlemlerimizi yaparken DateTime modülünü kullanıyor olacağız. Bu 
	modüldeki fonksiyonlara geçmeden önce sık kullanılan property'lere bir 
	bakalım.</p>
	<h3>Propertyler</h3>
	<p> <span class="keywordler">Date ve Date$</span>: O anki günü verirler. İlki Variant döndürürken ikincisi 
	String döndürür. Aralarındaki fark, Variant olanda null işlemi yapabilmenizdir. 
	Bir diğer fark da Date ile bir sayıyı matematiksel işleme tabi 
	tutabilirsiniz ancak Date$ ile bunu yapamazsınız .</p>
	<pre class="brush:vb">Debug.Print Date+1 'yarının tarihini verir
Debug.Print Date$+1 'hata verir. String olduğu için.</pre>
	<p> Aşağıdaki kod, üzerinde çalıştığımız dosyanın son değişikliğin tarihinin 
	bugünden küçük olması durumunda çalışır.</p>
	<pre class="brush:vb">'fso tanımının yapıldığı kodlar
If fso.DateModified &lt; Date Then
   'diğer kodlar
End if</pre>
	<p> <span class="keywordler">Time ve Time$</span>: Şu anın saatini verir. 
	aralarındaki fark Date/Date$ arsındaki farkların aynısıdır.</p>
	<p> <span class="keywordler">Now</span>: Şu anın tarih ve saatini verir. Date ve Time'ın birleşimi gibi 
	düşünülebilir.</p>
	<h3>Değişken tanımlamalar</h3>
	<p>DateTime modülünün içindeki <strong>Date</strong> tipi ile hem tarih hem saat tanımlanabilir. 
	Aşağıdaki örneklerde tüm değişkenli Variant tanımlıyoruz.</p>
	<pre class="brush:vb">
Dim tarih1 'As Date
Dim tarih2 'As Date&lt;
Dim stringtarih1 'As String
Dim stringtarih2 'As String
Dim numeriktarih1 'As Long
Dim numeriktarih2 'As Double
Dim diyezlitarih1 'As Date
Dim diyezlitarih2 'As Date
Dim hucredentarih1 'As Date
Dim hucredentarih2 'As Date
	</pre>
	<p>Örnek kodlara geçmeden önce, Range nesnesinin <strong>Value</strong> ve <strong>Value2</strong> 
	özelliklerinin farkını 
	tekrar edelim. Value, hücrenin içeriğini alırken, Value2 tarihleri numerik 
	değere çevirerek depolar. (Tabi siz bir değişkeni Date olarak tanımladıysanız, ona 
	Value2 sonucunu aktarsanız bile sayısal değer değil yine tarihsel formattaki 
	değeri depolanır.)</p>
	<h3>Dönüş tipleri</h3>
	<pre class="brush:vb">tarih1 = Now
tarih2 = DateSerial(1979, 1, 21)
stringtarih1 = "21.01.1979"
stringtarih2 = "01.01.1979"
numeriktarih1 = 42574
numeriktarih2 = 42574.127
diyezlitarih1 = #1/21/1979# 'amerikan formatında olmak zorunda, önce ay ve ayraç oalrak da "/" girili "." dğeil. gün girerken 01 
girsen bile enterea basıp alta geçince kaybolurlar.
diyezlitarih2 = #1/21/1979 4:30:00 AM# 'amerikan 
formatında olmak zorunda, önce ay ve ayraç olarak da /
hucredentarih1 = range("a1").Value 'hücre içeriği:24.07.17 13:14
hucredentarih2 = range("a1").Value2 'hücre içeriği:21.01.1979</pre>
	<p>Şimdi bunların çeşitli özelliklerini yazdıralım.</p>
	<pre class="brush:vb">
Debug.Print "tarih1", vbTab, tarih1, TypeName(tarih1),VarType(tarih1), IsDate(tarih1)
Debug.Print "tarih2", vbTab, tarih2, vbTab, TypeName(tarih2), VarType(tarih2), IsDate(tarih2)
Debug.Print "stringtarih1", vbTab, stringtarih1, vbTab, TypeName(stringtarih1), VarType(stringtarih1), IsDate(stringtarih1)
Debug.Print "stringtarih2", vbTab, stringtarih2, vbTab, TypeName(stringtarih2), VarType(stringtarih2), IsDate(stringtarih2)
Debug.Print "numeriktarih1", vbTab, numeriktarih1, vbTab,TypeName(numeriktarih1), VarType(numeriktarih1), IsDate(numeriktarih1)
Debug.Print "numeriktarih2", vbTab, numeriktarih2, vbTab,TypeName(numeriktarih2), VarType(numeriktarih2), IsDate(numeriktarih2)
Debug.Print "diyezlitarih1", vbTab, diyezlitarih1, vbTab,TypeName(diyezlitarih1), VarType(diyezlitarih1), IsDate(diyezlitarih1)
Debug.Print "diyezlitarih2", vbTab, diyezlitarih2, TypeName(diyezlitarih2), VarType(diyezlitarih2), IsDate(diyezlitarih2)
Debug.Print "hucredentarih1", hucredentarih1, vbTab,TypeName(hucredentarih1), VarType(hucredentarih1), IsDate(hucredentarih1)
Debug.Print "hucredentarih2", hucredentarih2, vbTab,TypeName(hucredentarih2), VarType(hucredentarih2), IsDate(hucredentarih2)</pre>
	<p>Çıktısı aşağıdaki gibi olacaktır.</p>
	<p><img src="/images/vbafunctarihsel1.jpg"></p>
	</div>
	<h2 class="baslik">Dönüşüm İşlemleri</h2>
	<div class="konu">
	<p>Bazen, gösterim tipi metinsel olan tarihleri tarih tipine çevirip onlar 
	üzerinden işlemlerinize devam etmek istersiniz. Bunun için VBA bize 4 yöntem 
	sağlmaktadır. Dördü de küçük nüanslar&nbsp;göstermektedir.</p>
	<p><span class="keywordler">1.Yöntem:CDate(Exp)</span>:CDate, bir ifadeyi tarih yaparken tipini de tarih yapar. 
	Üstelik saat v.s bilgisi varsa bunları da korur. İçine her tür ifadeyi 
	alabilir.</p>
	<pre class="brush:vb">
x1 = CDate(tarih1)
x2 = CDate(tarih2)
x3 = CDate(stringtarih1)
x4 = CDate(stringtarih2)
x5 = CDate(numeriktarih1)
x6 = CDate(numeriktarih2)
x7 = CDate(diyezlitarih1)
x8 = CDate(diyezlitarih2)
x9 = CDate(hucredentarih1)
x10 = CDate(hucredentarih2)	</pre>
	<p>Şimdi bunları yazdıralım.</p>
	<pre class="brush:vb">Debug.Print x1, TypeName(x1), VarType(x1)
Debug.Print x2, TypeName(x2), VarType(x2)
Debug.Print x3, TypeName(x3), VarType(x3)
Debug.Print x4, TypeName(x4), VarType(x4)
Debug.Print x5, TypeName(x5), VarType(x5)
Debug.Print x6, TypeName(x6), VarType(x6)
Debug.Print x7, TypeName(x7), VarType(x7)
Debug.Print x8, TypeName(x8), VarType(x8)
Debug.Print x9, TypeName(x9), VarType(x9)
Debug.Print x10, TypeName(x10), VarType(x10)</pre>
	<p>Çıktısı şöyle olacaktır:</p>
	<p><img src="/images/vbafunctarihsel2.jpg"></p>
	<p><span class="keywordler">2.Yöntem:DateValue(Metin)</span>: Bu fonksiyon ile 
	metin formatındaki tarihler gerçek tarihe çevrilir. Saat v.s detayını 
	korumaz. Argüman olarak sadece metin alır.</p>
	<pre class="brush:vb: Bu fonksiyon ile 
	metin formatındaki tarihler numerik değere çevrilir. Sonuç bir tam sayıdır. 
	Saat v.s bilgisni kourmaz, sadece tairhi döndürür. Saat v.s bilgsi için de 
	TimaValue kullanlabilir.</p>
	<pre class="brush:vb">DateValue("21.01.1979") '21 Ocak 1979 tarihini verir
DateValue(28876) 'hata verir. Çünkü argüman olarak sadece metin alır.</pre>
	<p><span class="keywordler">3.Yöntem:Format(exp, </span>
	format): sadece gösterim 
	şeklini tarihsel yapar, 
	dönüş tipi Stringtir. Dönen değerle tarihsel işlem yapılamaz.</p>
	<p>İkinci parametre olarak 
	önceden tanımlanmış formatlar da girilebilir, 
	kullanıcı tanımlı formatlar da. Bunların 
	detayın
	<a href="https://msdn.microsoft.com/VBA/Language-Reference-VBA/articles/format-function-visual-basic-for-applications">
	buradan</a> ulaşabilriiniz.</p>
	<pre class="brush:vb">y1 = Format(tarih1, "Short Date") 'önceden tanımlanmış parametre
y2 = Format(tarih2, "dd.mm.yyyy") 'bu ve aşağıdakiler ise kullanıcı tanımılı paremtredir
y3 = Format(stringtarih1, "dd.mm.yyyy")
y4 = Format(stringtarih2, "dd.mm.yyyy")
y5 = Format(numeriktarih1, "dd.mm.yyyy")
y6 = Format(numeriktarih2, "dd.mm.yyyy")
y7 = Format(diyezlitarih1, "dd.mm.yyyy")
y8 = Format(diyezlitarih2, "dd.mm.yyyy")
y9 = Format(hucredentarih1, "dd.mm.yyyy")
y10 = Format(hucredentarih2, "dd.mm.yyyy")</pre>
	<p>Şimdi de bunların çıktısını alalım</p>
	<pre class="brush:vb">Debug.Print y1, TypeName(y1), VarType(y1)
Debug.Print y2, TypeName(y2), VarType(y2)
Debug.Print y3, TypeName(y3), VarType(y3)
Debug.Print y4, TypeName(y4), VarType(y4)
Debug.Print y5, TypeName(y5), VarType(y5)
Debug.Print y6, TypeName(y6), VarType(y6)
Debug.Print y7, TypeName(y7), VarType(y7)
Debug.Print y8, TypeName(y8), VarType(y8)
Debug.Print y9, TypeName(y9), VarType(y9)
Debug.Print y10, TypeName(y10), VarType(y10)</pre>
	<p>Çıktı sonucu aşağıdaki gibidir:</p>
	<pre><img src="/images/vbafunctarihsel3.jpg"></pre>
	<p><span class="keywordler">4.Yöntem:FormatDateTime(d,constant))</span>:FormatDateTime da string döndürür, ama kullanımı daha 
	basittir, seçenekler sınırlıdır. 
	Seçeneklerde constantlar var, ön tanımlı  veya 
	kullanıcı tanımlı parametre  yok.</p>
	<pre class="brush:vb">z1 = FormatDateTime(tarih1, vbShortDate)
z2 = FormatDateTime(tarih2, vbLongDate)
z3 = FormatDateTime(stringtarih1, vbLongDate)
z4 = FormatDateTime(stringtarih2, vbLongDate)
z5 = FormatDateTime(numeriktarih1, vbLongDate)
z6 = FormatDateTime(numeriktarih2, vbLongDate)
z7 = FormatDateTime(diyezlitarih1, vbLongDate)
z8 = FormatDateTime(diyezlitarih2, vbLongDate)
z9 = FormatDateTime(hucredentarih1, vbLongDate)
z10 = FormatDateTime(hucredentarih2, vbLongDate)</pre>
	<p>Şimdi de çıktı alalım.</p>
	<pre class="brush:vb">Debug.Print z1, vbTab, TypeName(z1), VarType(z1), IsDate(z1) 'short olduğu ve içinde gün ismi geçmediği için True,  aşağıdakiler false false false
Debug.Print z2, TypeName(z2), VarType(z2), IsDate(z2)
Debug.Print z3, TypeName(z3), VarType(z3), IsDate(z3)
Debug.Print z4, TypeName(z4), VarType(z4), IsDate(z4)
Debug.Print z5, TypeName(z5), VarType(z5), IsDate(z5)
Debug.Print z6, TypeName(z6), VarType(z6), IsDate(z6)
Debug.Print z7, TypeName(z7), VarType(z7), IsDate(z7)
Debug.Print z8, TypeName(z8), VarType(z8), IsDate(z8)
Debug.Print z9, TypeName(z9), VarType(z9), IsDate(z9)
Debug.Print z10, TypeName(z10), VarType(z10), IsDate(z10)</pre>
	<p>Sonuç aşağıdaki gibi olacaktır:</p>
	<p><img src="/images/vbafunctarihsel5.jpg"></p>
	</div>
	<h2 class="baslik">Toplama Çıkarma</h2>
	<div class="konu">
	<p>VBA'de, Exceldeki yaptığımız gibi +1/-1 diyerek toplama 
	çıkarma yapılamıyor. Matematiksel işlemler için belirli fonksiyonlar 
	kullanılmalıdır.</p>
	<ul>
		<li>DateDiff/DateAdd</li>
		<li>DateValue("21.01.1979")+1 </li>
		<li>Dateserial(1979,1,21)+1 </li>
		<li>CDate("21.01.1979")+1 
	</li>
		<li>#"li üzerinde işlem: #5/22/97# - #1/10/97#
</li>
		<li>Date döndüren herhangi bir fonksiyonla. 
	(Date/Now veya UDF gibi)</li>
	</ul>
<pre class="brush:vb">
c1 = diyezlitarih1 + 1 'datevalue veya DateAdd demeye gerek kalmadan doğrudan kullanılabilir
c2 = DateAdd("d", 1, stringtarih1)
c3 = stringtarih1 + 1 'istenmeyen sonuç: noktaları uçup ekleme yapar, yani yıla eklenmiş olur, 21011980
c4 = DateValue(stringtarih1) + 1
c5 = CDate(stringtarih1) + 1
c6 = tarih1 - tarih2
c7 = DateDiff("d", tarih2, tarih1)
c8 = (tarih1 - tarih2) / 30 'yaklaşık ay sayısı
c9 = DateDiff("m", tarih2, tarih1) 'kesin ay sayısı
c10 = (tarih1 - tarih2) / 365 
</pre>

<p>Şimdi de çıktılarını alalım.</p>
<pre class="brush:vb">
Debug.Print vbNewLine
Debug.Print c1, vbTab, TypeName(c1), IsDate(c1)
Debug.Print c2, vbTab, TypeName(c2), IsDate(c2)
Debug.Print c3, vbTab, TypeName(c3), IsDate(c3)
Debug.Print c4, vbTab, TypeName(c4), IsDate(c4)
Debug.Print c5, vbTab, TypeName(c5), IsDate(c5)
Debug.Print c6, TypeName(c6), IsDate(c6)
Debug.Print c7, vbTab, TypeName(c7), IsDate(c7)
Debug.Print c8, TypeName(c8), IsDate(c8)
Debug.Print c9, vbTab, TypeName(c9), IsDate(c9)
Debug.Print c10, TypeName(c10), IsDate(c10)</pre>
	<p>Çıktı şöyle olacaktır:</p>
	<p><img src="/images/vbafunctarihsel4.jpg"></p>
	</div>
	<h2 class="baslik">Diğer işlemler</h2>
	<div class="konu">
	<h3>IsDate ile "Tarih mi?" sorgulaması</h3>
	<p>Bir değerin tarih olup olmadığını sorgulanması için içeriğin  
	Short  
	Date veya saatli  
	Short  
	Date olması lazım. 
	Long Date veya numerik olursa&nbsp; 
	tarih olarak algılamaz.</p>
	<p>Aşağıdaki örnekte, shcedule edilmiş bir makro ile, içindeki dosya 
	isimleri "Falanfilan raporu - 25.07.2017 Sonuçları.xlsm" gibi olan dosyalara 
	bakıyor ve belirtilen süreden önce eski olanları sildiriyorum. Belirli süreyi 
	tespit etme işlemini "Diğer kodlar" bölümünde yapıyorum, şuan bu ksıım 
	önemli değil. Her rapor için tespit ettiğim bu süre değerini 
	bir <a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">dictionary</a>'de 
	depoluyorum. Mid(isim, Len(isim) - 19, 10) ile kestiğim kısım 25.07.2017'a 
	denk gelen kısım ve bunun tarih olup olmadığını sorguluyorum. Ancak bazen 
	ilgili klasörde manuel kaydedilmiş dosyalar olabiliyor, bunların silinmesini 
	istemiyorum, keza içinde Format ifadesi geçen dosyaların da silinmesini 
	istemiyorum.</p>
	<pre class="brush:vb">
Sub eskilerisil()
'Diğer kodlar
For Each d In 
	klasor.Files
	isim = fso.GetBaseName(d)
	If Len(isim) &gt; 19 Then
		If IsDate(Mid(isim, Len(isim) - 19, 10)) Then tarih = DateValue(Mid(isim,Len(isim) - 19, 10))
	Else
	    tarih = DateSerial(2099, 12, 31)
	End If

	If Not aysonumu( Date - dict(kls)) = True And _
	tarih &lt; Date - dict(kls) And InStr(isim, "Format") = 0 Then
		Kill d
	End If
Next d
End Sub</pre>
	<h3>Saat karşılaştırma</h3>
	<p>Hour'la, TimeSerial veya TimeValeu ile yapabiliriz.</p>
		<pre class="brush:vb">If Hour(Now)&gt;13 Then 'dakika detayı gerekli değilse
If TimeValue(Now) &gt; #1:00 PM# Then
If TimeValue(Now) &gt; #13:00# Then 'otomatik PM'e çevrilir
If TimeValue(Now) &gt; #13:30:00# Then 'otomatik PM'e çevrilir
If TimeValue(Now) &gt; TimeSerial(13, 0, 0) Then</pre>
	<h3>Bir tarihin belirli kısımlarını alma</h3>
	<h4>Parça fonksiyonları</h4>
	<pre class="brush:vb">tarih="31.12.2016 13:35:00"

Debug.Print Year(tarih) '2016
Debug.Print Month(tarih) '12
Debug.Print Day(tarih) '31

'Hour, Minute ve Second ile de saat, dakika e saniye döndürülür</pre>
	<h4>DatePart</h4>
	<p>Yukarıdak parça parça alan fonksiyonların hepsini tek bir fonksiyon ve 
	parametre ile de yapabilriz, mesela yıl almak için DatePart("yyyy",tarih) 
	kullanılabilir. Ancak yukarıdakileri kullamak daha kolaydır. Yani basit 
	parça alma için bu fonksiyonu kallanmaucağz. Bununla beraber yılın kaçıncı günü, 
	yılın hangi çeyreği, 
	yılın hangi haftası v.s için 
	bu fonksiyon  lazım. Mesela bu 
	söylediğim son 3 amaç için sırayla y,q,ww 
	parametrelerini kullanırız.</p>
	<p>Dikkat:<strong>y</strong> ile <strong>yyyy </strong>karıştırmayın. yyyy yılı 
	veriken tek y yılın gününü verir, d de ayın gününü verir. Parametre 
	olarak kullanılabilecek tüm fadelere yine yukarda linkini verdiğimiz
	<a href="https://msdn.microsoft.com/VBA/Language-Reference-VBA/articles/format-function-visual-basic-for-applications">
	bu sayfadan</a> erişebilirsiniz.</p>
	<h4>Ay ve hafta adı</h4>
	<pre class="brush:vb">Debug.Print MonthName(Month(tarih)) 'Aralık
Debug.Print WeekdayName(Weekday(tarih)) 'Pazar</pre>
		<h4>TimeSerial ve DateSerial</h4>
		<p>Yıl,Ay, Gün veya Saat,Dakika,Saniye bilgilerinin bir yerden temin 
		edilmesi durumunda da belirli tarih ve saatler elde edilebilmektedir.</p>
		<pre class="brush:vb">Debug.Print DateSerial(2017,1,5)
Debug.Print TimeSerial(8,3,12)
</pre>
	<h3>Bazı özel tarihler</h3>
	<pre class="brush:vb">Aybaşı:DateSerial(year(Date),month(Date),1)
Aysonu: DateSerial(Year(Date), Month(Date)+1 , 1)-1 veya daha sade olarak DateSerial(Year(Date), Month(Date) + 1, 0)
Haftabaşı: Date - WeekDay(Date, vbUseSystem) + 1 ‘veya vbMonday
Haftasonu:Date - WeekDay(Date, vbUseSystem) + 7
Yılbaşı: DateSerial(Year(Date), 1, 1)
Yılsonu: DateSerial(Year(Date)+1, 1, 0)</pre>
	<h3>Timer ile geçen süreyi hesaplama</h3>
	<p> Timer property'sini genelde bir prosedürün ne kadar sürede çalıştığını 
	bumak amaçlı kullanırız. </p>
	<pre class="brush:vb">Sub timerkontrol()
Dim başlangıç As Single
Dim bitiş As Single
Dim i As Long

başlangıç = Timer

For i = 1 To 100000000 'Bu yapı For-Next döngüsüdür. Şimdilik bu döngünün nasıl kullanıldığını bilmiyor olabilirsiniz, buna takılmayın. Sonraki bölümlerde detaylıca incelenecek.
k = k + 1
Next i

bitiş = Timer

MsgBox ("İşlem süresi:" &amp; vbNewLine &amp; Round(bitiş - başlangıç, 2) &amp; " saniyedir.")
End Sub</pre>
</div>
	</span>
	</asp:Content>
