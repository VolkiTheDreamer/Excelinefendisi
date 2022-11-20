<%@ Page Title='Fonksiyonlar MetinselFonkisyonlar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
	<p><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></p>
	</td><td><asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>String Fonksiyonları</h1>
	<h2 class="baslik">Yerleşik Metin Fonksiyonları</h2>
	<div class="konu">
<h3>Karakter Fonksiyonları</h3>
<p><span class=" keywordler">Asc</span> ve <span class=" keywordler">Chr</span>: Exceldeki, CODE ve CHAR fonksiyonlarının benzeridir. Sırayla bir karakterin ASCII kodunu ve bir 
ASCII kodunun karakter karşılığını verirler. Yani birbirlerini zıttı şeklinde çalışırlar.
</p>

<pre class="brush:vb">
Debug.Print Asc("a") '97
Debug.Print Chr(64) '@
</pre>

<p>Bunların W ile biten iki versiyonu daha var. Bunlarla da ASCII'nin genişletilmiş kümesi olan UNICODE karakter işlemleri yapılır.</p>

<pre class="brush:vb">
For i = 1 To 65535
    Cells(i, 1).Value = ChrW(i)
Next i
</pre>

<h3>Parça alma ve pozisyon fonksiyonları</h3>
<p>
<span class=" keywordler">Left(string,n)</span>:Metnin solundan istenilen uzunlukta(n) parça keser.<br>
<span class=" keywordler">Right(string,n)</span>:Metnin sağından istenilen uzunlukta(n) parça keser.<br>
<span class=" keywordler">Mid(string,k,[n])</span>:Metnin ortasından belirtilen indeksten(k) itibaren belirtilen uzunlukta(n) karakter keser. Son parametre girilmezse 2.parametreden itibaren tümünü keser.<br>
<span class=" keywordler">Len(string)</span>:Metnin uzunluğunu(boyutunu) verir.<br>
<span class=" keywordler">InStr([n], string, substring, [Compare] )</span>:n.karakterden itibaren aramaya başlayarak bir metin içinde başka bir metni veya karakteri arar, indeks numarasını(sırasını) döndürür. Bulamazsa 0 döner. Son opsiyonel parametrenin 
varsayılan değeri vbBinaryCompare'dir(constant olarak 0), yani default arama şekli case-sensitivedir. Küçük/büyük harf ayrımı olmadan arama yapılması isteniyorsa bu parametreye vbTextCompare(1) girilir. Ama bu parametre girildiğinde ilk parametrenin de girilmesi gerekir.<br>
<span class=" keywordler">InStrRev(string, substring, [n], [Compare])</span>:InStr fonksiyonunun aramayı sondan yapan versiyonudur. Ancak bulunan indeks yine baştan sayarak bulunan indekstir. 
Ör:"Ardahan"da a'yı aratırsak, sondan 2. indekste bulur, bunun da baştan sayılan 
indeksi 6'dır, sonuç 6 olur.<br>
<span class=" keywordler">StrReverse(string)</span>:Metni terse çevirir.<br>
<span class="keywordler">Replace(metin,aranan,neyledeğiştir,başlanacak yer,adet,compare)</span>: 
İlgili metin içinde aranan metni istenilen metinle değiştirir. Başlangıç indeksi 
1'den farklı verilebilir, ve sadece belirli adet değişiklik yapılması 
istenebilir.</p>
	<p>
	String tipi ilginç bir tiptir. Programlama dünyasına aşinaysınız 
	bilirsiniz, string yapısı referans tipli ve hantal bir yapıdır. 
	Üzerinde manipülasyon yapıldığında yeni bir kopyası oluşturulur. Replace 
	işleminde de bu olur. Bu da şu demek; Replace edilecek bir metin olmasa 
	bile kopya oluşturulur. O yüzden eğer döngüsel ve uzun bir replace işlemi 
	yapacaksanız önce replace edilecek eleman var mı diye bakmakta fayda var. 
	Aşağıdaki gibi:</p>
	<pre class="brush:vb">If InStr(metin, aranan) &lt;&gt;0 Then 'eleman varmı
   metin=Replace(metin, aranan, değiştir)
End If</pre>
	<p>Bu tür optimazasyon teknikleri için <a href="http://www.aivosto.com/vbtips/stringopt.html">şu siteye</a> bakabilirsiniz.</p>


<p>Yukardaki açıklamalarda metin olarak verilen herşey bir hücrenin içeriği de olabilir. Bu arada yine yukarıdaki açıklamalardan anlaşıldığı üzere []'ler içindeki parametreler opsiyonel parametrelerdir.</p>

<pre class="brush:vb">
metin="Mustafa Kemal Atatürk"
Debug.Print Left(metin,3)  'Mus
Debug.Print Right(metin,3)  'ürk
Debug.Print Mid(metin,3,2)  'st
Debug.Print Mid(metin,3)  'stafa Kemal Atatürk
Debug.Print Len(metin)  '21
Debug.Print InStr(metin,"türk")  '18
Debug.Print InStr(metin,"Türk")  '0, çünkü case-sensitive
Debug.Print InStr(1,metin,"Türk",1)  '18
Debug.Print InStr(metin,"K")  '9, Kemal'in k'si
Debug.Print InStr(metin,"k")  'Atatürk'ün sonunudaki k
Debug.Print InStr(metin,"a")  '5
Debug.Print InStr(10,metin,"a")  '12. Aramaya 10'dan başlar ama bulduğu konumun indeksi 1'den itibaren sayılır. yani bu örnekte 3 değil, 12 bulunur.
Debug.Print InStr(10,metin,"z")  '0
Debug.Print InStrRev(metin,"a")  '17
Debug.Print InStrRev(metin, "a", 16) '12
Debug.Print StrReverse("volkan") 'naklov
</pre>


<h3>Dönüşüm Fonksiyonları</h3>
<p>
<span class=" keywordler">LCase(string)</span>:Metni küçük harfe çevirir.<br>
<span class=" keywordler">UCase(string)</span>:Metni büyük harfe çevirir.<br>
<span class=" keywordler">Str(string) ve CStr(expression)</span>:Str, numerik değeri metinsel ifadeye dönüştürür. 
Aldığı parametrenin tamamen numerik bir 
değer olması gerekir. Str, bu numerik ifadeyi başında bi boşluk olacak şekilde 
metne çevirir. Ör:123'ü " 123" şeklinde 4 karekterli bir metin yapar. CStr ise alfanumerik bir parametre alır 
ve başında boşluk olmadan dönüştürür. 
Bu bağlamda Str bana biraz anlamsız ve gereksiz geliyor. String dönüşümlerinde 
sadece CStr'yi kullanın derim.<br>
<span class=" keywordler">Val(string)</span>:Metinsel ifadeyi rakamsal ifadeye dönüştürür. Dönüş değeri double'dır.<br>
<span class=" keywordler">StrConv(string,tip)</span>:Metni istenen formattaki bir metne çevirir. Tip olarak vbUpperCase(1),vbLowerCase(2),vbProperCase(3),vbUnicode(64) ve vbFromUnicode(128) değerleri girilebilir. Hepsi de 
aşikar olduğu için ayrıca açıklamaya gerek bulmuyorum, aşağıdaki örnekler de zaten yeterli olacaktır.<br>

</p>

<pre class="brush:vb">
Debug.Print LCase("VOLKAN") 'volkan
Debug.Print UCase("volkan") 'VOLKAN
Debug.Print Str(123) '"123"
Debug.Print Val("123") '123
Debug.Print StrConv("merhaba dünya",1) 'MERHABA DÜNYA
Debug.Print StrConv("MERHABA DÜNYA",2) 'merhaba dünya
Debug.Print StrConv("MERHABA DÜNYA",3) 'Merhaba Dünya
</pre>


<p>Kodunuzda kullanıcından bilgi girişi istediğinizde ve bunu bir değerle(şifre 
v.s) karşılaştırdığınızda küçük/büyük ayrımı önemli olacağı için girilen değeri büyük harfe çevirip, karşılaştırdığınız metni de büyük harf hazırlarsanız kullanıcı kaynaklı sorunları çözmüş olursunuz.</p>

<pre class="brush:vb">
sifre=InputBox("Şifreyi giriniz")

If UCase(sifre)="ABCD" Then
' diğer kodlar
Else
   MsgBox "Yanlış şifre girdiniz"
   Exit Sub
End If</pre>

<h3>Boşluklar</h3>
<p>
<span class=" keywordler">Trim(string)</span>:Metnin solundaki ve sağındaki tüm boşlukları siler.<br>
<span class=" keywordler">LTrim(string)</span>:Metnin solundaki tüm boşlukları siler.<br>
<span class=" keywordler">RTrim(string)</span>:Metnin sağındaki tüm boşlukları siler.<br>
<span class=" keywordler">Space(n)</span>:n kadar boşluk üretir.<br>

</p>

<pre class="brush:vb">
For i = 1 To 10
   Cells(i, 1).Value = Space(25 - Len(Cells(i, 1).Value)) & Cells(i, 1).Value
Next
</pre>

<img alt="" src="/images/vbaspacefunc.jpg">


	<h4>Sıfır uzunluklu metin(boş metin)</h4>
	<p>""(içi boş çift tırnak) yazarak sıfır uzunluklu metin elde 
	edebiliyoruz. Bir değerin boş metin olup olmadığını anlamak için <strong>if 
	degisken=""</strong> yöntemi sık kullanılmaktadır. Ne var ki bu yöntem çok sağlıklı 
	değildir. Özellikle büyük bir döngü içindeyken kesinlikle kaçınılmalıdır.</p>
	<p>Alternatif olarak <strong>Len</strong> veya <strong>LenB</strong> kullanılabilir.</p>
	<pre class="brush:vb">If LenB(x)=0 Then
  'diğer kodlar
End If</pre>
	<p>Boş metin atamaları da vbNullString şeklinde yapmanızı tavsiye ederim, "" 
	olarak değil.</p>
	<p>Bu konuda daha detaylı bilgi için 
	<a href="../Fasulye/NeNeredeNasil_NullNothingEmptyveIlkdegeratama.aspx">buraya</a> bakın.</p>
	<h3>Split ve Join</h3>
<p>

<strong>Split</strong> ile, belirli bir ayraçla birbirinden ayrılmış kelimeleri 
birbirinden ayırıp tek boyutlu bir dizi elde ederiz. Mesela bir hücrede ; ile 
ayrılmış sicil numaralarını birbirinden ayırıp, bu kişileri mail sistemindeki 
alıcı listesine tek tek ekleyebiliriz. Bu mail gönderim örneğinin tamamını 
<a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx">şu sayfada</a> göreceğiz, ancak şuan bizi ilgilendiren kısmına bakalım.</p>
	<pre class="brush:vb">Dim emailgrubu As Variant
'........ diğer kodlar

emailgrubu=Split(Activecell.Offset(0,1).Value, ";")
'ilgili hücredeki metin 12345;12456;12894 ise, bunlar birbirinden ayrılacak ve 3 elemanlı bir dizi elde edilecektir
'....... diğer kodlar</pre>
	<p>

	Bir başka örnek de bir hücredeki kelimeleri saydıran veya belirli bir kelimeyi 
	seçen bir Function yazmak olabilir. Benim kullandığım böyle bir fonksiyon 
	var. Microsoft geliştiricileri böyle kritik bir fonksiyonu neden hala yerel 
	fonksiyon listesine eklemiyor, gerçekten hayret ediyorum.</p>
	<pre class="brush:vb">Function kelimesec(hucre As Range, kaçıncı As Byte, Optional ayrac As String = " ")
Dim kelimeler As String

kelimeler= Split(hucre.Value2, ayrac)
kelimesec=kelimeler(kaçıncı - 1)
End Function</pre>
	<p>

	Function konusunu henüz incelemediyseniz çok dert etmeyin, öğrendiğiniz 
	zaman gelip tekrar bu örneği inceleyebilirsiniz.</p>
	<p>

	<strong>Join </strong>ise, Splitin tersi mantıkta çalışır. Dizi 
	elemanlarını, belirli bir ayraç ile metin olarak birleştirir. Hemen örneğe 
	bakalım.</p>

<pre class="brush:vb">
Dim siciller() As Integer
Dim birlesiksiciller As Stiring

'dizi eleman sayısı bir yerden okunur, bu x olsun
Redim siciller(1 to x)
For i = 1 to x
   siciller(i)=cells(i,2).Value 'dizi elemanlarına değer atanıyor
Next i

birlesiksiciller=Join(siciller, ";")
</pre>


<h3>Metin Formatlama</h3>
<p>

String modülünün <span class="keywordler">Format</span> fonksiyonu oldukça 
faydalı bir fonksiyon olmakla birlikte metinler üzerinde kullanımından ziyada 
tarihsel ve numerik alanları üzerinde kullanımı daha yaygındır. O yüzden bu 
kısımdan ziyade Tarih ve Numerik fonksyion sayfalarında ele alacağız. Yine de 
sınırlı olan metin formatlama için
<a href="https://msdn.microsoft.com/VBA/Language-Reference-VBA/articles/format-function-visual-basic-for-applications">
bu sayfaya</a> bakabilrsiniz.</p>
	<h3>

	$ işaretli fonksiyonlar</h3>
	<p>

	VBA'de bazı fonksiyonların aynısının sonu $ ile biten versiyonları 
	mevcuttur. Bunların $'sız versiyonları <strong>string tipli variant </strong>döndürürken, $'lı 
	versiyonları standart string döndrür. Bunun daha derin bir anlamı var, o da 
	şu. $'sız olanlar, üzerinde işlem yaptığı metin null değer ise hata vermekzen 
	$'lı olanlar hataya neden olur. Ayrıca $'sız olanlar string döndürdükleri 
	için hafıza ve dolayısıyla performans avantajı sunarlar.</p>

<pre class="brush:vb">
'Şu kod hata vermez
Dim x
x = Null
Debug.Print Left(x, 3)

'Bu kod hata verir
Dim x
x = Null
Debug.Print Left$(x, 3)
</pre>
	</div>

	<h2 class="baslik">Regex</h2>
	<div class="konu">
	<p>Regex dünyası başlı başına ayrı bir dünya olmakla birlikte burda küçük birkaç örnek yapacağız. Konuyu daha iyi kavramak adına <a href="https://regexr.com/">şu</a> ve <a href="https://regex101.com/">şu</a> sayfalarda denemeler yapmayı ihmal etmeyin. Bunlar, programlama dilinden bağımsız olarak genel regex kullanımına yönelik test sayfalarıdır. Genel yapı dilden bağısmız olarak aynı olmakla birlikte syntax tabiki dilden dile değişmektedir.</p>
        <h3>Nedir ve ne işe yarar?</h3>
        <p>Yukarıda, <strong>Instr</strong> ile belirli bir metni başka bir metin içinde arıyorduk, keza <strong>Replace</strong> ile de belirli ifadeleri değiştiriyorduk. Ancak ya elimizde direkt bir ifade yoksa, ve onun yerine belirli bir şablon varsa? Örneğin bir metnin içinde &quot;sız, siz, suz, süz&quot; ifadelerinden biri var mı diye bakmak istiyoruz. Tek tek bu 4 sorgulamayı yapmak yerine &quot;s-herhangi bir karakter-z&quot; kalıbını aramak daha mantıklı olurdu değil mi?</p>
        <p>Keza içinde herhangi bir rakam geçen metinleri, veya bir rakam bir sayı bir rakam(Ör:5a3, 4r8) gibi bir desen içeren metinleri filtrelemek isteyebiliriz.</p>
        <p>İşte bu tür, desen(pattern) bazlı ifadelere <strong>düzenli ifadeler (regular expressions)</strong> deniyor, bunların kısaltması da <strong>regex </strong>oluyor.</p>
        <p>Örnek kullanım alanları</p>
        <ul>
            <li>Mail adresi, telefon numarası analizlerinde ve/veya kurala uygun veri girişi yapılmasında</li>
            <li>Karakter maskelemede(şifreleme)</li>
            <li>html parsing veya code syntax işaretlemede. (Gerçi mümkün mertebe html parse etmede regex kullanmayın. Bunlar için diğer araçları kullanın)</li>
        </ul>
        <h3>Nasıl kullanılır?</h3>
        <p>Öncelikle Regex kullanabilmeniz için Tools&gt;Reference&#39;ten <strong>Microsoft</strong> <strong>VBScript Regular Expressions 5.5 </strong>kütüphanesini eklemeniz gerekiyor. (veya Late Binding ile de kullanabilirsiniz, ama bu durumda tabiki intellisense&#39;i unutun)</p>
        <p>Sonra bazı kavramları&nbsp;bilmek gerekiyor. Bu kavramlara geçmeden önce şunu belirtmek isterim ki, eğer başka bi dilde Regex kullandıysanız işlerin VBA&#39;de biraz farklı olduğunu görebilirsiniz.</p>
        <h4>Literal ve Meta karakterler</h4>
        <p>Literal karakterler bildiğimiz karakterlerdir. Örneğin bir metinde abc (literal) ifadesini arıyorsak, parametre olarak direkt abc&#39;yi veririz.</p>
        <p>Metakarakterler ise biraz daha soyuttur. Bunlardan birden fazla karaktere denk gelebilrler.</p>
        <p>
            <img alt="regex" src="../../images/regex.jpg" style="width: 1233px; height: 817px" /></p>
        <p>
            <strong>Önemli</strong>: Farkettiyseniz bazı işlemler için daha kısayollar tasarlanmış. Mesela sayı seçimi için [0-9] yerine \d yapabiliyoruz. Ama mesela eğer kelime kelime arama yapacaksanız, &quot;.&quot; gibi tüm whitespace&#39;leri de içeren bir operatörü kullanmamanız gerekir. Mesela, bir mail adresi incelemesi yapalım. Amacımız @ işaretinden önceki kısmı bulmak. Bir mail adresinden neler olur, harf, rakam , nokta veya alt tire değil mi. Bunu normalde şöyle ifade edebiliri. &quot;[a-zA-Z0-9_.]+@&quot;. Bunu daha kısa tutmak için &quot;.+@&quot; yaparsak &quot;<strong><a href="mailto:benim%20mail%20adresim%20volkan.yurtseven@hotmail.com">benim mail adresim volkan.yurtseven@hotmail.com</a></strong>&quot; metni içinden &quot;<strong>benim mail adresim volkan.yurtseven@</strong>&quot; döndürür. Kısaltma yapmak için nokta kullanmadık, peki \S kullanabilir miyiz? Yine olmaz, çünkü bu durumda yanlışkla yazılmış bir &quot;volkan+yurtseven@hotmail.com&quot; &#39;daki + işareti koşulu sağladığı için o da gelir. O yüzden sırf kodu kısa tutmak amacıyla yanlış pattern kullanmamalısınız. 
            Ancak şöyle bir alternatif de mevcut: &quot;(\w|\.)+@&quot;, yorumu basit: bir alfanumerik karekter veya nokta&#39;dan en az bir tane olsun sonra bir de @.</p>
        <h3>Regex VBA </h3>
        <h4>Property ve Metodlar</h4>
        <p>
            Regex nesnesinin propertyleri şunlardır.</p>
        <ul>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">Pattern</strong><span>:</span> En önemli propertysi bu olup, aradığımız patterni yazarız.</li>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">IgnoreCase</strong>: True ise, küçük büyük harf ayrımı yapmaz. (Türkçe karakter için aşağı bakınız)</li>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">Global</strong><span>:</span> True ise tüm eşleşmeler, False ise ilk eşleşme</li>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">MultiLine</strong><span>:&nbsp;</span>True ise, satır satır bakılır.</li>
        </ul>
        <p>
            Metodlar ise şunlardır.</p>
        <ul>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">Test:</strong> Patterne bakar, bulursa True bulamazsa False döner</li>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">Replace</strong><span>:</span> Eşleşmeleri başka bir ifade ile değiştirir</li>
            <li style="box-sizing: border-box;"><strong style="box-sizing: border-box; font-weight: 700;">Execute</strong><span>:&nbsp;Eşleşmeleri bir MatchCollection nesnesi olarak döndürür</span></li>
        </ul>
        <p>
            Şimdi de birkaç örnek yapalım.
        </p>
        <p>
            İlk olarak bir prosedür örneği. Bu prosedürle bir alandaki hücrelerde parantez içinde bulunan ifadeleri sildireceğiz.</p>
        <p>
            Örnek değerlerimiz aşağıdaki gibi:</p>
        <p>
            <img src="../../images/vba_regex2.jpg" style="width: 471px; height: 100px" /></p>
        <p>Burdaki parantez içindeki bilgileri yokedip şu hale getirmek sitiyoruz.</p>
        <p>
            <img src="../../images/vba_regex3.jpg" style="width: 263px; height: 115px" /></p>
        <pre class="brush:vb">
Sub ParantezYoket()

Dim Str As String
Dim Replace_Str As String
Dim regexObject As New RegExp
Dim matches As MatchCollection
Dim match As match

With regexObject
    .Pattern = "\s?\([^)]+\)"
                '\s?:parantez öncesi boşluk varsa veya yoksa, ?'nin 0 veya 1 tekrar olduğunu unutmayın
                '\( : açılış parantezi
                '[^)]+: bir veya daha fazla parantez olmayan karakter
                '\) : kapanış parantezi
    .Global = True
End With

Set alan = Range("a1:a3")
For Each cell In alan
    Set matches = regexObject.Execute(cell.Value)
    For Each match In matches
      temp = regexObject.Replace(cell.Value, "")
    Next match
    cell.Value = temp
Next cell

End Sub        </pre>
        <p>Şimdi bir de UDF örneği yapalım. Bu örnekte, bir kolondaki bulunan kişi isimlerinin(2-3-4 kelimeden oluşabilir) baş harflerini bırakıp kalanını yıldızlama işlemi yapacağız, yani maskeleme uygulayacağız.</p>
        <p>
            <img src="../../images/vba_regex4.jpg" style="width: 265px; height: 110px" /></p>
        <pre class="brush:vb">
Function Maskele(metin As String, Optional pattern As String = "\B[A-Za-z]")
Dim reg As New RegExp
reg.Global = True
reg.pattern = pattern

Maskele = reg.Replace(metin, "*")

End Function            </pre>
        <p>
&nbsp;Bu örnekte sadece ilk karekterleri maskelemek istedik, o yüzden optional parametre verdik, ama istenirse başka bir pattern de verilerek onların maskelenmesi sağlanabilir.</p>
        <h5>Türkçe karakter ve Unicode</h5>
        <p>Yukarıdaki maskeleme örneğinde isimlerden birinde Türkçe karakter olsaydı sıkıntı yaşardık. Mesela &quot;şükran dağıstan&quot; &quot;şük*** d*ğıs***&quot; olarak maskelenirdi. Bunu engellemek için unicode ifadelerini kullanmak gerekiyor ancak bu örnek için ben bir türlü uygun çözümü bulamadım, zira bu unicode konusu biraz karışık. Onun yerine <strong>metin</strong> değişkenini şöyle değiştirme yoluna gittim: &quot;metin = Replace(Replace(Replace(Replace(Replace(metin, &quot;ş&quot;, &quot;s&quot;), &quot;ü&quot;, &quot;u&quot;), &quot;ı&quot;, &quot;i&quot;), &quot;ç&quot;, &quot;c&quot;), &quot;ğ&quot;, &quot;g&quot;)&quot;. Tabiki ilk karakter de türkçe karakterse bunu da değiştirmiş oluyoruz, o yüzden ideal bir çözüm değil. Diğer dillerin çoğunda unicode desteği internal olarak geliyor, malesef VBA&#39;de bu destek yok.</p>
	</div>
</asp:Content>
