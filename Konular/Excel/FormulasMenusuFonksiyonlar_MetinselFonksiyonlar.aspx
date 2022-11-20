<%@ Page Title='FormulasMenusu1 MetinselFormuller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'>

<table><tr><td>

<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr>

</table></div>


<!--***********************************************************************************-->
<h1>Metinsel Formüller</h1>
<p>Excel her ne kadar sayılarla uğraşan bir program görüntüsü verse de, çok sayıda metin manipülasyonu da yapmamız gerekecektir. Excel'in bu konuda oldukça büyük bir metin fonksiyon kütüphanesi var. Bazen bunları tekli bazen de içiçe birkaçını birden kullanmak gerekecektir. Hatta bazı formüller tahmininizeden çok daha uzun olabilecektir.</p>

<p>Bununla birlikte bazı durumlarda Excelin metin formüllkerinin yetersiz kaldığını göreceksiniz. Veya yazdığınız formüller aşırı uzun olacak, bunları her ihtiyaç duyduğunuzda böyle uzun uzun yazmak çok zor olabilecektir. İşte böyle durumlarda UDF oluşturma zamanı gelmiştir. Ancak siz yine de uzun formüllerle başa çıkmak için çeşitli <a href="FormulasMenusuDiger_PufNoktalari.aspx#uzunformul">pratik yöntemleri</a> kullanabilirsiniz.</p>

<p>Son olarak, bazı durumlarda formül yazmak yerine <a href="/Konular/Excel/HomeMenusu_Doldurma.aspx">Filling(Doldurma)</a> işlemi yapmanız da sizi uzun formüllerden kurtarabilir.</p>

<p>Şimdi bu kategorideki önemli formüllere bakalım. Fonksiyonları belirli alt kategoriler altında topladım, buna göre ilerleyeceğiz.</p>

<!--********************************************************************************************************************************************-->
<h2 class="baslik">Metinden parça alma</h2>


<div class="konu">
<p><span class=" keywordler">TRIM(Metin)</span>:Trim, bir hücredeki kelimeler arasındaki tekil boşluklar dışındaki tüm boşlukları temizler. Kelimeler arasında birden fazla boşluk varsa bunları da 1e indirir. A2 hücresinde "&#160&#160&#160Volkan&#160&#160&#160&#160Yurtseven&#160&#160&#160&#160&#160&#160&#160&#160" yazdığını düşünelim.</p>
<pre class="formul">=TRIM(A2) //Volkan Yurtseven</pre>


<p><span class=" keywordler">CLEAN(Metin)</span>: Clean, ekran görünmeyen ilk 32 ASCII karakteri temizler. Mesela Enter'ın kodu 10 olup bunu temizler. TRIM'le birlikte kullanıldığında temizlenecek boşluk benzeri karakter sayısı daha çoğalmış olur.</p>
<pre class="formul">=CLEAN(TRIM(A2))</pre>


<p><span class=" keywordler">LEFT(Metin,kesilecek karakter sayısı)</span>:LEFT, bir hücrenin içindeki metinin solundan, belirtilen adette karakteri keser. Kaç karakter kesmeniz gerektiğini bilmediğiniz bazı durumlarda <strong>FIND, SEARCH, LEN</strong> gibi diğer fonksiyonlardan yararlanabilrsiniz.  A2 hücresinde "TR123456" gibi, ilk iki hanesi ülke kodu olan metinler olduğunu düşünelim. Ülke kodunu almak için soldaki 2 haneyi keseriz.</p>
<pre class="formul">=LEFT(A2;2) //TR</pre>

<p>Mesela, bir hücre grubunda isim ve soyisimler var diyelim, siz ilk ismi almak istiyorsunuz, aşağıdaki gibi bir formül yazabilirsiniz. A2 hücresinde "Volkan Yurtseven" yazdığını düşünelim.(FIND fonksiyonunu biraz aşağıda göreceğiz.)</p>
<pre class="formul">=LEFT(A2;FIND(" ";A2)-1) //Volkan</pre>

<p><span class=" keywordler">RIGHT(Metin,kesilecek karakter sayısı)</span>:RIGHT, bir hücrenin içindeki metinin sağından, belirtilen adette karakteri keser. Yine, kaç karakter kesmeniz gerektiğini bazı durumlarda <strong>FIND, SEARCH, LEN</strong> gibi diğer fonksiyonlardan yararlanabilrsiniz. A2 hücresinde "TRM456" yazdığını düşünelim.sadece sağdaki 3 hane olan sayıları almak istiyoruz.</p></p>
<pre class="formul">=RIGHT(A2;3) //456</pre>

<p>Yine bir hücre grubunda isim ve soyisimler var diyelim, siz soyismi almak istiyorsunuz, Bunu yapmak LEFT ile isim almak kadar kolay değil, çünkü 2 isimli kişiler işleri zorlaştırmaktadır. Buna ait örneği en alttaki Çeşitli Örnekler bölümünde bulabilirsiniz. Biz yine de sadece bir isimli kişilerin olduğu bir listede bunu nasıl yapabiliriz, buna bi bakalım. A2 hücresinde "Volkan Yurtseven" yazdığını düşünelim.</p>

<pre class="formul">=RIGHT(A2;LEN(A2)-FIND(" ";A2)) //Yurtseven</pre>

<p><span class=" keywordler">MID(Metin,kesmeye başlanacak yer,kesilecek karakter sayısı)</span>:MID bir hücrenin içindeki metinin ortasından bir yerden belirtilen adette karakteri keser. Nereden başlayacağınızı ve kaç karakter kesmeniz gerektiğini bilmediğiniz bazı durumlarda <strong>FIND, SEARCH, LEN</strong> gibi diğer fonksiyonlardan yararlanabilrsiniz.</p>

<p>Mesela elinizde 25 haneli hesap numaraları var diyelim. Bunların 12 ile 15 arasındaki karakterleri "şube kodu" olsun. Şube kodunu buradan almak için 12'den başlayıp, 4 karakter almamız gerekir.</p>

<pre class="formul">=MID(A2;12;4)</pre>
</div>
<!--********************************************************************************************************************************************-->
<h2 class="baslik">Parça bulma ve değiştirme</h2>
<div class="konu">
<p><span class=" keywordler">FIND(bulunacak metin, neyin içinde,[nerden başlanacak]):</span>FIND, bir metin içinde başka bir metin/karakter arayıp onu bulduğu yerin konumunu(kaçında karakter olduğunu) gösterir. Aranan karakterden çok sayıda bulunsa bile ilk bulunduğu yerin konumu gelir. Üçüncü parametre opsiyonel olup default(varsayılan) değeri 1'dir. <strong>Bu fonksiyon, büyük küçük harf ayrımına duyarlıdır</strong>. Aşağıdaki örnekte, hücredeki ilk boşluk karakterinin konumu gelir.</p>

<pre class="formul">=FIND(" ";A2)</pre>

<p>Bu yöntem, her hücrenin ikinci kelimesini bulmak için kullanılabilir. Bunun için üçüncü parametre olan arama pozisyonunu, ilk boşluğu bulduğumuz yer + 1 şeklinde belirleyeriz. Ancak sonrasında kaç karakter keseceğimizin cevabı biraz daha karışıktır. Önce formüle bakalım, sonra nasıl işlediğine.</p>

<pre class="formul">=MID(A2;FIND(" ";A2)+1;FIND(" ";A2;FIND(" ";A2)+1)-FIND(" ";A2)-1)</pre>

<p>İlk parametre basit, bunu geçelim. İkinci parametreyi FIND(" ";A2)+1 diyerek belirledik. Son parametre FIND(" ";A2;FIND(" ";A2)+1)-FIND(" ";A2)-1 formülüyle belirledik. Bunu da parçalara ayıralım. Öncelikle ikinci boşluk karakterini bulmamız lazım, ikinci boşluğu bulmak için de aramaya ilk boşluğu bulduğumuz konumdan sonra başlamalıyız. ilk boşluğu nasıl bulmuştuk, FIND(" ";A2) ile, buna 1 ekliyoruz. Şimdi elimizde ikinci boşluğu bulduğumuz FIND(" ";A2;FIND(" ";A2)+1) formülü var. Ama bu kadar karakter kesersek beklenmeyen bir sonuç olur. O yüzden son olarak, ilk boşluğun konumu bulduğumuz formülü de son formülümüzden çıkarmamız lazım: FIND(" ";A2;FIND(" ";A2)+1)-FIND(" ";A2)-1</p>

<p>Böyle uzun formüllerde F9 tuşu ile formülün hangi parçasının ne döndürdüğünü görmek oldukça pratiklik sağlamaktadır. Bu özelliği sık sık kullanmanızı tavsiye ederim.</p>

<p>Şimdi diyeceksiniz ki, ikinci kelimeyi bulduk da, üçüncü/dördüncü/v.s kelimeleri nasıl buluruz? Böyle sürekli 3.boşluğu, 4.boşluğu bul, öçncekinin konumunda çıkar vs. yoluna girersek formülümüz saçma derecede uzar ve pratik bir yöntem de olmaz. Bunu nasıl yapacağımızı, Çeşitli Örnekler bölümünde göreceğiz. Bir dieğr altenratif de UDF kullanmak veya Excelin sonraki sürümlerini beklemek olacaktır. Ben hala kelime sayma, kelime seçme gibi temel formüllerin 2016 sürümünde bile olmamasını hayretle karşılıyorum. Neyseki çözümsüz değiliz, UDF'ler bize bu konularda büyük kolaylıklar sağlıyor.</p>


<p><span class=" keywordler">SEARCH(bulunacak metin, neyin içinde,[nerden başlanacak])</span>:FIND gibi çalışır. FIND'dan farkı küçük-büyük harf duyarlılığı yoktur, ayrıca joker karakterleri(? ve *) kullanmanıza izin verir.
</p>

<p>Bu fonksiyon, filtrelemeye alternatif olarak kullanılabilir ve Filtrelemeden daha esnektir. Mesela aşağıdaki formülle bir hücrede Volkan kelimesi var mı diye bakıyorum. </p>

<pre class="formul">=ISNUMBER(SEARCH("Volkan";A2))</pre>

<p>Bu da FIND gibi başka fonksiyonlarla kombine bir şekilde kullanılır.</p>

<p><span class=" keywordler">REPLACE(değişiklik yapılacak metin, kaçıncı karakterden başlanacak, kaç karakter değişecek, neyle değiştirilecek):</span>Konum bilgilerinden yararlanmayı gerektiren bu fonksiyon ile bir metin içinde belirli bir karakteri veya karakter grubunu başka bir karakter veya karakter grubu ile yer değiştirtiriz. Değiştirilen karakter ile yerine konan karakterler aynı uzunlukta olmak zorunda değildir. Mesela bir hücredeki ilk 10 karakteri sadece "?" karakteri ile değiştirebilirsiniz. Örneği de şöyle olacaktır. Metnimiz bir kredi kartı numarası olsun: 1111 2222 3333 4444</p>
<pre class="formul">=REPLACE(A2;1;10;"?") //?3333 4444</pre>

<p>FIND ve REPLACE sıkça beraber kullanılır. Nereden başlanacak ve kaç karakter değiştirilecek gibi soruları bulurken FIND fonksiyonundan faydalanırız. Mesela bir hücre grubu içinde - işareti olan yere kadarki tüm metni silmek istesek ne yapardık? Yolumuz şu olurdu:"-" işaretinin konumunu FIND ile bul, 1'den başlayıp bu konuma kadar olan tüm karakterleri "" ile yer değiştir(Bir metni "" ile yer değiştirmek onu silmektir). Metnimiz "Fatih-İstanbul" olsun.</p>
<pre class="formul">
=REPLACE(A2;1;FIND("-";A2);"") //İstanbul
//Tam tersi işi yapmak içinse - işaretini bulduğumuz yerden başlayıp, 'metnin tüm uzunluğu - işaret konumu+1' kadar gidip bunları yokederiz
=REPLACE(A2;FIND("-";A2);LEN(A2)-FIND("-";A2)+1;"")//Fatih
</pre>

<p>Konumdan ziyade doğrudan içerik değiştirmekle ilgileniyorsak bir alttaki fonksiyona bakmamız gerekir.</p>

<p><span class=" keywordler">SUBSTITUTE(değişiklik yapılacak metin, neyi değiştireceğiz, neyle değiştireceğiz, [kaçıncı]):</span>Bu fonksiyon ile bir metindeki belirli karakterleri(veya metinleri) başka karakterlerle değiştiriyoruz. Son parametre seçimlik olup kaçıncı eşleşmeden sonra değişklik yapılması gerektiğini belirtmiş olursunuz, belirtmezseniz tüm eşleşmeler için işlem yapılır. Basit bir örneğe bakalım.</p>

<pre class="formul">
=SUBSTITUTE(A2;"ı";"i") //Tüm ı'lar i yapılır
=SUBSTITUTE(A2;"0";"",2) //İlk sıfır dışındaki tüm sıfırları yok eder.
</pre>

<p><strong>NOT:</strong>Bu fonksiyon küçük/büyük harf ayrımına duyarlıdır.</p>

<p id="supertrim">Yukarda CLEAN ve TRIM fronkisyonlarının birlikte kullanımı durumunda boşluk ve ilk 31 ascii kodunsa sahip görünmez karaterleri silebildiğini söylemiştik. Ancak görünmeyen başka karakterler de vardır ve bunlar için SUBSTITUTE'lu formül veya 
<a href="../VBAMakro/Fonksiyonlar_ExcelicinUDFKullaniciTanimliFonksiyonlar.aspx#supertrim">UDF</a> yazmak gerekir. Mesela birçok durumda karşınıza çıkabilecek olan non-breaking-space karakterinin kodu 160 olup bunu CLEAN ve TRIM ile yokedemezsiniz, SUBSTITUTE ile yoketmeniz gerekir. Tüm böyle gereksiz karakterleri yoketmek için açağıdaki gibi bir formül yazılabilir.</p>
	<pre class="formul">=TRIM(CLEAN(SUBSTITUTE(A2;CHAR(160);"")))</pre>

</div>

<!--********************************************************************************************************************************************-->
<h2 class="baslik">Dönüştürme</h2>
<div class="konu">
<p><span class=" keywordler">CHAR(Sayı)</span>:Kendisine verilen sayı parametresi ile ilgili Ascii karakterlerini döndürür. Sayı 1-255 arası değer alabilir. Özellikle formüller içinde Enter veya tırnak kullanmak için kullanılır. Özellikle bir hücreden mail body'si okurken paragraflar arasına Enter koymak amacıyla kullanılır.</p>

<pre class="formul">
=CHAR(10) //Enter görevi görür. 
=CHAR(34) //Tırnak işaretidir.
</pre>

<p>A kolonunda A1 hücresinden başlayıp aşağı doğru 1'den 255e kadar sayıları yazıp, yanlarına bu formülü yazarak hangi sayı ile hangi kodun döndüğünü görebilirsiniz.</p>

<p><span class=" keywordler">CODE(Karakter)</span>:CHAR'ın tersi gibi çalışır</p>
<pre class="formul">=CODE("@") //64 döndürür</pre>

<p><span class=" keywordler">UNICODE ve UNICHAR(2013)</span>:CHAR ve CODE'un 1-255in dışındaki sayılarla da çalışmasını sağlayacak şekilde genişletilmiş halleridir. Kendiniz deneyip görebilrsiniz. Bunlarla her tür karakteri basabilir, veya ilgili karakterin kodunu elde edebilirsiniz.</p>

<p><span class=" keywordler">LOWER(Metin)</span>:Verilen metnin tüm harflerini küçük harf yapar.</p>
<pre class="formul">=LOWER("MERHABA VOLKAN") //merhaba volkan</pre>

<p><span class=" keywordler">UPPER(Metin)</span>:Verilen metnin tüm harflerini büyük harf yapar.</p>
<pre class="formul">=LOWER("merhaba volkan") //MERHABA VOLKAN</pre>

<p><span class=" keywordler">PROPER(Metin)</span>:Verilen metnin içindeki kelimelerin sadece ilk harflerini büyük, diğerlerini küçük harf yapar</p>
<pre class="formul">=PROPER("MERHABA VOLKAN") //Merhaba Volkan</pre>

<p id="formattext"><span class=" keywordler">TEXT(Metin,DönüştürmekİstediğinizFormat)</span>:Elinizdeki metni daha okunaklı ve anlaşılır hale getirmek için belirli bir formata sokmaya yarar. Özellikle başka formüllerle birarada kullanıldığında anlamlıdır. </p>
<p>Mesela aşağıdaki resimde görüldüğü üzere, A1 hücresinde =TODAY() formülü var, yani dosyayı açtığımızda değişen bir tarih içeriği ile karşı karşıyayız. Biz bu tarihi B2 hücresinde kullanmak istiyoruz, "XXX tarihi Performans Raporu" diye de raporu bastıracağız. Formülü aşağıdaki gibi yazarsak; </p>

<pre class="formul">=A1 & " Performans Raporu"</pre>

<p>görüntüsü şöyle olur, ki bunu istemeyiz. O yüzden bu tarihi formatlamamız gerekir. İşte burada <strong>TEXT</strong> formülü devreye girer.</p>


	<img src="/images/excelmetinseltext1.jpg">

<pre class="formul">=TEXT(A1;"dd.mm.yyyy") & " Performans Raporu"</pre>

<p>Sonuç istediğimiz gibi olacaktır:</p>

	<img src="/images/excelmetinseltext2.jpg">

<p><a href="https://support.office.com/en-us/article/Format-a-date-the-way-you-want-8e10019e-d5d8-47a1-ba95-db95123d273e">Burada</a> tarih formatlarını nasıl kullanacağınızla ilgili detay bilgi bulunmaktadır. 
(Bölgesel ayarlara göre dd.mm.yyyy veya gg.aa.yyyy şeklide bir pattern girilmesi 
gerekebilir. İlki İngilizce day day.month month.year year year year 
kelimelerinin ilk harfleri; ikincisi Türkçe gün gün.ay ay.yıl yıl yıl yıl 
kelimelerinin ilk harfleridir)</p>

<p><strong>Örnek 2:</strong>Şube isimleriyle şubelerin Hedef Gerçekleştirme(HG%) oranlarını birleştirmek istiyorsunuz diyelim. A kolonunda şube isimleri var, B kolonunda HG% oranları, ama HG% oranları şu şekilde: 0,98745210023. Siz mesela bu değeri %98,7 görmek istiyorsunuz. Formülümüz şöyle olacaktır</p>

<pre class="formul">=A1&"-"&TEXT(B1;"%0,0")</pre>


<p><a href="https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68">Burada</a> da sayı formatlarıyla ilgili detaylı bilgi bulunmaktadır.</p>

<p><span class=" keywordler">VALUE(Metin/Tarih)</span>:Parametre olarak verilen ifade sayıya çevrilebilecek bir metinse veya tarihse bunu sayıya çevirir. Özellikle başka bir veri kaynağından gelen datalarda sayılar bazen metin formatında sola dayalı şekilde gelir. Bunları sayıya çevirmede kullanılabilir.</p>

<pre class="formul">=VALUE("500") //500</pre>

</div>

<!--********************************************************************************************************************************************-->
<h2 class="baslik">Diğer</h2>
<div class="konu">
<p><span class=" keywordler">CONCAT(Alan)-(2016)</span>:Belirli bir alandaki hücre içeriklerini aralarında bir ayraç olmadan birleştirir.</p>
<pre class="formul">=CONCAT(A2:C2)</pre>

<p><span class=" keywordler">CONCATENATE(Metin1,Metin2,....)</span>:Parametre olarak girilen metinleri birleştirir. CONCAT'tan farklı olarak metinler/hücreler tek tek girilir. Ayraç belirtmek istenirse de her metin arasına tek tek ayraçlar da girilmek zorundadır.</p>
<pre class="formul">=CONCATENATE(A2;" ";B2;" ";C2)</pre>
<p>Bu yöntem, ilgili metinleri & işareti ile birleştirmeye benzer.</p>
<pre class="formul">=A2&" "&B2&" "&C2)</pre>

<p><span class=" keywordler">TEXTJOIN(Ayraç,BoşHücrelerDikkateAlınsınmı,Alan)-(2016)</span>:Belirli bir alandaki hücre içerikleri, aralarında belirlenen ayraç olacak şekilde birleştirir. Boş hücreler dikkate alınsın mı, alınmasın mı, bunun için de TRUE/FALSE şeklinde bir parametre girilir.</p>
<pre class="formul">=TEXTJOIN(" ",TRUE,A2:C2</pre>

<p><span class=" keywordler">EXACT() </span>:İki hücrenin aynı olup olmadığını test eder. Bu kontrolü yaparken küçük/büyük harf ayrımını da dikkate alır. Mesela A1'de Volkan ve B1'de VOLKAN yazıyor olsun.</p>
<pre class="formul">
=A1=B1 //TRUE döndürürken
=EXACT(A1;B1) //FALSE döndürür
</pre>

<p>Bu fonksiyon, başka fonksiyonlarla birarada kullanılabilir ve bu şekilde kullanımı daha yaygındır.</p>
<pre class="formul">
=INDEX(Z:Z;MATCH(EXACT(A2);Y:Y;0))
</pre>

<p><span class=" keywordler">LEN(Metin):</span>Bir hücredeki metnin karakter sayısını verir. Ör:A kolonundaki metinlerden sadece 5 haneli olanları seçmek istiyoruz diyelim, şu formülü yazıp en sona kadar aşağı çekelim ve sonra da 5 olanları filtreleyelim.</p>
<pre class="formul">=LEN(A2)</pre>

<p>Bu formül de kendi başına kullanımdan ziyade, aşağıdaki Çeşitli Örnekler bölümünde göreceğiniz üzere diğer fonksiyonlarla birlikte daha çok kullanılır.</p>

<p><span class=" keywordler">REPT(metin,tekrarsayısı)</span>: Bu fonksiyon ile belirli bir metni(genelde tek bir karakter) belirtilen sayıda tekrar ettiriz. Bu fonksiyon da genelde başka fonksiyonlarla birlikte kullanılır.</p>

<p>Diyelim ki elinizde farklı uzunluklarda müşteri numaraları var, 2 haneden 8 haneye kadar değişkenlik gösteriyor. BT ekibinize bu müşteri numaralarını içeren bir liste göndereceksiniz, ancak BT ekibi diyor ki "bana bu müşteri numaralarını 10 haneli gönder. Başlarında 0 olsun(BT ekipleri bu tür taleplerle gelirken genelde sayısal verilerin önünde 0, metinsel verilerin önünde/sonunda boşluk isterler). Bu durumda formülümüz şöyle olacak:</p>

<pre class="formul">=REPT(0;10-LEN(A2)) & A2</pre>

<p>Önce LEN(A2) ile kaç karakter olduğunu buluyoruz. Sonra bunu 10'dan çıkararak kaç tane 0 ekleyeceğimzi buluyoruz, REPT ile bu kadar 0 üretiyoruz, ve son olarak da bunu metnin kendisi ile birleştiriyoruz.</p>

</div>


<!--********************************************************************************************************************************************-->
<h2 class="baslik">Çeşitli Örnekler</h2>
<div class="konu">

<h4  class="baslik">Bir hücredeki kelimeleri saymak</h4>
<div>
<p>İzlenecek yol: Hücredeki toplam karater sayısını LEN ile bulalım, buna X diyelim. Sonra hücredeki boşlukları SUBSTITUTE ile yokedip kalan kısmın uzunluğunu bulalım, buna da Y diyelim. X-Y+1 aradığımız sorunun cevabı olacaktır.</p>
<pre class="formul">=LEN(A1)-LEN(SUBSTITUTE(A1;" ";""))+1</pre>
</div>


<h4 class="baslik">Bir hücredeki metinden son kelimeyi almak</h4>
<div>
<p>İzlenecek yol: Hücredeki bütün boşlukları 30 adet _ ile değiştiriyoruz. Buna X diyelim. X'in sağdan 30 karakterini alalım, buna Y diyelim. Sonra da _ karakterlerini "" değiştirelim yani bunları uçuralım, voila!</p>
<pre class="formul">=SUBSTITUTE(RIGHT(SUBSTITUTE(A1;" ";REPT("_";30));30);"_";"")</pre>
<p>Aşama aşama bakalım: Hücre içieriği şu olsun:<strong>Mustafa Kemal Atatürk</strong></p>
<pre>
X=Mustafa______________________________Kemal______________________________Atatürk
Y=_______________________Atatürk
Sonuç=Atatürk
</pre>

<p>Neden 30 sayısını kullandık. Çünkü bi kelimenin 30 harften daha uzun olacağını sanmıyorum. Sizin elinizde daha uzun kelimelerden oluşan bir liste varsa 30 yerine 50 veya 100 de kullanabilirsiniz.</p> 
</div>

<h4 class=baslik>Bir hücredeki metinden n. kelimeyi almak</h4>
<div>
<p>İzlenecek yol: Hücredeki bütün boşlukları 100 adet _ ile değiştiriyoruz. Buna X diyelim. X'in ortadan (n-1)*100+1. karakterinden seçmeye başlayıp 100 karakter alalım, buna Y diyelim. Sonra da _ karakterlerini "" değiştirelim yani bunları uçuralım, evvet!</p>
<p>Aşama aşama bakalım: Hücre içeriği yine aynı olsun:<strong>Mustafa Kemal Atatürk</strong>Bu sefer ikinci kelimeyi seçeceğiz. (2-1)*100+1=101. karakterden başlayıp 100 karakter seçeceğiz ve ilerleyeceğiz</p>
<pre class="formul">=SUBSTITUTE(MID(SUBSTITUTE(A1;" ";REPT("_";100));101;100);"_";"")</pre>
<pre white-space="pre-wrap">
X=Mustafa____________________________________________________________________________________________________Kemal____________________________________________________________________________________________________Atatürk
Y=_____________Atatürk
Sonuç=Atatürk
</pre>

<p>Burada 100ü kullanma sebebim, n'in değerine göre kolay çarpım yapma isteğimdendir.</p> 

<p>Gerek bunu gerek bir üsttekini, en dıştaki SUBSTITUTE yerine TRIM ile de yapabilirdik, tabi _ yerine boşluk kullanarak. Sadece son örneği yapalım.</p>
<pre class="formul">=TRIM(MID(SUBSTITUTE(A1;" ";REPT(" ";100));101;100))</pre>

</div>


<h4 class=baslik>Bir hücredeki metinden ilk n kelimeyi almak</h4>
<div>
<p>İzlenecek yol: Önceki yöntemlere benzer olarak, hücredeki bütün boşlukları 100 adet _ ile değiştiriyoruz. Buna X diyelim. X'in soldan n*100. karakterini alalım, buna Y diyelim. Sonra da Y içindeki 100 tane _ karakterini "" ile değiştirelim yani bunları uçuralım, yeni değerimiz Z olsun. Z içinde de en sonda kalan 100den az sayıdaki _ işaretlerini uçurulım. Bukkadar basit!!</p>

<p>Aşama aşama bakalım: Hücre içeriği şu olsun:<strong>Batı karedeniz bölge müdürlüğü</strong>:İlk 2 kelimeyi seçeceğiz. Boşlukların 100 adet _'e çevrildiği metinden 2*100=200 karakteri soldan kesip ilerleyeceğiz</p>
<pre class="formul">=SUBSTITUTE(SUBSTITUTE(LEFT(SUBSTITUTE(A1;" ";REPT("_";100));200);REPT("_";100);" ");"_";"")</pre>
<pre white-space="pre-wrap">
X=Batı____________________________________________________________________________________________________karedeniz____________________________________________________________________________________________________bölge____________________________________________________________________________________________________müdürlüğü
Y=Batı____________________________________________________________________________________________________karedeniz_______________________________________________________________________________________
Z=Batı karedeniz_______________________________________________________________________________________ //bunda 87 _ işareti var
Sonuç=Batı karadeniz
</pre>

<p>Farkettiyseniz bu sefer formül iyice komplike oldu. Bu noktada gerçekten UDF kullanmak en iyi çözüm olacaktır. Aşağıda gördüğünüz gibi oldukça basit bir kullanımı var. Daha önce belirttiğim gibi Microsft bu tür fonksiyonları neden hala metin fonksiyonları içine almıyor, anlamış değilim. Neyse ki UDF teknolojisi var.</p>

<pre class="formul">=ilknkelime(A1;2)</pre>

<p>İlgili UDF de aşağıdaki gibi yazılabilir. Detaylarını burada girmiyorum, gerek VBA temellerini, gerek dizileri gerek fonksyion konusunu bilmeniz gerekiyor. Bu detayları ilgili VBA safyalarında bulabilirsiniz.</p>

<pre class="brush: vb">
Function ilknkelime(hucre As Range, kaç As Byte, Optional ayrac As String = " ")
    'normal bir cümlede ayrac boşluk olacğaı için ayracı girmene gerek yok, zaten default olarak " " atadım.
    'ama atıyroum içeriği / ile ayrılmış bir hücre varsa 3.parametreyi / olarak girersin
    Dim kelimeler As Variant
    Dim i As Byte
    
    kelimeler = Split(hucre.Value2, ayrac)
    
    For i = 0 To kaç - 1
        geçici = geçici & kelimeler(i) & ayrac
    Next i
    
    ilknkelime = Mid(geçici, 1, Len(geçici) - 1)
    
End Function
</pre>

</div>




<h4 class=baslik>Bir hücredeki metinden sondan x kelime almak</h4>
<div>
<p>Bir önceki örneğin neredeyse aynısı mantıkta hazırlanır. Sadece soldan n*100 yerine sağdan n*100 _ işareti alınır. </p>

<p>Aşama aşama bakalım: Hücre içeriği yine aynı olsun:<strong>Batı karedeniz bölge müdürlüğü</strong>:Sondan 2 kelimeyi seçeceğiz. Boşlukların 100 adet _'e çevrildiği metinden 1*100=100 karakteri sağdan kesip ilerleyeceğiz</p>
<pre class="formul">=SUBSTITUTE(SUBSTITUTE(RIGHT(SUBSTITUTE(A1;" ";REPT("_";100));200);REPT("_";100);" ");"_";"")</pre>
<pre white-space="pre-wrap">
X=Batı____________________________________________________________________________________________________karedeniz____________________________________________________________________________________________________bölge____________________________________________________________________________________________________müdürlüğü
Y=____________________________________________________________________________________________bölge____________________________________________________________________________________________________müdürlüğü
Z=____________________________________________________________________________________________bölge müdürlüğü
Sonuç=bölge müdürlüğü
</pre>

</div>


<h4 class=baslik>Bir hücredeki metinden sondan x kelime hariç almak</h4>
<div>
<p>Bu örnekte formül biraz daha uzayackatır. En azından benim aklıma gelen yöntem bu olmuştur. Daha kısa yazılabilir mi bilmiyorum, üzerine eğilmek gerekir ancak kesinlikle UDF kullanmak en akıl karı iş olacaktır. Biz yine de Excel içindeki yerleşik fonksiyonlarla halletmeye çalışalım.</p>

<p>İzlenecek yol: Önceki yöntemlere benzer olarak, hücredeki bütün boşlukları 100 adet _ ile değiştiriyoruz. Buna X diyelim. X'in toplam uzunluğuna U diyelim. U'dan n*100 eksiği olan noktadan itibarenki karakterleri uçuralım, bu Y olsun. Y'nin içindeki 100 uzunluktaki _ işaretlerini uçuralım, bu Z olsun. Son olarak da kalan tüm _ işaretlerini uçuralım.</p>

<p>Aşama aşama bakalım: Hücre içeriği yine aynı olsun:<strong>Batı karedeniz bölge müdürlüğü</strong></p>
<pre class="formul">=SUBSTITUTE(SUBSTITUTE(REPLACE(SUBSTITUTE(A1;" ";REPT("_";100));LEN(SUBSTITUTE(A1;" ";REPT("_";100)))-200+1;200;"");REPT("_";100);" ");"_";"")</pre>
<pre white-space="pre-wrap">
X=Batı____________________________________________________________________________________________________karedeniz____________________________________________________________________________________________________bölge____________________________________________________________________________________________________müdürlüğü
U=321
Y=Batı____________________________________________________________________________________________________karedeniz________
Z=Batı karedeniz________
Sonuç=Batı karedeniz
</pre>

<p>Yukarda belirttiğim gibi formül daha da komplike oldu. Bunun için yazacağımız UDF ise aşağıdaki gibi olacaktır.</p>

<pre class="formul">=sonxkelimehariç(A1;2)</pre>

<p>İlgili UDF de aşağıdaki gibi yazılabilir. Detaylarını burada girmiyorum, gerek VBA temellerini, gerek dizileri gerek fonksiyon konusunu bilmeniz gerekiyor. Bu detayları ilgili VBA safyalarında bulabilirsiniz.</p>

<pre class="brush: vb">
Function sonxkelimehariç(hucre As Range, kaç As Byte, Optional ayrac As String = " ")
    'normal bir cümlede ayrac boşluk olacğaı için ayracı girmene gerek yok, zaten default olarak " " atadım.
    'ama atıyroum içeriği / ile ayrılmış bir hücre varsa 3.parametreyi / olarak girersin
    Dim kelimeler As Variant
    Dim i As Byte
    
    kelimeler = Split(hucre.Value2, ayrac)
    
    For i = 0 To UBound(kelimeler) - kaç
        geçici = geçici & kelimeler(i) & ayrac
    Next i
    
    sonxkelimehariç = Mid(geçici, 1, Len(geçici) - 1)
        
End Function
</pre>

</div>



<h4 class=baslik>Bir hücredeki belirli bir karakterden kaç tane geçtiğini bulmak</h4>
<div>
<p>"-" karakterini sayacağız</p>

<p>İzlenecek yol:Hücrenin uzunluğu X olsun. Hücredeki tüm "-" karakterlerini "" ile yokedelim ve kalan metnin uzunluğunu ölçelim, bu da Y olsun. Aradığımız cevap X-Y'dir</p>

<pre class="formul">
X=LEN(A1)
Y=LEN(SUBSTITUTE(A1;"-";""))
Çözüm=X-Y

=LEN(A1)-LEN(SUBSTITUTE(A1;"-";""))
</pre>
</div>



<h4 class=baslik>Hem 1 hem 2 isimli kişilerin olduğu listelerle uğraşmak</h4>
<div>
<p>Diyelim ki, 2 isimli kişilerin ikinci ismini, 1 isimlilerin ise doğal olarak ilk ismini alacaksınız. </p>
<p>İzlenecek yol: Hücredeki kelime sayısına bakıp yukardaki formüllerden birini kullanmak. Kelime sayısı 2 ise ilk kelimeyi al, kelime sayısı 3 ise 2.kelimeyi al. Bunu tek formül yazmak çok uzun olabilir, IF'lı bir yapı olacağı için Filling de yapılamaz, en güzel çözüm alsında UDF yaratmaktır ancak biz burada formülümüzü yazacağız. Fakat yine de formülü tek seferde yazmak yerine iki ayrı hücreye de yazabilirisiniz. İlk hücrede kelime sayısı olur. İkinci hücrede de bu sayıyı kontrol eden bir formül</p>

<pre class="formul">İlk formülümüz B1de olsun:
=LEN(A1)-LEN(SUBSTITUTE(A1;" ";""))+1
</pre>
<p></p>
<pre class="formul">İkinci formülümüz C1'de olsun:
=IF(B1=2;LEFT(A1;FIND(" ";A1)-1);TRIM(MID(SUBSTITUTE(A1;" ";REPT(" ";100));100;100)))
</pre>
</div>



</div>
</asp:Content>

