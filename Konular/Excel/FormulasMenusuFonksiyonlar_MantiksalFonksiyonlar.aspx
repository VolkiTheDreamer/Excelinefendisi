<%@ Page Title='FormulasMenusu1 MantiksalFormuller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Formulas Menüsü(Fonksiyonlar)'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>


<!--********************************************************************************************************************************************-->
<h1>Mantıksal Fonksiyonlar</h1>
<p>Mantıksal fonksiyonlar, ya sonucu ya da içeriği TRUE/FALSE olan fonksiyonlardır.</p>
<p>Bunlar öyle komplike fonksiyonlar değillerdir. Zaten tak başlarına kullanımları olmayıp birçok formülün içine girerler. Şimdi kısa kısa bunları inceleyelim.</p>

<h2 class="baslik">IF ve IF türevleri</h2>
<div class="konu">
<h3>IF</h3>
<p>IF ile bir koşul sağlandığında veya sağlanmadığında hangi sonucu göstereceğimizi belirleriz. Genel Syntax'ı <span class=" keywordler">IF(kontrol;DOĞUYSA ŞU, YANLIŞSA ŞU)</span> şeklindedir.</p>

<p>Örneğin, A kolonunda kişilerin yaşları var ve bu yaşlara göre kişilerin reşit olup olmadığını B kolonuna yazdıralım.</p>

<pre class="formul">=IF(A2>=18;"Reşit";"Reşit Değil")</pre>

<h4>İçiçe IF</h4>

<p>IF'in bir de içiçe kullanımı vardır ve içiçe en fazla 64 IF yazabilirsiniz(Excel 2007den önce bu sayı 7 idi). Ancak herhalde 8-10 IF dışında çok fazla IF'e ihtiyacınız olmayacaktır. Gerçekten ihtiyaç varsa, geçici bir yere Lookup tablosu yapıp orada Vlookup formülünü kullanmak daha pratik olacaktır. <a href="FormulasMenusuFonksiyonlar_LookupAramaFonksiyonlari.aspx#iciceif">Burada</a> bunu nasıl yapacağınızı bulabilirsiniz.</p>

<pre class="formul">=IF(A2>=65;"Yaşlı";IF(A2>=40;"Orta Yaşlı";IF(A2>18;"Genç";"Çocuk")))</pre>

<p><img src="/images/exceliciceif.jpg"></p>
	<p>İçiçe IF yazarken dikkat edilmesi gereken bir husus var. Excel'de Between diye bir operatör olmadığı için bu amacı < ve > işaretlerini AND operatörü ile birlikte kullanarak veya sorgulamaya önce en büyük değerden başlayıp aşağı doğru inerek veya < işaretini kullanıp yukarı doğru çıkarak yapabiliriz. Bu yukardaki örnekte ben önce en büyük değerden başlayıp aşağıda doğru inme yolunu seçtim. </p>

<p>Yine farkettiyseniz son değer olan Çocuk için ayrı bi IF yazmadım, zira bu, önceki şartlar sağlanmadığında otomatik olarak yazdırılacak kısımdır.</p>

<p>Bir diğer önerim de şu olacaktır <strong>ALT+Enter</strong> tuş kombinasyonuna basarak kodunuzun okunurluğunu artırabilirsiniz</p>

<pre class="formul">=IF(A2>=65;"Yaşlı";
                         IF(A2>=40;"Orta Yaşlı";
                                                IF(A2>18;"Genç";
                                                                   "Çocuk")))</pre>

<p>Son bir önerim daha var, o da, içiçe IF'leri yazarken her IF'ten sonra açtığınız parantezi önce kapatın, sonra ilk olarak True kısmın formülünü yazın, sonra da kalan kısma geçici olarak FALSE yazın. Yukarıdaki örnek üzerinden gidecek olursak, önce şöyle yazın, IF(), sonra IF(A2>=60;"Yaşlı";FALSE), sonra FALSE'ı silip içini dolurun, ve son FALSE'a kadar bunu yapın, yani sırayla şöyle olsun: </p>

<pre class="formul">IF()
=IF(A2>=60;"Yaşlı";FALSE)
=IF(A2>=60;"Yaşlı";IF())
=IF(A2>=60;"Yaşlı";IF(A2>=40;"Orta Yaşlı";FALSE))
=IF(A2>=60;"Yaşlı";IF(A2>=40;"Orta Yaşlı";IF()))
=IF(A2>=60;"Yaşlı";IF(A2>=40;"Orta Yaşlı";IF(A2>=18;"Genç";FALSE)))
=IF(A2>=60;"Yaşlı";IF(A2>=40;"Orta Yaşlı";IF(A2>=18;"Genç";"Çocuk")))
</pre>

<h3>IFS(2016)</h3>
<p>IFS ile yukardaki gibi içiçe birsürü IF dizmekten kurtulmuş oluyorsunuz.</p>

<p>Syntax'ı şöyledir: </p>

<pre class="formul">=IFS(Sorgu1, Sorgu1DoğruysaDeğer1, Sorgu2, Sorgu2DoğruysaDeğer2,.......Sorgu127, Sorgu127DoğruysaDeğer127)</pre>

<p>Hiçbir koşul TRUE döndürmezse N/A görünür. O yüzden en son koşulu <strong>1=1;"hiçbirkoşul sağlanmadı"</strong> gibi bir çift ile bitirmekte fayda var, veya <strong>TRUE;"hiçbirkoşul sağlanmadı"</strong> </p>

<p>Bir örnek verelim.</p>

<pre class="formul">=IFS(A2>65;"Yaşlı";A2>40;"Orta Yaşlı";A2>18;"Genç";TRUE;"Çocuk")</pre>

<p>Bunu şu şekilde yapmak belki biraz daha uygun olur, zira olur da, hatalı bir giriş varsa bunları da ele almış olursunuz. Mesela yaş bilgisi boş ise, negatif ise veya  sayısal olmayan bir değer ise uyarı mesajı çıkar.</p>
<pre class="formul">=IFS(A2>65;"Yaşlı";A2>40;"Orta Yaşlı";A2>18;"Genç";A2>0;"Çocuk";TRUE;"Uygun bir yaş değildir")</pre>

<p>IFS fonksiyonu her ne kadar 127 adet sorgulama yapma imkanı verse de, çok adetli sorgulamalarda yukarda bahsettiğim Lookup yöntemini kullanmak daha mantıklı olacaktır.</p>


<h3>IFERROR</h3>
<p>IFERROR'u, yazdığımız bir formül hatalı sonuç döndürdüğünde hata değil de başka birşey görünsün istediğimizde kullanırız. Mesela bi Vlookup işlemi yaptığımız ve aradığımız değeri bulamadığımızda normalde N/A döndürür, ama bi de en altta diptoplam aldırdığımızı düşünün, o kolonda bi N/A varsa diptoplam da hata çıkacaktır, o yüzden hata durumunda N/A yerine 0 görünmesini isteyebiliriz. Ör:</p>

<pre class="formul">=IFERROR(VLOOKUP(A2;D:E;2;0);0)</pre>

<p>Yine normal IF'te olduğuğ gibi bunu da aşama aşama aşağıdaki gibi yazarız</p>

<pre class="formul">=IFERROR(TRUE;FALSE)
=IFERROR(VLOOKUP();0)
=IFERROR(VLOOKUP(A2;D:E;2;0);0)
</pre>

<p>NOT:2007 öncesi günlerde bunu aşağıdaki gibi uzun bir yolla yapardık.</p>
<pre class="formul">
IF(ISERROR(VLOOKUP(A2;D:E;2;0));0;VLOOKUP(A2;D:E;2;0))
</pre>



<h3>IFNA(2013)</h3>

<p>IFERROR'dan farklı olarak IFNA ile, sadece N/A hatası oluştuğunda başka bir değer gösterilir, diğer hata sonuçları olduğu gibi kalır.</p>

<p>Şimdi diyelim ki yine Vlookup yapacaksınız, sadece N/A'ların yani lookup yaptığınız yerde bulunmayan değerlerin gelmemesini istiyorsunuz. Diğer hatalar gelsinki, o hatalara özgü duruma müdahale edesiniz. Mesela Lookup listesinde kişilerin hedef gerçekleştirme oranlarını(HGO) getirmek istiyorsunuz, kişin hedefi 0 ise <abbr title="Gerçekleşen/Hedef şeklinde hesaplanır">HGO</abbr> sıfıra bölme hatası olan #DIV/0! verir. Bunu sıfırlamak demek, bu hatayı baskılamak demektir, halbuki siz N/A dışındaki hataları baskılamak yerine onları düzeltmek istersiniz. Mesela 0 hedefli kişiler kimse bunları alıp hedefleme ekibinden bunlara hedef vermesini isteyebilirsiniz.
</p>

</div>

<!--********************************************************************************************************************************************-->

<h2 class="baslik">AND/OR/NOT</h2>
<div class="konu">

<h3>AND</h3>
<p>Bu fonksiyon sıklıkla IF'li formüllerde kullanılır. Aynı anda birden fazla(255e kadar) koşulun doğru olması durumunda TRUE kısmın değeri döner, koşullardan biri sağlanmazsa FALSE kısmın değeri döner. Mesela bir hücreye hem A1 hücresinin 1 olduğu hem de C2'nin boş olduğu durumda 100, aksi halde 0 yazdırmak istersek şu formülü yazarız.</p>

<pre class="formul">=IF(AND(A1=1;ISBLANK(C2));100;0)</pre>

<p>Aşağıda ise daha uzun bir formül örneği görüyoruz. Burada belki dikkatinizi çekmiştir. Formül içinde <span class=" keywordler">Named Range</span>(özel isimli hücre) kullanmışım. Mesela <strong>bugünayno</strong> diye bir Name var, bunun içeriği şöyle: <strong>=DAY(TODAY())</strong>, yani ayın kaçıncı günü olduğunu veriyor. Bir diğer Name olan hftno var, onun da içeriği şu:<strong>=WEEKDAY(TODAY();2)</strong>, yani bugün haftanın kaçıncı günü onu veriyor. Bu tür Name'ler kullanmak hem formülümüzü daha kısaltır hem de daha anlaşılır hale getirir. Formüle geri dönecek olursam, önce ne yapmaya çalıştığımı belirteyim. Bu, erken kredi kapanmalarını öngören bir formülün bir parçası ve kendisi de başka bir hücredeki Vlookup formülünün ilk paramteresi olarak iş görüyüor. Bunun da şöyle bir hikayesi var:Ayın belli günlerinde kredi kapanmaları daha yüksek oluyor. Aşağıdaki tabloda temisili rakamları görüyorsunuz. Şimdi görüldüğü gibi ayın 15'inde kamu kurumları banka aracılığı ile maaş ödemelerini yaparlar ve çalışanların eline para geçtiği için de bunların bazısı kredilerini erken kapatır, yani vadesini beklemez. Ayrıca ayın ilk ve son günleri de özel kurumlar maaş ödmesi yaptığı için bu günlerde de ortalamaya göre bir miktar daha fazla kapama olur. Gelelim formüle: Bazen ayın 15'i veya 1'i haftasonuna denk gelebilir ve maaş ödemesi 1 veya 2 gün sarkabilir. O yüzden diyoruz ki, <strong>Eğer Bugün ayın 2 veya 3'ü ise ve aynı zamanda Pazartesi ise 1 yazsın yani bugünü ayın 1'i gibi düşünelim, bu şart sağlanmazsa(Ör:Çarşamba olup ayın 4'üdür, Pazartesi olabilir ama ayın 5idir, veya ayın 2sidir ama Salıdır) şu şarta bakalım:Bugün ayın 16 ve 17si ise ve aynı zamanda Pazartesi ise 15i gibi düşünelim, bu şart da sağlanmazsa bugün 28/29 Şubatsa 31'i gibi düşünelim, diğer tüm durumlarda bugün ayın kaçıysa onu yazalım</strong></p>

<pre class="formul">=IF(AND(bugünayno<=3;hftno=1);1;IF(AND(bugünayno>15;bugünayno<18;hftno=1);15;IF(AND(MONTH(TODAY())=2;bugünayno>=28);31;bugünayno))) </pre>


<p><img src="/images/andorkredi.jpg" height="10%" width="10%" class="zoomla"></p>
<p>Bir diğer kullanım alanı da IF'siz olup, sonucun TRUE/FALSE olmasını sağlar, 
ve mesala sadece TRUE olanları filtrelemek isteyebilirsiniz.</p>
<p><img src="/images/IFsizAnd.jpg"></p>
<h3>OR</h3>
<p>OR'un kullanımı da AND'e çok benzer. Bunda koşullardan herhangi biri doğruysa TRUE döndürür, hiçbiri sağlanmıyorsa FALSE döndürür. Aşağıda küçük bir örnek bulunmkatadır. A2 bugünden daha büyükse veya B2 0 ise TRUE, diğer durumlarda FALSE döner.</p> 

<pre class="formul">=OR(A2>TODAY();B2=0)</pre>

<h3>NOT</h3>
<p>TRUE/FALSE döndüren bir formülün sonucunu tersine döndürmek için kullanılır. Genel kullanım yeri, günlük konuşma diline yakın bir anlam çıkarması sayesinde bir hücrenin DOLUMU olduğunu göstermek amacıyla ISBLANK iledir. </p>
<pre class="formul">=NOT(ISBLANK(A2))</pre>
<p>İstediğimiz şeyi, tabiki ISBLANK yazıp FALSE olanları filtrelemek şeklinde de yapabilirdiniz ama dediğim gibi konuşma diline yakın olması adına bu şekilde kullanımı daha uygundur. Bu arada bu sorgulama şeklini IF(A2<>"";TRUE;FALSE) veya IF(LEN(A2)>0;TRUE;FALSE) şeklinde de yapabilirdik. Excel'de bir sonuca ulaşmanın birden çok şekli olduğunu hep aklınızda bulundurun.</p>

</div>

</asp:Content>
