<%@ Page Title='Dosyaislemleri Dosyaokumaveyazma' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
	<div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Dosya işlemleri'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>
	<h1>Dosya Okuma ve Yazma işlemleri</h1>
		<h2 class="baslik">Giriş</h2>
		<div class="konu">
	<p> İçine veri yazılan birçok dosya tipi bulunmaktadır, ancak biz VBA 
	konsepti içinde genellikle text(csv dahil) dosyalarıyla çalışacağız. SQL 
	kodlarını depoladığınız ".sql" uzantılı dosyaları da bu kapsamda 
	düşünebilirsiniz. Excel 
	dosyalarla çalışmayı zaten <a href="DortTemelNesne_Workbook.aspx">Workbook</a> 
	ve <a href="DortTemelNesne_Application.aspx#filefolder">Application</a> 
	bölümlerinde görmüştük. </p>
		<p> Text dosyaları bilgi depolayıp okumanın kolay 
		bir yolunu sunarlar. Özellikle settings(ayar) bilgileri veya aşamalı bir 
	sürecin durum bilgilerini(log) okumakta/yazmakta oldukça kullanışlıdırlar. </p>
		<p> VBA'de iki tür okuma/yazma yöntemi bulunuyor. Öncelikle biz VB6'dan 
		miras gelen klasik okuma yazma yöntemine bakacağız.</p>
		<p> <strong>UYARI</strong>:Buradan itibaren aşağıda göreceğiniz tüm dosya 
		okuma işlemlerinde, dosya okuma hep sağa ve aşağı yönlüdür. Döngüsel 
		işlemlerde "bir sağ kolona/karaktere veya bir alt satıra geç" tarzında 
		ilave bir kod ifadesi yoktur. Bu işlem otomatik olmaktadır.</p>
            </div>

		<h2 class="baslik">Klasik Okuma/Yazma işlemleri</h2>
		<div class="konu">
		<h3> Dosya açma ve Kapama</h3>
		<h4> Dosya Açma</h4>
			<p> Okuma işlemi için de de Yazma işlemi için de öncelikle dosyanın 
		açılması gerekir. Bunun için <span class="keywordler">Open</span> 
		fonksiyonu kullanılır. Aşağıdaki gibi bir syntax'a sahiptir.</p>
	<p><strong>Open <em>dosyayolu</em> For <em>mod</em> [<em>Erişim tipi</em>] [lock] 
	As Dosyano</strong></p>
		<ul>
			<li><strong>Dosyayolu</strong>:Dosyanın bulunduğu tam adres. 
			Ör:C:\deneme\deneme.txt</li>
			<li><strong>Mod</strong>:<strong>Input </strong>ise okuma, <strong>
			Output </strong>ise yazma, <strong>Append </strong>ise dosya sonuna ekleme 
			yapılır. 2 tane daha var ama bize bu üçü yeter. Output seçildiğinde 
			mevcut dosya varsa ezilip içeriği yeniden oluşturulur, olmayan bir 
			dosya girildiyse bu dosya yaratılır.</li>
			<li><strong>Erişim tipi ve Lock tipi</strong>:Opsiyoneldirler. Dosya 
			açıkken, başkalarının ne yapabileceğini gösterir. Biz bunları 
			kullanmayacağız, o yüzden default değerleri devreye girecek.</li>
		</ul>
		<h4>Freefile</h4>
		<p>Dosyalar açıldığında onlara bir sıra numarası verilir. Bu numara 
		manuel belirtilebileceği gibi, çok sayıda okuma yazma yapılan bir prosedür 
		içinde o andaki müsait sıra numarasını veren <strong>Freefile</strong> 
		deyimi de kullanılabilir. Manuel giriş için sıra numarası # ile 
		kullanılır. #1 gibi.</p>
			<h4>Dosyayı Kapama</h4>
			<p>Dosyayı Close ifadesi ile kaparız, ancak parametre olarak dosya 
			adresi değil, numarasını alır.</p>
			<p>Dosyayı kapatmazsak, tekrar aynı dosyayı açmaya çalıştığımızda 
			"Dosya zaten açık" hatası alırız.</p>
			<h4>Kod</h4>
			<p>Bu durumda, örnek bir dosya açma kodu şağıdaki gibi olacaktır.</p>
			<pre class="brush:vb">adres = "C:\Users\Volkan\Desktop\denemeler\dosya1.txt"
Open adres For Input As #1 'veya 1 veya FreeFile
'çeşili işlemler

Close #1</pre>
		<h3>Dosyadan veri okuma</h3>
			<h4>Kolon kolon bilgi okumak</h4>
		<p>Aşağıdaki bilgileri içeren bir text dosyamız olsun. Kişinin adı, yaşı 
		ve baba adı bilgileri var.<br><br>volkan ,38 ,ismail<br>ayşe ,40 ,murat<br>serkan ,35 ,osman</p>
		<p>Buradan ilk kaydın yaş bilgisini almak istiyoruz diyelim. Bunun için
		<span class="keywordler">Input </span>deyimini kullanıp, istediğimiz 
		kolon sayısı kadar değişken belirleyip kolon bilgilerini bu değişkenlere 
		atıyoruz.</p>
		<pre class="brush:vb">Sub teksatırdan_tekkolon_okuma()
adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt

Open adres For Input As 1
Input #1, adı, yaş 'kolon sayısı kadar paremetre alır. # zorunlu
Debug.Print yaş '38 verir
'ikinci bi Debug.Print yaş yazsak bile yine 38 yazar. sadece tek satır okumasu var.
Close 1

End Sub</pre>
		<p>Bu kod tabiki sadece ilk satır için bilgi döndürür. Satırda ilerleme 
		yapamıyoruz. Özellikle her defasında üzerine yazma yapılan tek satırlık 
		bilgi içeren dosyalarda kullanışlıdır. Mesela ikinci kolonunda, bilginin 
		yazdırıldığı tarihi veya kişiyi gösteren bir dosyadan bu tarihi veya 
		kişiyi elde etmek 
		isteyebiliriz. Böylece bu dosyaya en son ne zaman bilgi yazıldığını veya 
		kimin tarafından yazıldığını elde edebiliriz.</p>
		<h4>Tek satırdan kısmi bilgi okuma</h4>
			<p>Input ifadesini <strong>Input(karakteradedi, dosyano)</strong> 
			şeklinde kullandığımızda belirli adette karakter okumuş oluruz.</p>
			<pre class="brush:vb">Sub teksatır_kısmen_okuma()
adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open adres For Input As 1
x = Input(6, 1) '1 nolu dosyadan 6 karakter oku
Debug.Print x 'volkan
End Sub</pre>
			<h4>İlk satırın tamamını okuma</h4>
		<p>Yine yukarıdaki dosyamı elimizde bulunsun. Bu sefer ilk satırın 
		tamamını elde edeceğiz. Bunu <span class="keywordler">Line Input</span> 
		deyimi ile yapıyor ve içeriği ikinci parametredeki değişkene atıyoruz .</p>
		<pre class="brush:vb">Sub teksatır_tamamını_okuma()

adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open adres For Input As 1 
Line Input #1, metin 'ilk satırı oku ve metin değişkeninde depola
Debug.Print metin 'volkan,38,ismail yazar

Close 1

End Sub</pre>
		<h4>x adet satırı tek tek okuma</h4>
		<p>Bu işlem Line Input'un bir For Next döngüsü ile kullanımı ile yapılabilir. 
		Belli sayıda satır bilgisinin yeterli olduğu durumlarda kullanılır.</p>
		<pre class="brush:vb">Sub x_adet_satır_oku()

adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open adres For Input As 1
For i=1 to 2
  Line Input #1, metin
  Debug.Print metin
Next i

Close 1

End Sub</pre>
<h4> Dosyadaki tüm metni okuma</h4>
		<h5> 1.Yöntem:Dosyadaki karakter sayısı kadar okumak</h5>
		<p> Bunu 
		<span class="keywordler">LOF</span> ifadesi ile yapıyoruz. Bu, Length Of File'ın 
		kısaltılmışıdır, yani dosyadaki karakter sayısını verir. Input ile 
		birleştirerek de dosyadaki tüm karakter sayısını oku demiş oluruz. 
		"volkan naber" şeklinde 12 karakterli bir metni içeren bir dosyada;</p>
		<p> LOF(1):12 döndürür</br>
Input(LOF(1),1):1 nolu dosyayı tamamen okur:"volkan 
		naber"</p>
		<p> Örnek bir kodumuz ise şöyledir. Bu kodda ayrıca
		<span class="keywordler">Seek</span> ifadesini de kullandık. Bununla 
		dosyada belirli bir sıradaki karaktere konumlanıyoruz, ki bunu genelde 
		belirli bir sırayla ilerlediten sonra tekrar ilk karaktere dönmek için 
		kullanırız. Syntax: <strong>Seek dosyano, konum</strong></p>
	<pre class="brush:vb">
Sub DosyaOkuTümü1()

adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"
Open adres For Input As 1

Debug.Print LOF(1) 'Length Of File Of File

içerik = Input(LOF(1), 1) '1 nolu dosyanın hepsini oku
Debug.Print içerik

Seek 1, 1 '1 nolu dosyanın 1.karakterine yan en başa git
içerik = Input(LOF(1) - 10, 1) '1 nolu dosyanın son 10 karekteri hariç oku
Debug.Print içSeek 1, 5 '1 nolu dosyanın 5.karakterine git
içerik = Input(10, 1) '1 nolu dosyanın 5.karakterinden sonraki ilk 10 karekterini okurini oku
Debug.Print içerik

Close 1

End Sub</pre>
	<p> Aynı mantıkla bir şekilde dosyanın ilk x karekterini okumak için içerik 
	değişkenine <strong>Input(x,1);</strong> son x karekteri hariç okuma yapmak 
	isterseniz içerik değişkenine <strong>Input(LOF(1)-x,1)</strong> şeklinde 
	atama yaparsınız.</p>
		<p> Bu yöntemde dosya içindeki metnin kaç satırda yer aldığı önemli 
		değildir. Tüm metin tek bir değişkende depolanır.</p>
		<h5> 2.Yöntem:Dosya sonuna kadar satır satır okumak</h5>
		<p> Yukarıda gördüğümüz belli sayıdaki satırları tek tek okumadan farklı 
		olarak tüm satırları tek tek okuyoruz. Satırların bittiğini
		<span class="keywordler">EOF</span>(End Of File'ın kısaltması )özelliği 
		ile anlıyoruz.</p>
		<pre class="brush:vb">
Sub DosyaOkuTümü2()
tamadres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open tamadres For Input As 1

Line Input #1, satırmetni
içerik = satırmetni

Do Until EOF(1)
  Line Input #1, satırmetni
  içerik = içerik &amp; vbNewLine &&amp; satırmetni
Loop

Debug.Print içerik

Close 1
End Sub
</pre>
		<p>
		NOT: İçerik değişkenini oluşturuken kayıtlar arasınsa <strong>vbNewLine
		</strong>koyarak 
		satırbaşı yapıyoruz. Ancak ilk başta değişkenin içi boş olacağı için 
		fazladan bir boş satır oluşmaması için en başta bir kezliğine döngüye 
		girmeden ilk satırın atamsını yapıyorum sonrasında döngü içinde <strong>vbNewLine</strong> ekliyorum.</p>
		<p>
		Bu arada istenirse ilgili metinler <strong>vbNewLine </strong>denmeden satır satır değil 
		de ardışık bir şekilde de biraraya getirilebilir.</p>
		<h5>Tüm içeriği bir diziye aktarmak</h5>
		<p>Bunu da kendi içinde iki ayrı yöntemle yapabiliriz. İlk yöntemde 
		satır satır okur ve her satırı bir collectiona atarız. Özellikle her 
		satırın başına/sonuna başka bir metin eklemek gereken durumlarda bunu 
		kullanabiliriz. İkinci yöntemde ise tüm metni okuyup Enterları(vbCrLf veya vbNewLine) Split ederek 
		diziye atayabilirsiniz.</p>
		<pre class="brush:vb">
'1)Collectiona atama
Do Until EOF(1)
  Line Input #1, metin
  coll.Add metin
Loop

'2)Diziye atama
içerik=Input(LOF(1), 1)
dizi = Split(içerik, vbCrLf)</pre>
<p>Bu iki yöntemle de yukardaki örneği yapalım.</p>

<pre class="brush:vb">
Sub dosyaoku3()

adres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open adres For Input As 1

Debug.Print "------önce dizi yöntemi-----"
içerik = Input(LOF(1), 1)
dizi = Split(içerik, vbCrLf)

For Each satır In dizi
    Debug.Print "Prefix:" + satır + " -Suffix"
Next satır

'veya
Debug.Print "------collection yöntemi-----"
Seek 1, 1
Dim coll As New Collection
Do Until EOF(1)
  Line Input #1, metin
  coll.Add metin
Loop

For Each satır In coll
    Debug.Print "Prefix:" + satır + " -Suffix"
Next satır

Close 1

End Sub
</pre>
		<h5>Virgülle ayrılmış metinleri hücrelere yazdırmak</h5>
		<p>Virgülle ayrılmış tüm değerleri farklı kolonda olacak şekilde satır 
		satır Excele yazdırmak isteyebilirsiniz. Bunun için satır satır okuma 
		yapmamız gerekir ve her satırı Split ile bir diziye atayabilir, sonra da ilgili hücrelere bu 
		dizi elemanlarını döngüsel şekilde yazdırabiliriz.</p>
		<pre class="brush:vb">
Sub DosyaOkuTümü4()
tamadres = "C:\Users\volkan\Desktop\denemeler\dosya1.txt"

Open tamadres For Input As 1

i = i + 1
Do Until EOF(1)
  Line Input #1, x
  dizi = Split(x, ",")
  For j = 0 To UBound(dizi)
    Cells(i, j + 1) = dizi(j)
  Next j
  i = i + 1
Loop

Close 1
End Sub</pre>
		<h3>Dosyaya veri yazma</h3>
		<h4>Sıfırdan yazma</h4>
		<p>
	Dosyaya veri yazdırmak için dosyayı <strong>Output </strong>modunda açmamız 
		gerekir. Yazdırma eylemi için iki fonksiyonumuz var.
		<span class="keywordler">Write</span><strong> </strong>ve
		<span class="keywordler">Print</span>. Write, yazdırılan ifadeyi " " içine alarak yazarken 
	Print böyle bir işlem yapmaz.</p>
		<p>
		Her iki deyim de ardışık kullanımlarında satır satır yazdırır. Yani 
		birbirini takip eden metinler şeklinde yazılmaz. Örneğin;</p>
		<pre class="brush:vb">ad="Volkan"
soyad="Yurtseven"

Print ad
Print soyad

'bu kodun çıktısı aşağıdaki gibidir
'Volkan
'Yurtseven
'VolkanYurtseven değil</pre>
		<p>
		Eğer yazdırılan metinlerin ardışık yazdırılması isteniyorsa aşağıda 
		anlattığım <strong>TextStream </strong>nesnesini kullanmanız gerekir.</p>
		<p>
		Şimdi veri yazdırmaya ait küçük bir kod yazalım.</p>
	<pre class="brush:vb">
Sub DosyaYazTekSatır()

tamadres = "C:\Users\volkan\Desktop\denemeler\dosya2.txt
ff = FreeFile ' o an uygun olan dosya numarası verilir. 

Open tamadres For Output As ff 

Print #ff, "bu ilk satır"
Print #ff, "bu da ikinci satır"

Close ff 'kapatırken kaydeder, ayrı bir save işlemi yoktur

End Sub</pre>
		<p>
		<strong>NOT</strong>:Eğer dosya mevcut değilse otomatikman oluşturulur, 
		varsa ezilir ve üzerine yazılır.</p>
		<h4> Varolan dosyaya ekleme yapmak</h4>
		<p> Varolan bir dosyanın sonuna ekleme yapmak istiyorsak bu işi <strong>Append</strong> 
		deyimi ile yaparız. Eğer dosya mevcut değilse 
		Outputta olduğu gibi otomatikman oluşturulur.</p>
	<pre class="brush:vb">
Sub DosyaAppend()

tamadres = "C:\Users\volkan\Desktop\denemeler\dosya2.txt
ff = FreeFile ' o an uygun olan dosya numarası verilir. 

Open tamadres For Append As ff 

Print #ff, "bu üçüncü satır"

Close ff 'kapatırken kaydeder, ayrı bir save işlemi yoktur

End Sub</pre>
		<h4> Excelden okuyup dosyaya yazma</h4>
		<p> 
		İçiçie iki For Next ile satır ve sütunlarda dolaşırız, araya "," veya 
		istediğimiz başka bir ayraç ekleriz.</p>
		<p> 
		Aşağıdaki örnekte 10 satır 3 kolondan oluşan bir listeyi metin dosyasına 
		yazdırıyoruz.</p>

	<pre class="brush:vb">
Sub Exceldenyaz()

tamadres = "C:\Users\volkan\Desktop\denemeler\dosya3.txt
ff = FreeFile

Open tamadres For Output As ff

For i = 1 To 10
    For j = 1 To 3
        If j = 3 Then
           satırmetin = satırmetin + Trim(Cells(i, j).Value)
        Else
           satırmetin = satırmetin + Trim(Cells(i, j).Value) + ","
        End If
    Next j
    
    Print #ff, satırmetin
    satırmetin = ""

Next i

Close ff

End Sub</pre>

		<h4>
		Metin değişikliği yapmak</h4>
		<p>
		Bazen tek bir dosyada bazense birçok metin dosyasında aynı anda bir 
		metin değişikliği yapmak isteriz. Birçok dosyayı elde etmeyi 
		<a href="Dosyaislemleri_DosyaveKlasorerisimi.aspx">bir önceki 
		sayfada </a>görmüştük. Bir döngü ile bu klasörleri/dosyaları elde ettikten 
		sonra yapmanız gereken iş 4 aşamadan oluşur:</p>
		<ul>
			<li>Önce dosyayı açmak</li>
			<li>Metni okumak</li>
			<li>Aradığımız metni bulup değiştirmek</li>
			<li>Dosyaya tekrar yazdırmak</li>
		</ul>
		<pre class="brush:vb">
Sub MetinDeğiştir()

Dim adres As String
Dim içerik As String

adres = "C:\deneme\deneme.txt"

Open adres For Input As 1
içerik = Input(LOF(1), 1)
Close 1
  
içerik = Replace(içerik, "abc123", "abc345")

Open adres For Output As 1
Print #1, içerik
Close 1

End Sub</pre>
		<p>
		Yapılan değişkliği illa dosyaya kaydetmek zorunda değilsiniz. Mesela 
		benim bazı SQL'leri tuttuğum metin dosyalarım var, içinde değişkenlerin 
		olduğu bölümler var. Bu dosyaları bir nevi şablon olarak tutuyorum, onların 
		üzerinde değişiklik yapmıyorum, onun yerine dosyayı okuyup bi değişkene 
		atıyorum ve onun üzerinde replace işlemi yapıp, SQL metni olarak işleme 
		sokuyorum. En aşağıdaki örnekler bölümünde bu işlemi görebilirsiniz.</p>
		<p>
		<span class="dikkat">UYARI/ÖNERİ:</span>Yapılandırılmış tipteki(belli format, uzunluk ve kolonlardan oluşan) büyük metin dosyalarında büyük çaplı 
		değişiklikler yapılacaksa bu dosyaya ADO ile bağlanıp Update işlemi 
		yapılması daha hızlı sonuç verecektir. I/O yöntemi ile replace işlemi 
		küçük dosyalarda tercih edilmelidir.<br>
</p>
</div>
		<h2 class="baslik">TextStream nesnesi</h2>
		<div class="konu">
		<p> Dosyalara yazma ve okumanın bir diğer yolu da 
		<span class="keywordler">TextStream</span> nesnesi 
		yoluyladır.&nbsp;Niye böyle bi yöntem daha var? Bu sınıf, aslında 
		web sayfalarında <strong>VBscript </strong>diliyle yazılmak üzere tasarlanmış bir sınıftı ama sonradan VBA içinde de kullanıma alındı. O yüzden 
		<strong>Scripting</strong> Runtime librarysi içindedir 
		ve&nbsp; references menüsünden eklenmesi gerekir.</p>
		<p> Ben şahsen hem okunurluk hem de kullanım kolaylığı açısından 
		TextStream nesnesini kullanmayı tercih ediyorum, ancak her zaman olduğu 
		gibi başkalarının yazdığı kodları okumanız/kullanmanız gerekebileceği 
		için her yöntemi bilmekte fayda var. Aşağıda göreceğiniz 
		üzere TextStream'in bazı ek özellik ve metodları da onu ayrıcalıklı 
		kılmaktadır.</p>
		<h3> Erişim &amp; Yaratım</h3>
		<p> Bir TextStream nesnesine erişmek için <strong>FSO'</strong>nun
		<span class="keywordler">CreateTextFile</span> veya 
		<span class="keywordler">OpenTextFile</span><strong> </strong>metodlarını kullanabileceğimiz gibi 
		<strong>File </strong>nesnesinin 
		<span class="keywordler">OpenAsTextStream</span> nesnesini de kullanabiliriz.</p>
		<h4> CreateTextFile</h4>
		<p> <strong>Syntax: </strong>
		fso.CreateTextFile(dosyadı,Owerrite?,Unicode desteği?)</p>
		<p> Aşağıdaki kod ile varolan bir dosyayı, eğer mevcutsa üzerine 
		yazdırarak(yani içini boşaltarak) Türkçe karakterleri de destekleyecek 
		şekilde açıyoruz.</p>
		<pre class="brush:vb">'global fso nesnesnin var olduğunu düşünerek ilerliyoruz
Dim ts As TextStream
Set ts = fso.CreateTextFile("c:\deneme\deneme.txt", True, True) </pre>
		<p> Eğer ikinci parametreyi False olarak kullanmak yani dosya 
		mevcutsa onu ezmeyelim istiyorsak, aşağıdaki gibi dosyanın varlığını kontrol ederek 
		açmalıyız yoksa "dosya zaten mevcut" hatası alırız.</p>
		<pre class="brush:vb">If Not fso.FileExists("C:\Users\Volkan\Desktop\denemeler\deneme1.txt") Then
   Set ts = fso.CreateTextFile("C:\Users\Volkan\Desktop\denemeler\deneme1.txt", False, True)
   ts.Write ("merhaba")
End If</pre>
		<h4> OpenTextFile</h4>
		<p> <strong>Syntax:</strong>fso.OpenTextFile(dosyadı,I/O 
		tipi,MevcutdeğilseYaratılsınmı?,Format)</p>
		<pre class="brush:vb">Set ts = fso.OpenTextFile("C:\deneme\deneme.txt",ForWriting,True,TristateFalse)</pre>
		<p>I/O tipi,  <span class="keywordler">ForWriting</span> açıldığında içerik ezilir. Bu CreateTextFile'ın ikinci 
		parametresinin True olarak açılmasıyla aynı etkidedir.</p>
		<p> <span class="keywordler">ForReading</span> ile okuma yaparsınız, yazmaya izin verilmez.</p>
		<p> <span class="keywordler">ForAppending</span> ile en sona konumlanır ve oraya yazarsınız, böylece 
		mevcut içerik silinmemiş olur.</p>
		<p> Üçüncü parametreyi, dosya mevcut değilse yaratmak istediğinizde True olarak kullanırız. Eğer burası 
		False ise ve aradığınız dosya yoksa 
		hata alırsınız. O yüzden ya burayı True yapmalısınız ya da dosyanın 
		mevcut olup olmadığını kontrol etmelisiniz. Mesela aşağıdaki kod ile, 
		dosya mevcut ise sonuna ekleme yapmak istiyoruz, mevcut değilse 
		yaratarak açıyoruz.</p>
		<pre class="brush:vb">
If Not fso.FileExists("C:\Users\Volkan\Desktop\denemeler\deneme2.txt") Then
    Set ts = fso.OpenTextFile("C:\Users\Volkan\Desktop\denemeler\deneme2.txt", ForWriting, True, TristateFalse)
Else
    Set ts = fso.OpenTextFile("C:\Users\Volkan\Desktop\denemeler\deneme2.txt", ForAppending, False, TristateFalse)
End If		</pre>
			<p>
			Son parametre Unicode desteği ile olup olmayacağını verir.</p>
		<h4> OpenAsTextStream</h4>
		<p> Elinizde bir File nesnesi varsa bunun <strong>OpenAsTextStream
		</strong>metodunu kullanarak da metin dosyalarını açabilirsiniz. Gerçi 
		File nesnesi için de yine bir FSO nesnesi gerekiyor. O yüzden her ikisini 
		de yaratmak gerekecek. Eğer File nesnesini başka birşey için 
		kullanmayacaksanız boşuna bu zahmete gerek yok, direkt FSO ve onun 
		metodları yeterli 
		olacaktır.</p>
		<p> <strong>Syntax:</strong>File.OpenAsTextStream(I/O modu,Format)</p>
		<p> İki parametre de opsiyonel olup default değerleri sırasıyla 
		ForReading ve TristateFalse'tur. Aşağıda bir örnek 
		bulunmakta.</p>
		<pre class="brush:vb">
Dim f As File, ts1 As TextStream, ts2 As TextStream
Set f = fso.GetFile("C:\Users\Volkan\Desktop\denemeler\deneme2.txt")

Set ts1 = f.OpenAsTextStream 'default değerlerle açıldı
x = ts1.ReadAll
ts1.Close

Set ts2 = f.OpenAsTextStream(ForAppending, TristateMixed)
y = ts2.Write("yeni")
ts2.Close</pre>
		<h3>TextStream Üyeleri</h3>
		<h4>Metin okuma şekilleri</h4>
		<pre class="brush:vb">ts.Read(5) 'Bulunulan yerden itibaren 5 karakter okur
ts.ReadLine 'Bulunulan satırı okur
ts.ReadAll 'Tüm dosya içeriğini okur</pre>
		<p>Üç yöntemde de bir değişkene atama işlemi yapılmalıdır.</p>
		<p>Ör: içerik=ts.ReadLine</p>
		<h4>Yazma şekilleri</h4>
		<pre class="brush:vb">ts.Write(metin):Dosyaya metni yazar
ts.WriteLine(metin):Dosyaya metni yazar ve bir alt satıra geçer
ts.WriteBlankLines(5):Dosyaya 5 adet boş satır ekler</pre>
		<p><span class="keywordler">Line </span>ile cursor'ın o anki satır 
		numarasını elde ederiz.</p>
		<p><span class="keywordler">Close </span>metodu ile TextStream nesnesini 
		kapatarız.</p>
		<p>Dosyada okuma yaparken, belirli koşullar durumunda o satırı
		<span class="keywordler">SkipLine</span> ile atlayarak bi sonraki satıra 
		geçebilriz. Aşağıdaki kod ile, sayısal bir ifadeyle başlayan herşeyi bir 
		collectiona atayıp en son da bunları yazdırıyoruz. Read ile bir karakter 
		okuduktan sonra kalanını ReadLine yaparken başına ilk okuduğumuz kısmı 
		eklediğimize dikkatinizi çekmek isterim. Örnek dosyamız aşağıdaki gibi 
		olsun</p>
			<pre>1-birinci satır
2-ikinci satır
falanfilan
3-üçüncü satır
falanfilan
4-dördüncü satır</pre>
		<pre class="brush:vb">
Sub Satıratla()
Dim ts As TextStream
Dim col As New Collection

Set ts = fso.OpenTextFile("C:\Users\Volkan\Desktop\denemeler\deneme1.txt", ForReading, False, TristateMixed)

Do
    kelime = ts.Read(1)
    If IsNumeric(kelime) Then
       col.Add kelime + ts.ReadLine
    Else
       ts.SkipLine
    End If
Loop Until ts.AtEndOfStream

For Each Item In col
    Debug.Print Item
Next Item	
End Sub</pre>
			<p>
			Çıktı ise şöyle olacaktır:</p>
			<pre>1-birinci satır
2-ikinci satır
3-üçüncü satır
4-dördüncü satır</pre>
			<p>
			Hepsi bir arada bir örneğimi aşağıdaki gibi olabilir:</p>
			<pre class="brush:vb">
Sub çeşitli_üyeler()
Dim ts As TextStream
Const dosya As String = "C:\Users\Volkan\Desktop\denemeler\ts_üyeler.txt"
Set ts = fso.CreateTextFile(dosya, True, True)

ts.WriteLine ts.Line &amp; "-" &amp; Now
ts.Write ts.Line &amp; "-": ts.WriteBlankLines (1) 'Dosyaya 1 adet boş satır ekler
ts.WriteLine ts.Line &amp; "-" &amp; Environ("username")
ts.WriteLine ts.Line &amp; "-" &amp; "selam"
ts.WriteLine ts.Line &amp; "-" &amp; "naber"


Debug.Print ts.Line
ts.Close

Set ts = fso.OpenTextFile(dosya, ForReading, False, TristateMixed)
x = ts.ReadLine ' Bulunulan satırı okur
ts.skipline 'ilgili satırı atlar
y = ts.Read(5) 'Bulunulan yerden itibaren 5 karakter okur, artık 3. satırdayız: 3-Vol
z = ts.ReadAll 'Cursordan itibaren tüm dosya içeriğini okur, baştan itibaren değil :kan4-selam5-naber

Debug.Print z

End Sub	</pre>
</div>
		<h2 class="baslik" id='logger'>Olay/Hata Logu tutan bir uygulama(Logger Prosedürü)</h2>
		<div class="konu">
		<p> Hergün belli satlere schedule edilmiş(<a href="DortTemelNesne_Application.aspx#OnTime">Application.Ontime</a> 
		aracılığı ile) makrolarınızın olduğunu 
		düşünün. Bunlar içinde çeşitli aşamaları gün/saat başta olmak üzere 
		diğer önemli bilgilerle birlikte kayıt altına almak, nerde hata alınmış, 
		bunları görmek isteyebilirsiniz, hatta iyi bir programcı olarak görmek 
		istemelisiniz.</p>
		<p> Keza, bölümünüz için raporlara ulaşım amacıyla hazırladığınız bir 
		arayüz(<a href="Formlar_Kontroller.aspx#kokpit">Kokpit Formu</a>) olması 
		durumunda, kim ne zaman hangi rapora girmiş, en çok hangi rapor 
		kullanılıyor, Kokpiti en çok kim kullanıyor gibi soruların cevabını elde 
		etmek için bir log kaydı da tutmak isteyebilirsiniz.</p>
		<p> İşte bu amaçlarla dosya yazma/okuma işlemlerini kullanabiliriz. </p>
		<h3> Otomasyon süreçlerinde Logger kullanımı</h3>
		<p> Diyelim ki aşağıdaki prosedür günün belirli saatlerinde çalışıyor. 
		Çalışmanın belirli aşamalarını(kritik önemde veya ana işlerin 
		öncesinde/sonrasında) kayıt altına alıyoruz. Ayrıca bir hata oluşursa yine bunu da 
		kayıt altına alalım.</p>
		<p> Otomasyon süreçlerinde Log tutmanın bir alternatifi kendinize veya 
		ilgili kişilere mail attırmak 
		olacaktır. Ancak çalışan çok fazla iş varsa mail kalabalığında 
		boğulursunuz. O yüzden log sistemi daha güzel bir seçenektir.</p>
		<p> Şimdi aşağıdaki örnekte Log dosyasmızda Tarih/Saat, Kullanıcı, 
		bilgisayar adı, rapor adı, log tipi, varsa hata kodu, açıklama kolonları 
		olmak üzere 7 kolon bilgi bulunmaktadır. Bunun ilk 3'ü Logger fonksiyonu 
		içinde dinamik olarak ele alınmakta, son 4 parametre ise Logger 
		fonksiyonuna KrediRaporu modülünden argüman olarak gönderilmektedir. 
		Diğer hususlar şöyledir.</p>
		<ul>
			<li>Rapor ismi toplamda 50 hane olacak şekilde ayarlanır. 50den kısa 
			olan rapor isimleri için başına 50ye tamalamanacak kadar boşluk 
			eklenir. Bunun amacı datayı Excele aktardığınızda aynı hizada görünemleri içindir. Bu 
			sizin dünyanızda daha yüksek bir sayıya ayarlanabilir.</li>
			<li>Bilgisayar ismi de yine aynı şekilde 10 haneye tamamlanmaktadır. 
			Bu da sizin dünyanızda daha yüksek ayarlanabilir.</li>
			<li>Hata yoksa hata kodu olarak 0 gönderilmektedir.</li>
		</ul>
		<pre class="brush:vb">
'*****Logger'ı çağıran prosedür*****
Sub KrediRaporu()

On Error GoTo hata
raporLoggerAd="KrediRaporu"

'çeşitli işler
Logger WorksheetFunction.Rept(" ", 50 - Len(raporLoggerAd)) & raporLoggerAd, "OK", 0, "Bölme işlemi başlayacak"
'çeşitli işler
Logger WorksheetFunction.Rept(" ", 50 - Len(raporLoggerAd)) & raporLoggerAd, "OK", 0, "Bölme işlemi bitti"
'çeşitli işler
Logger WorksheetFunction.Rept(" ", 50 - Len(raporLoggerAd)) & raporLoggerAd, "OK", 0, "Rapor başarıyla çalıştı"

Exit Sub
hata:
Logger WorksheetFunction.Rept(" ", 50 - Len(raporLoggerAd)) & raporLoggerAd, "Hata", Err.Number, Replace(Err.Description, vbNewLine, vbNullString)

End Sub 

'*****Logger prosedürümüz*****
Sub Logger(rpr As String, logtip As String, hatano As Integer, açıklama As String)

    On Error GoTo hata

    Dim dosya As String
    Dim dosyano As Variant

    dosyano = FreeFile
    dosya = gunlukklasor + "\GünlükRaporlarLog.txt"
 
    Open dosya For Append As #dosyano
    Print #dosyano, CStr(Now), Environ("UserName"), WorksheetFunction.Rept(" ", 10 - Len(Environ("computername"))) & Environ("computername"), rpr, logtip, hatano, açıklama
    Close #dosyano  

    Exit Sub

hata:

Call mail_logger_hata(rpr, alicilar) 'Log prosedüründe bir şekilde hata önceden berlirlenmiş alınırsa alıcılara özel formatta mail atılır

End Sub		</pre>
		<h3> Kokpit uygulamalarında Logger kullanımı</h3>
		<p> Userform konusunda gördüğümüz
		<a href="mailto:Formlar_Kontroller.aspx#kokpit">Kokpit uygulamalarında</a>, 
		uygulamayı kullanan kişlerin aktivitelerini aşağıdakine benzer bir kod 
		ile kayıt altına alabiliriz.</p>
		<pre class="brush:vb">
'Form üzerindeki bir butona tıklanınca
Sub Btn_KrediRaporAc()
On Error Goto hata
   'rapor açma kodları  
   Rapor="KrediRapor"
   frekans="Günlük"
   Call detayraporlogu(Rapor, frekans)

Exit Sub
hata:
On Error Goto -1
On Error Goto hata2

'burada hata kaydını tutan bir log kaydı(aşağıdaki log prosedürünü gölgede bırakmasın diye detayına girmedim)
Exit Sub

hata2:
'Diskte yer olmaması, veya kullanıcının ilgili diske yazma yetkisinin olmaması gibi bir sebeple hata olması durumunda
Call LogHata 'bu sefer maille size bilgilendirme yapılır
End Sub

'*****Logger prosedürümüz*****
Sub detayraporlogu(ByVal Rapor As String, ByVal frekans As String)

    If Environ("UserName") = sizinuserınız Then Exit Sub 'kendimizi loglamıyoruz
     
    i = FreeFile
    Open adres & "Kokpitlog_detayrapor.txt" For Append As i
    Print #1, Environ("UserName"), Date, Time, frekans, Rapor
    Close #1

End Sub</pre>
		<p>
		Daha sonra bu text dosyasını bir Excel dosya içine aktarırı veya her 
		açıldığında refresh olan bir bağlantı kurararak gelen data üzerinde pivot 
		tablolarınızı oluşturabilirsiniz. Metin dosylarından bağlantı kurmak için
		<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">buraya</a> 
		bakabilrisiniz.</p>
		<h4>
		Logger içinde hata</h4>
		<p>
		Bir sebeple logger fonksiyonu içinde de hata olursa bunu da başka bir hata 
		bloğuyla ele alabilirsiniz. Veya ana gönderici modülde On Error goto -1 
		deyip 2. bir hata bloğu açabilirsiniz. Yukarıdaki ilk örnekte Logger 
		fonksiyonu içinde hata bloğu ile yakaladık. İkinci örnekte ise ana 
		prosedürde On Error GoTo -1 yöntemini kullandık.</p>
		</div>
		<h2 class="baslik">Çeşitli Örnekler</h2>
		<div class="konu">
		<h4 class="baslik">Settings işlemleri</h4>
				<div class="konu">
		<p>Bir dosyadan bir database'in kullanıcı adı ve şifresini okuma, veya 
		bir dosyanın path'ini okuma gibi işlemler de bu sayfada 
		öğrendiklerimizle yapılabilir.</p>
		<p>Diyelim ki jenerik bir Add-in yaptınız. Bu Add-indeki makrolardan bir 
		tanesi bir klasördeki bir Excel dosyasını açacak. İşte bu Excel dosyanın 
		yerinin sabit olmasının mümkün olmadığı, bunun hangi klasörde olacağını 
		kullanıcıya bırakmanız gereken durumlar olabilecektir. 
		SettingforAddin1.txt gibi bir dosya içine bu klasörün tam path'i 
		yazılabilir. Hatta buna birden fazla dosya için birden fazla klasör de 
		eklenebilir. İstenirse ";" ile ayırılır, istenirse satır satır yazılır, 
		hiç farketmez. Yukarıdaki yöntemlerden biriyle ilgili adresi elde etmek 
		oldukça kolaydır.</p>
		</div>
		<h4 class="baslik">SQL 
		metinlerini değiştirme</h4>
		<div class="konu">
		<p>
		Diyelim ki raporlama araçlarınız çok hantal ve katı. Siz de gerek 
		kendiniz gerek departmanınız için Excel içinden çalışan hızlı ve esnek 
		bir raporlama platformu oluşturdunuz. İlgili raporların SQL'ini bir 
		metin dosyası içine koydunuz. Kullanıcıya tarih ve müşteri listesi gibi 
		sorular sordurarak dosyadaki parametrik kısımlarla kullanıcının verdiği 
		cevapları replace ettirerek nihai SQL'inizi elde edersiniz. Böylece uzun 
		bir SQL'i VBA içine satır satır yazmaktan kurtulmuş olursunuz. VBA içine 
		de SQL kodu yazılabilir ama bu hem kodun uzun ve çirkin görünmesine 
		neden olur hem de çok zahmetli bir iştir, özellikle SQL onlarca hatta 
		yüzlerce satırdan oluşuyorsa.</p>
		<p>
		Kodumuz şöyle olabilir:</p>
		<pre class="brush:vb">Sub SQLDeğiştir()

tarih=InputBox("tarihi girin")
If tarih=vbNullString Then Exit Sub

adres="C:\SQLller\kredi.txt"
Open adres For Input As #1
içerik=Input(LOF(1),1)
Close #1

strSQL=Replace(içerik,"trh",tarih) 'SQLi elde ettik
'bundan sonra SQL'i çalıştıracak kodlar devreye girer

End Sub</pre>
</div>
	</div>

</asp:Content>

