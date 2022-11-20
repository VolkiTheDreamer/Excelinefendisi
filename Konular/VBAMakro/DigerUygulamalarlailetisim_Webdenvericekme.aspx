<%@ Page Title='Webden veri çekme' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<script runat="server">

    protected void Page_Load(object sender, EventArgs e)
    {

    }
</script>


<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'>
        <table>
            <tr>
                <td>
                    <asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td>
                <td>
                    <asp:Label ID='Label2' runat='server' Text='Diğer Uygulamalarla iletişim'></asp:Label></td>
                <td>
                    <asp:Label ID='Label3' runat='server' Text='5'></asp:Label></td>
            </tr>
        </table>
    </div>
    <h1>Webden veri çekme</h1>
    <p>
        Web&#39;den farklı formatlarda veri çekilebilmektedir. pdf, imaj gibi yapısal olmayan dosyaları ahariç bırakırsak hedefimizde 3 tür dosya tipi olabilir.</p>
    <ul>
        <li>html: Özellikle &quot;table&quot; tarzındaki elementleri parse ederek hedefe ulaşırız. Bunun için HTML DOM bilgisi gereklidir.</li>
        <li>xml: Burda da XML DOM&#39;a aşina olmak gerekli. Ayrıca XPath ve XQuery kavramlarını da bilmekte fayda var.</li>
        <li>json: Bu, VBA&#39;in dictionarysine benzer. Bunun için hazır modülleri kullanacağız.</li>
    </ul>
    <p>
        Yöntemlere baktığımızda ise 3 ana kütüphane ile hedefe ulaşabiliriz. Bunlar;</p>
    <ul>
        <li>Internet Explorer ile html parsing yapılabilir</li>
        <li>XmlHttpRequest nesnesi ile html veya xml parsing yapılabilir, webservis de çağırabiliriz</li>
        <li>SOAP nesnesi ile webservisleri çağırabiliriz</li>
    </ul>
    <p>
        Bunların dışında 3rd party kütüphaneler(Ör:<a href="https://florentbr.github.io/SeleniumBasic/">Selenium</a>, <a href="https://github.com/VBA-tools/VBA-Web">VBA WEB tools</a>) de var ama biz bunlara girmeyeceğiz. Ayrıca Excel&#39;in yerleşik QueryTable nesnesi ile de web sorgulamaları yapılabilmekte ama bunla da kafanızı karıştırmak istemiyorum, kendim de çok kullandığımı söyleyemem.</p>
    <p>
        Hedef her ne olursa olsun, dönen sonuç üzerinde çeşitli işlemler yapmamız gerekecek. Bazen tek bir yöntem yeterliyken bazen burda göreceğimiz yöntemleri birleştirerek kullanmak gerekebilecektir.</p>
    <h2 class='baslik'>HTML elemanlarını parse etmek</h2>
    <div class='konu'>
    <p>
        <strong>ÖNEMLİ NOT:</strong> Aşağıdaki örnekteki kod geçerliliğini 
	yitirmiştir. Zira ilgili sayfanın tasarımcısı HTML kodlarını değiştirmiş olabilir. Yeni 
	örnek dosyayı <a href="../../Ornek_dosyalar/Makrolar/webdenveri2.xlsm">buradan</a> indirebilirsiniz, ancak aşağıda hala ilk haline göre olan 
	açıklamaları bulacaksınız. Bahsekonu sitede sürekli bir güncelleme hali 
	olabileceği için aynısını her defasında kendi siteme yansıtamıyacağım. Ancak 
	kodların çalışmaması durumunda <a href="../../iletisim.aspx">iletişim</a>
        sayfamdan bana bilgi verirseniz sadece örnek dosyayı güncelleyip siteye 
	ekleyebilirim.
    </p>
    <p>
        <strong>EDİT(16.02.2020)</strong>:Bu tür güncelleme süreciyle uğraşmamak için 2.örnek olarak kendi web sitemden bir örnek ekledim. buradaki kod her zaman çalışır olacaktır.
    </p>
        <p>
            <a href="../../Ornek_dosyalar/Makrolar/webdenveri.xlsm">Buradan</a> örnek 
	uygulamayı indirebilirsiniz.
    </p>
        <p>
            Hemen belirtmek isterim ki, bu yöntemi uygulamak için biraz da olsa Html, Html DOM(Document Object Model) ve Javascript(bazen 
		de css) bilgisi gereklidirBunlara aşina değilseniz 
		uygun bir tutorial sitesinden(w3school olabilir) temelleri almanızda 
		fayda var.
        </p>
        <p>
            Web'den tüm bir sayfa içeriğini VBA ile alabilmeniz mümkün olsa da 
		biz bu sitede daha çok verilerle ilgilendiğimiz için veri okumayla ilgili kısma 
		yoğunlaşacağız. Verilerin bulunduğu HTML elemanları da büyük çoğunlukla 
		<strong>table</strong> elemanları içinde bulunmaktadır.
        </p>
        <p>
            Lafı uzatmadan hemen örneğimize geçelim. Diyelim ki aşağıdaki sitede(<a href="https://kur.doviz.com/">https://kur.doviz.com/</a>) 
		bulunan döviz tablosunu Excel içine almak istiyoruz.
        </p>
        <p>
            <img src="../../images/vbawebhtml1.jpg"></p>
        <p>
            Şimdi bu noktada ilk yapmamız gereken tabloda bir yere gelip sağ 
		tıklamak ve "<strong>İncele</strong>" demek olacaktır.
        </p>
        <p>
            <img src="../../images/vbawebhtml2.jpg"></p>
        <p>
            Bunu yaptığımızda browser penceremizde bir bölme açılır ve seçtiğimiz 
		kısma konumlanarak onun elemanlarını bize gösterir.
        </p>
        <p>
            <img src="../../images/vbawebhtml3.jpg"></p>
        <p>
            Bu kısmı biraz kurcalarsak <strong>thead</strong> ve <strong>tbody</strong> kısımlarını görürüz. thead 
		kısmında ilgili listenin başlıkları yazmakta. tbody kısmında ise veriler 
		bulunmakta.(Daha geniş bir görüntüyü, sağ tıkladıktan sonra İncele yerine 
		<strong>Sayfa kaynağını görüntüle</strong> seçeneği ile görebilirsiniz)
        </p>
        <p>
            <img src="../../images/vbawebhtml4.jpg"></p>
        <p>
            Gördüğünüz üzere, bizim hedef kitlemiz "<strong>hisse-tablo</strong>" class'ına sahip 
		olan "tr" tag'leri. Yanız thead içinde de hisse-tablo var bi tane(her ne 
		kadar tam class adı "hisse-tablo hisse-tablo-row1" görünse de). 
		Dolayısıyla biz ilk "hisse-tablo" class'lı elemanı değil 
		sonrasındakilere bakmalıyız.
        </p>
        <p>
            Şimdi üstteki resimlerden belli olmadığı için tüm sayfa kaynağından 
		bakacak olursak, bu <strong>tr</strong> tag'inin açılmış halini daha açıkça görebiiriz. 
		Bunların içinde aşağıdaki gibi <strong>td </strong>taglerini görüyoruz. Bizim ihtiyacmız 
		olan bunlardan döviz adı, alış ve satış fiyatları; yani sırayla 1, 4, ve 
		5. elemanlar. Javascriptte indexler 0'dan başladığı için&nbsp; 0, 3 ve 4. 
		elemanlar. Bu td elemanları tr elemanlarının bir alt seviyesi olduğu için 
		bunlara tr'nin <strong>çocukları</strong> denir, yani <span class="keywordler">children
        </span>özelliği ile bunlara ulaşılabilir.
        </p>
        <p>
            <img src="../../images/vbaweb7.jpg">
        </p>
        <p>Şimdi kod yazımı için hazırlığımızı yapalım. </p>
        <p>
            Kurgumuz aşağıdaki gibi olacak. B1 hücresinden döviz kurlarının 
		kaynağını seçeceğiz.(Serbest piyasa mı yoksa başka bir kurumun fiyat 
		bilgileri mi diye). Buna göre dönen sonuç A5-C5 range'inden aşağıya doğru 
		gelecek. Yani bir worksheet_change event'i sözkonusu.
        </p>
        <p>
            <img src="../../images/vbaweb8.jpg">
        </p>
        <p>
            İlk olarak bir <a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">dictionary'ye</a> ihtiyacımız olacak. Bunu B1'den seçilen değerin karşılığı 
		olarak Url'ye bir ek yapmak için kullanacağız. Örneğin B1'de kaynak olarak 
		"Serbest Piyasa" seçilirse url eki Serbest-Piyasa oluyor. Bunun için 
		global geçerli olması gereken bir dict değişkenini Module1 içinde tanımladım. Buna hem 
		workbook_open eventinden hem de worksheet_change eventi içinde erişeceğiz. 
		Workbook_open içinden dictionary'yi dolduruyoruz ve B1 hücresine data 
		validation yapıyoruz.
        </p>
        <pre class="brush:vb">
'Module1 içeriği
Public dict As New Scripting.Dictionary

'ThisWorkbook içeriği
Private Sub Workbook_Open()

dict.Add "Serbest Piyasa", "Serbest-Piyasa"
dict.Add "Akbank", "Akbank"
dict.Add "Denizbank", "Denizbank"
dict.Add "Merkez Bankası", "Merkez-Bankasi"

For Each k In dict.Keys
    liste = liste & "," & k
Next k

    With Range("b1").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Mid(liste, 2, Len(liste) - 1)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub</pre>
        <p>
            Kodun bundan sonraki kısmı iki ayrı yöntemle yapılabilir. İlk olarak Internet Explorer yöntemini kullanıcaz, sonra da HTTP 
		yöntemini.
        </p>
        <h3>IE ile kod yazımı</h3>
        <p>
            Bunun için Internet Explorer'ın(IE) bir örneğini yaratıcaz, 
		bunu gerçekten açmıyoruz tabi, sadece bellekte açıyoruz. O yüzden IE'nin nesne modeline ihtiyacımız var. 
		(Bu yöntem için PC'nizde IE olması gerektiği aşikar.)
        </p>
        <p>
            Ayrıca HTML elemanları ile çalışacağımız için HTML nesne modeline de 
		ihtiyacımız bulunuyor.
        </p>
        <p>
            Bunları eklemek için Tools&gt;Reference üzerinden 
		aşağıdaki libraryler eklenir:
        </p>
        <p>
            <img src="../../images/vbawebhtml5.jpg"></p>
        <p>
            Worksheet_change eventinin kodu da aşağıdaki gibi olacaktır. Kod 
		içinde gerekli tüm açıklamalar bulunmakta.
        </p>
        <pre class="brush:vb">
Private Sub Worksheet_Change(ByVal Target As Range)

Dim IE As InternetExplorer
Dim elements As IHTMLElementCollection 'birçok tr tagini tutacak olan html collectionımız
Dim url As String
Dim url_ek As String
Dim r As Integer

On Error GoTo hata
If Not Target.Address = "$B$1" Then Exit Sub 'sadece B1'e tıklanırsa ektif olsun
Application.EnableEvents = False 'data çekilirken recursive şekilde tetiklenme olmasın diye eventleri geçici olarak pasifleştiriyoruz

Application.StatusBar = "Lütfen bekleyiniz..."
url_ek = dict([B1].Value)
[b2].ClearContents

başlangıç = Timer 'süre ölçmek istersek diye koydum

'önce eski veriyi temizleyelim
Range("Table1").Select
Selection.ClearContents

Set IE = New InternetExplorer
IE.Visible = False
url = "https://kur.doviz.com/" & url_ek
IE.navigate url 'ilgili sayfayı bellekte açıyoruz

'Bu kısım önemli, IE'yi bellekte açar açmaz sayfa yüklemesi hemen olmaz
Do While IE.Busy Or IE.READYSTATE <> 4 'sayfa yüklememsi tamamlanana kadar bekliyoruz
    DoEvents
Loop

Application.StatusBar = "Veri Çekiliyor..." 'veri çekilmeye başlandığında statusbarı güncelliyoruz
Set elements = IE.document.getElementsByClassName("hisse-tablo") 'classı "hisse-tablo" olan tüm elemanları elements değişkenine atıyoruz
'Application.Wait (Now + TimeValue("00:00:05")) 'eğer eksik veri geliyorsa tam olarak belleğe alamıyordur, bunun için gereken süre kadar bekleriz, burayı duruma göre sizin ayarlamanız gerekecek
r = 5 'başlangıç satırı

For i = 1 To elements.Length - 1 '0'dan değil 1den başlıyoruz, çünkü başlığın olduğu kısmı atlıyoruz
    If elements(i).ID = "linkUnit" Then GoTo atla 'RUS rublesindeki sonraki çizgiyi atlıyoruz

    Cells(r, 1).Value = elements(i).Children(0).innerText 'döviz adı, 0 indeksli yani 1. eleman
    'Küsuratlarda problem çıkmaması için virgülleri nokta ile replace ediyoruz
    Cells(r, 2).Value = Replace(elements(i).Children(3).innerText, ",", ".") 'alış fiyatı 3 indeksi yani 4.eleman
    Cells(r, 3).Value = Replace(elements(i).Children(4).innerText, ",", ".") 'Satış fiyatı 4 indeksi yani 5.eleman
    r = r + 1
atla:
Next i

bitiş = Timer
Debug.Print bitiş - başlangıç 'geçen süreyi yazdırıyoruz
IE.Quit 'ilgili sayfada reklam v.s varsa arka planada müzik çalmaya devam eder, bu satırla arka planda açık olan Internet Exlporer'dan çıkarız ve reklam müziği sona erer
Set IE = Nothing: Set htmldoc = Nothing: Set elements = Nothing 'ilgili objeleri bellekten atıyoruz

Range("Table1[[Alış]:[Satış]]").Select
Selection.NumberFormat = "0.00"
    
[lastruntime].Value = Now 'güncelleme zamanını yazıyoruz
[lastruntime].Select

Application.StatusBar = "İşlem tamam"
Application.EnableEvents = True 'Enableevent özelliğini tekrar aktive ediyoruz
Exit Sub

hata:
    Set IE = Nothing: Set htmldoc = Nothing: Set elements = Nothing
    Application.StatusBar = ""
    MsgBox "Bi hata oluştu" + Err.Description
    Application.EnableEvents = True
    
End Sub	</pre>
        <p>
            For Next döngüsünü aşağıdaki gibi de yapabilirdik, ancak bu sefer ilk If 
		kontrolünü tüm döngü boyunca yapmak durumunda kalırdık, ki bu da 
		performansı yoran bir işlem olurdu, o yüzden yukarıdaki yöntem daha 
		hızlıdır.&nbsp;Ama bu tür kontrollerin gerekmediği durumlarda For Each 
		döngüleri daha pratik olmaktadır.
        </p>
        <pre class="brush:vb">
'En başta da şu tanım yapılmalıdır
Dim element As IHTMLElement

For Each element In elements
    If element.className = "hisse-tablo hisse-tablo-row1" Then GoTo atla 'başlığın olduğu kısmı atlıyoruz
    If element.ID = "linkUnit" Then GoTo atla 'RUS rublesindeki sonraki çizgiyi atlıyoruz

    Cells(r, 1).Value = element.Children(0).innerText 'döviz adı, 0 indeksli yani 1. eleman
    Cells(r, 2).Value = Replace(element.Children(3).innerText, ",", ".") 'alış fiyatı 3 indeksi yani 4.eleman
    Cells(r, 3).Value = Replace(element.Children(4).innerText, ",", ".") 'Satış fiyatı 4 indeksi yani 5.eleman
    r = r + 1
atla:
Next element</pre>
        <p>
            Gördüğünüz üzere, html elemanlarını parse ederken çeşitli metodlar ve 
		özellikler kullanırız. Bunlar özetle aşağıdaki gibi olup özellikle 
		bunları araştırmanız ve öğrenmenizi tavsiye ederim.
        </p>
        <ul>
            <li><span class="keywordler">getElementsByTagName</span>: Çok 
			elemandan oluşan bir&nbsp; IHTMLElementCollection collection'ı 
			döndürür</li>
            <li>
                <p>
                    <span class="keywordler">getElementsByClassName</span>: Çok 
			elemandan oluşan bir&nbsp; IHTMLElementCollection collection'ı 
			döndürür
                </p>
            </li>
            <li>
                <p>
                    <span class="keywordler">getElementById</span>: Tek bir 
			IHTMLElement elemanı döndürür
                </p>
            </li>
            <li>
                <p>
                    <span class="keywordler">innerText</span>: İlgili elemanın 
			içindeki metini verir.
                </p>
            </li>
            <li>
                <p>
                    <span class="keywordler">innerHTML</span>: İlgili elemanın tüm 
			HTML metnini verir.
                </p>
            </li>
            <li>
                <p>
                    <span class="keywordler">textContent</span>: İlgili elemanın 
			içindeki &lt;span&gt; v.s elemanları kapsayacak şekilde metnini verir.
                </p>
            </li>
        </ul>
        <p>
            <strong>NOT</strong>:Unutmayın ki, web siteleri zamanla değişebilir. 
		Bu nedenle kodunuzda 
		zaman zaman güncelleme yapmanız gerekebilir.
        </p>
        <p>NOT: Yukarıda IE.document.getElementsByClassName diyerek arada olşuan HTMLDocument tipli nesnesyi bypass ediyoruz. eğer html dokümanını kendisiyle işlem yapacaksak, mesela body&#39;sini okumak gibi, bunu bypass etmeden bi değişkne atayabiliriz. Örneğin.</p>
        <pre>Dim htmldoc As MSHTML.HTMLDocument
......
Set htmldoc = IE.document
bodystr=htmldoc.body.innerText</pre>
        <h3>2.Örnek</h3>
        <p>
            Şimdi ise, konuyu pekiştirmek adına ve değişme ihtimali çok düşük bir sayfadan veri çekeceğiz. Değişme ihtimalinin çok düşük olması sayfanın benim web sitemdeki bir <a href="../Excel/Giris_PratikKisayollar.aspx">sayfa</a> olmasından. Olur da bir değişiklik yaparsam bu çok büyük ihtimalle “class” isminde olacaktır.</p>
            <p>Hemen örneğimize geçelim.</p>
                <p>Veri çekeceğimiz tablo ve html görüntüsü aşağıdaki gibidir.</p>
        <p>
            <img alt="" src="/images/vba_webdenveri0.jpg" />
        </p>
        <p>
            <span>Buna göre kodlarımız şöyle olacaktır. Tüm açıklamaları kod içinde bulabilirsiniz.</span>
        </p>
        <pre class="brush:vb">
Sub webdenveri()
 
Dim IE As InternetExplorer
Dim tablolar As IHTMLElementCollection 'tablolarımızın ID'si yok, classı var, class olması demek 1den çok tablo olabilir demek,
'bu sayfamızda 1 tablo var gerçi, ama biz yine de bu şekilde ilerlemek durumundayız
Dim tbody As IHTMLElement 'tablodaki tbody elementini tutacak
Dim r As Integer
 
On Error GoTo hata
 
Set IE = New InternetExplorer
IE.Visible = False
url = "https://www.excelinefendisi.com/Konular/Excel/Giris_PratikKisayollar.aspx"
IE.navigate url 'ilgili sayfayı bellekte açıyoruz
 
'Bu kısım önemli, IE'yi bellekte açar açmaz sayfa yüklemesi hemen olmaz
Do While IE.Busy Or IE.READYSTATE <> 4 'sayfa yüklememsi tamamlanana kadar bekliyoruz
    DoEvents
Loop
 
Set tablolar = IE.document.getElementsByClassName("alterantelitable") 'classı "alterantelitable" olan tüm tabloları tablolar değişkenine atıyoruz, 1 tane var zaten
'Application.Wait (Now + TimeValue("00:00:05")) 'eğer eksik veri geliyorsa tam olarak belleğe alamıyordur, bunun için gereken süre kadar bekleriz, burayı duruma göre sizin ayarlamanız gerekecek
r = 1 'başlangıç satırı
Set tbody = tablolar(0).Children(0) 'ilk tablonun(her ne kadar 1 tane olsa da) ilk elementi, yani tbody
For i = 0 To tbody.Children.Length - 1 'tbody altındaki tüm child elementlar kadar, yani tr tagleri kadar, özetle tüm satırlarda döneceğiz
    Cells(r, 1).Value = tbody.Children(i).Children(0).innerText 'tbodynin ilk satırının ilk child elemanı, yani ilk td'si, yani ilk kolonu(ilk satır için td değil, th ama bizim için değişen birşey yok)
    Cells(r, 2).Value = tbody.Children(i).Children(1).innerText 'tbodynin ilk satırının ikinci child elemanı, yani ikinci td'si, yani ikinci kolonu
     r = r + 1
Next i
 
IE.Quit 'ilgili sayfada reklam v.s varsa arka planada müzik çalmaya devam eder, bu satırla arka planda açık olan Internet Exlporer'dan çıkarız ve reklam müziği sona erer
Set IE = Nothing: Set tablolar = Nothing: Set tbody = Nothing 'ilgili objeleri bellekten atıyoruz
Exit Sub
 
hata:
    Set IE = Nothing: Set tablolar = Nothing: Set tbody = Nothing
    MsgBox "Bi hata oluştu, " + Err.Description
  
End Sub</pre>
        <p>Sonuç aşağıdaki gibi olacaktır.</p>
        <p>
            <img alt="" src="/images/vba_webdenveri1.jpg" /></p>
        <p>Gerekli formatlama işlerini size bırakıyorum.</p>
        <h3>Http Yöntemi ile</h3>
        <p>Bu konuyu aşağıda işledim. Aslında burda yapılan da günün sonunda parsing olacak ancak ama HTTP yönteminin başka boyutları da olduğu için aşağıda bir bütün olarak almayıtercih ettim.&nbsp;</p>
    </div>

    <h2 class="baslik">XMLHttpRequest nesnesinin detayları</h2>
    <div class="konu">
        <h3>Http</h3>
        <p><a href="https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest">XMLHttpRequest</a> nesnesi, VBA dahil birçok programlama dilinde http bağlantısı kurmayı sağlayan bir nesnedir. Biz burada doğal olarak VBA için hazırlanmış olan kütüphaneyi kullanacağız.</p>
        <p>Konuya geçmeden önce HTTP hakkında genel bilginiz olması gerekiyor. Eğer bu konuda kendinizi yetersiz hissediyorsanız aşağıdaki linklerden ön bilgi alabilirsiniz.</p>
        <p>Genel http bilgisi</p>
        <ul>
            <li><a href="https://developer.mozilla.org/en-US/docs/Web/HTTP">https://developer.mozilla.org/en-US/docs/Web/HTTP</a> </li>
            <li><a href="https://www.ionos.com/digitalguide/hosting/technical-matters/http-header/">https://www.ionos.com/digitalguide/hosting/technical-matters/http-header/</a></li>
            <li><a href="https://code.tutsplus.com/tutorials/http-headers-for-dummies--net-8039">https://code.tutsplus.com/tutorials/http-headers-for-dummies--net-8039</a> </li>
            <li><a href="http://www.steves-internet-guide.com/http-basics/">http://www.steves-internet-guide.com/http-basics/</a> </li>
        </ul>
        <p>
            daha sonra tekrar buraya gelip devam edin.
        </p>
        <p>Evet, bu linklere baktıysanız görmüşsünüzdür ki; HTTP, host ile client arasındaki bir protokoldür. Biz client olarak bir host makinaya <strong>request</strong>(istek) göndeririz, ki bu istek bir URl şeklindedir; o makine de bize <strong>response</strong>(yanıt) döndürür.</p>
        <p>Örnek bir GET isteği şöyledir: <a href="https://wwww.falancasite.com/query?singer=tarkan&amp;yil=2012">https://wwww.falancasite.com/query?singer=tarkan&amp;yil=2012</a> </p>
        <p>Browserlardaki adres şubuğunua yazdığımız herşey bir GET isteğidir. POST&#39;ta ise daha çok hassas bilgi göndererek bilgi almaya çalışırız veya bir veritabanı kayıt işlemi gerçekleştiririz. Bu iki metod arasındaki farklar için <a href="https://www.w3schools.com/tags/ref_httpmethods.asp">https://www.w3schools.com/tags/ref_httpmethods.asp</a> sayfasındaki &quot;Compare GET vs. POST&quot; kısmına bakabilirsiniz.</p>
        <p>Bu arada Http denemeleri yapmak için <a href="https://httpbin.org/">https://httpbin.org/</a> ve <a href="http://ptsv2.com/">http://ptsv2.com/</a>&nbsp; sayfalarını kullanabilirsiniz.</p>
        <h4>VBA dünyası için ön bilgiler</h4>
        <ul>
            <li>XMLHttpRequest adı sizi yanıltmasın, ilk başta sadece XML varken bu isim verilmiş, her tür veri alma/gönderme için bu nesne kullanılabilir.</li>
            <li>XMLHttpRequest ile asenkron işlem de yapılabilir ki, AJAX denen teknolojide bu nesne kullanımı kritiktir. Ancak biz buradan hep senkron çalışacağız.</li>
            <li>Genel syntax şöyledir:<span class="keywordler"> XMLHttpNesnesi.Open(strMethod,&nbsp;strUrl,&nbsp;varAsync,&nbsp;strUser,&nbsp;strPassword)</span></li>
        </ul>
        <p>Open metodu bağlantıyı açar ama henüz bi bilgi gönderimi yoktur. Bu metodun parametrelerine bakalım:</p>
        <ul>
            <li>strMethod: GET veya POST değerini alır(başka da var ama bize bu ikisi yeterli). Hangisini seçeceğimizi bizim bilmemiz lazım, bunun için ilgili URL&#39;e gidip F12 tuşuna basarak network sekmesinden bunu görebiliriz. İlgili URL GET mi istiyor POST mu, bu bilgi burada görünüyor. Bi parametre vermiyorsak genelde GET olacaktır.</li>
            <li>strUrl: Gitmek istediğimiz URL</li>
            <li>varAsync: Senkron çalışacağımız için hep False atayacağız</li>
            <li>Son iki parametreyi otantikasyon gereken işlemlerde kullanacağız</li>
        </ul>
        <p>
            <span class="keywordler">Send</span> metodu ile de bağlantı isteğini(request) göndeririz ve daha sonra dönen bilgiyi de işleriz. Şimdi yapılacak işlere sırayla bakalım.
        </p>
        <h4>Kütüphane ekleme ve nesneyi kullanma</h4>
        <p>VBA&#39;de HTTP requesti yapmamızı sağlayan iki library/reference(ve 3 sınıf) bulunmakta. Bunlar <span class="keywordler">XmlHttpReuest</span>(hem XMLHttp hem ServerXMLHttp sınıfları bunda) ve <span class="keywordler">WinHTTP</span>&#39;dir. İkincisi, ilknin biraz daha gelişmiş versiyonudur denebilir. Hangi durumlarda hangisini kullanmak gerekir sorusuna ait bilgileri <a href="https://stackoverflow.com/questions/11605613/differences-between-xmlhttp-and-serverxmlhttp">bu</a> ve <a href="https://blog.srpcs.com/picking-the-correct-xmlhttp-object/">şu</a> sayfalardan bakabilirsiniz, bununla beraber biz burada sadece basit XmlHttp sınıfını göreceğiz. </p>
        <p>Genel olarak Http yöntemi yukarda anlattığımız IE yönteminden daha basit olabilmektedir. Özellikle, ilgili host bize web service, api veya json dosya ile bize daha yapısal bir formatla sonuç döndürüyorsa. Aralarındaki farkları aşağıda göreceğiz.</p>
        <p>Early Binding yapmak için aşağıdaki şekilde eklenebilir.</p>
        <p>
            <img alt="" src="/images/vba_webdenveri2.jpg" /></p>
        <p>Sonrasında ilgili nesneyi şöyle tanımlarız:</p>
        <pre class="brush:vb">
Dim req As New MSXML2.XMLHTTP60 &#39;Sondaki 60 şu: ilgili kütüphanenin 6.0 versiyonu
req.Open &quot;GET&quot;, requrl, False
req.Send</pre>
        <p>Late Binding için ise şöyle yazarız.</p>
        <pre class="brush:vb">
Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")</pre>
        <p>
            Şimdi isteğimizi(request) gönderdik, bize bir cevap(response) dönecek. Eğer dönen şey metinsel bir bilgi ise (html, xml, json, text türünde bir çıktı) bunu <span class="keywordler">responseText</span> propertysi ile okuruz, dönen cevap XML dokuman ise bunu <span class="keywordler">responseXml</span> propertysi ile alırız.
        </p>
        <p>
            Tabi öncesinde isteğimizin başarılı olup olmadığını <span class="keywordler">status</span> özelliği ile kontrol etmemizde fayda var.
        </p>
        <p>
            Hadi şimdi örneklerimize geçelim. Bir önceki örneğin devamı olarak ilerlemiyorsanız referans olarak <strong>HTML Object Library</strong>&#39;sini de eklemeyi unutmayalım. Ayrıca Dictionarylerden yararlanacağımız için <strong>Scripting Runtime</strong> da ekleyelim.
        </p>
        <h4>GET Örnekleri</h4>
        <h5>Basit bir GET isteği</h5>
        <pre class="brush:vb">
Sub basicget()
    Dim url As String
    Dim request As New MSXML2.XMLHTTP60
    
    url = "http://www.example.com"
    request.Open "GET", url, False
    request.send
    
    cevap = request.responseText
    Debug.Print cevap
        
End Sub
</pre>
        <h5>Parametre Gönderme</h5>
        <p>
            Bu örnekte benim kendi sitemden yaptığım bir sayfaya AnaKonu ve Altkonu parametrelerini vererek bu altkonu altında kaç konu olduğunu döndüren bir sorgulama yapacağız. 
            Parametreleri URL&#39;in bir parçası olarak eklemek yeterlidir. Siz de farklı bir Anakonu ve Altkonu vererek test edebilirsiniz.</p>
        <pre class="brush:vb">
Sub parametreliget()
    Dim url As String
    Dim request As New MSXML2.XMLHTTP60
    
    url = "https://www.excelinefendisi.com/httpapiservice/ResponseveRequestTarget.aspx?Anakonu=VBAMakro&Altkonu=Temeller"
    request.Open "GET", url, False
    request.send
    
    cevap = request.responseText
    Debug.Print cevap
        
End Sub</pre>
        <p>
            Bunun sonucu aşağıdaki gibi olacaktır. 
                Burada istersek özellikle işaretlediğim kısmı da html parsing ile elde edebiliriz, buna ait bir örneği bir altta bulabilirsiniz.</p>
        <p>
            <img src="../../images/vbawebhttp1.jpg" style="width: 965px; height: 541px" />
        </p>
        <html>
        <body>
            <form method="post" action="./ResponseveRequestTarget.aspx?Anakonu=VBAMakro&amp;Altkonu=Temeller" id="form1">
            </form>
        </body>
        </html>
        <h5>HTML ile birleştirip Parsing yapalım </h5>
        <p>Bu örnek yukarıda IE ile yaptığımız örneğin aynısıdır.</p>
        <pre class="brush:vb">
Sub htmlandhttp()
Dim oXMLHTTP As New MSXML2.XMLHTTP60
Dim htmlObj As New HTMLDocument
  
With oXMLHTTP
    .Open "GET", "https://www.excelinefendisi.com/Konular/Excel/Giris_PratikKisayollar.aspx", False
    .Send

    If .readyState = 4 And .Status = 200 Then
        Set htmlObj = CreateObject("htmlFile")
        htmlObj.body.innerHTML = .responseText
        
        Set tablolar = htmlObj.getElementsByClassName("alterantelitable")
        
        Set tbody = tablolar(0).Children(0)
        For i = 0 To tbody.Children.Length - 1
           Debug.Print tbody.Children(i).Children(0).innerText & " , " & tbody.Children(i).Children(1).innerText
           
        Next i

    End If

End With

End Sub&nbsp; </pre>
        <h5>Otantikasyon</h5>
        <p>
            Bu sefer <a href="http://ptsv2.com">http://ptsv2.com</a> sitesinden kendi oluşturduğum bir URL&#39;i kullanacağız. Burada user ve password bilgilerini sırasıyla volki ve tolki olarak veriyorum. İsterseniz önce aşağıdaki linki browserda kendiniz deneyin.
            Size bir User ve Password kutusu çıkarakcatır. Bu kodla bu kutulara da giriş yapmışız gibi oluyoruz.</p>
        <pre class="brush:vb">
Sub authanticated()
Dim req As New MSXML2.XMLHTTP60
url = "http://ptsv2.com/t/volkitolki/post"

With req
    .Open "GET", url, False, "volki", "tolki"
    .Send
    response = .responseText
End With

Debug.Print response

End Sub         &nbsp;</pre>
        <p>
            Bu kod çalıştığında &quot;naber kanka&quot; metni dönecektir.
        </p>
        <h5>Header bilgisi de belirtelim</h5>
        <p>
            İstek yaparken bazı bilgilerin headerda iletilmesi gerekebilir. Bunun için setRequestHeader metodu kullanılıyor. Mesela sadece belirli bir tarihten sonra oluşan taze bilgiği almak isterseniz şöyle bi header bilgis&nbsp; geçebilirsiniz: request.setRequestHeader "If-Modified-Since", "Sat, 24 Apr 2021 00:00:00 GMT" </p>
        <p>Aşağıdaki örnekte hedefin content-type&#39;ını da belirliyoruz. </p>
        <h5>Metin dışındaki bilgilerin(imaj, pdf v.s) alınması</h5>
        <p>Burda dönüş bilgisini <strong>responseText</strong> ile değil <strong>responseBody</strong> ile alıyoruz, ki bu bize <strong>byte</strong> tipinde bir dizi verir.</p>
        <pre class="brush:vb">
Sub dosyaveyaimaj()
Dim request As New MSXML2.XMLHTTP60
Dim ado As New ADODB.Stream 'raw byte olan body'yi okumak için. "open for binary" ve sonrasında "put" statement diyerek de yapılır deniyor ama ben başaramadım, ADODB ile oldu
Const dosya As String = "<a href="file:///E:/OneDrive/Masaüstü/httpimaj.jpg">E:\OneDrive\Masaüstü\httpimaj.jpg</a>" &#39;hedef dosyamız bu, bunun içine yazılacak

url_jpg = "https://www.excelinefendisi.com/anasayfa.jpg"

request.Open "GET", url_jpg, False
request.setRequestHeader "Content-Type", "image/jpg"
request.send

fileBytes = request.responseBody

With ado
    .Open
    .Type = adTypeBinary
    .Write request.responseBody 'raw byte olarak gelir
    .Position = 0
    .SaveToFile dosya, adSaveCreateOverWrite
End With

End Sub  </pre>
        <h5>
            Json dosya</h5>
        <p>
            Json konusun çok detaylı bi konu olduğu için <a href="#json">aşağıya</a> koydum.</p>
        <h5>
            XML dosya </h5>
        <p>
            <a href="https://github.com/VBA-tools">VBATools&#39;un</a> JSON için olduğu gibi&nbsp;
            XML için de converter <a href="https://github.com/VBA-tools/VBA-XML">modülü</a> var ancak şuan gelişm aşamasındaymış, gerçekten ben bizzat denedim hata aldım. Ama sorun değil, XML işleme için başka yöntemler var, zira XML Jsondan çok daha eski bi teknoloji olduğu için VBA içinde default gelen libraryler içinde XML parsing yapmaya yarayan sınıflar bulunuyor. Bu arada bu sınıfları kullanmak için biraz XML ve XML&#39;le alakalı kavramları(özellikle Xpath) bilmeniz gerekiyor, tıpkı yukardaki HTML objelerinde olduğu gibi. Bunun için <a href="https://www.w3schools.com/xml/">w3school&#39;dan</a> fadaylanabilrsiniz.</p>
        <p>
            XML responseları tabiki responseText ile text olarak elde edilebilir ama responsexml ile xml objesi olarak işlemenin daha çok avantajı vardır, yapısal bi nesne olması sebebiyle. Biz de sadece buna bakacağız. Aşağıdaki kodda XML&#39;i 3 farklı şekilde ele alacağız, o yüzden F8 ile ilerlemenizde fayda var. O sırada ele aldığımız nesnenin tam olarak ne tür bir nesne olduğunu bilmek için TypeName yazdırmanız faydalı olacaktır, ona göre en baştaki değişken deklerasyon bölümünde doğru tipli nesne yaratırsınız, bu da intellisenseten faydalanmanızı sağlayacaktır. Yani ilk başta hiç bir değişkeni tanımlamadan yola çıkabilir, kod aralarında TypeName yazdırarak sırayla bi sonraki nesneyi tanımlayabilirsiniz. Aşağıda, commentlenmiş böyle bi kaç satır görebilirsiniz.</p>
        <pre class="brush:vb">
Sub xmile()
Dim request As New MSXML2.XMLHTTP60
Dim respxml As MSXML2.DOMDocument60 'XML olarak dönen nesne bu olacak
Dim root As IXMLDOMElement 'en baştaki xml bilgisi(prolog) dışındaki ana root elementi döndürüler
Dim snodes As IXMLDOMSelection 'selectnode olarak alınnalar
Dim mynodelist As IXMLDOMNodeList 'getelementtsbytagname olarak alına
Dim nd As IXMLDOMNode 'nodelist içindeki her bir item
Dim cn As IXMLDOMElement, cn2 As IXMLDOMElement 'childnode olarak alınanlar

url = "https://www.w3schools.com/xml/simple.xml"
request.Open "GET", url, False
request.send

Set respxml = request.responseXML
'Debug.Print TypeName(respxml)
'Debug.Print respxml.ChildNodes.Length
Set root = respxml.DocumentElement
'Debug.Print TypeName(root)
'Debug.Print root.ChildNodes.Length


'tagnamelere göre
Set mynodelist = respxml.getElementsByTagName("food")
For Each nd In mynodelist
    Debug.Print nd.FirstChild.Text
Next nd


'child nodelarda dolaşarak: 2 derinlik için 2 for veya childnodes propertysi, 10 derinlik olsaydı 10 for veya 10 nested child olacaktı
Set xdoc = respxml.DocumentElement
For Each cn In xdoc.ChildNodes
    Debug.Print cn.ChildNodes(2).Text
Next cn

'xpath ile daha kolay. içiçe 10 for yerine tek satırda 10 "/" var. şimdi sadece tek for ile elde edilen liste üzerinde dolaşırız
Set snodes = root.SelectNodes("//food/name")
'veya aşağıdaki gibi
'Set snodes = respxml.SelectNodes("//breakfast_menu/food/name")
'Set snodes = respxml.ChildNodes(1).SelectNodes("//breakfast_menu/food/name")

For Each s In snodes
    Debug.Print s.Text
Next s
End Sub    &nbsp;</pre>
        <p>
            Bu arada olur da localinizdeki bir XML dosya ile çalışacaksanız bunu ilgili XML objesini <strong>Load</strong> metodu ile bir XML dosyayı ve <strong>LoadXML </strong>ile bir XML stringi okuyup sonrasında benzer işlemleri yapabilirsiniz.</p>
        <h4>POST Örnekleri</h4>
        <p>Şifre gibi hassas bir bilgi gönderilecekse, POST sorgulaması yapılır. Evet, düşünülenin aksine, POST sadece veritabanında güncelleme yapmak için kullanılmıyor, hassas bilgi gönderilerek yapılan sorgulamalarda da kullanılıyor. Biz de şimdi benim siteme üye olurken girdiğiniz mail ve şfire bilgilerinizle ne zaman üye olduğunuzu veya ETK izni verip vermediğinizi göreceksiniz.</p>
        <pre class="brush:vb">
Sub postornek()
    Dim xhr As New MSXML2.XMLHTTP60
    url = "https://www.excelinefendisi.com/httpapiservice/ResponseveRequestTarget.aspx"
    'parametrede boşluk, +, @  varsa ya WorksheetFunction.Encodeurl kullanın
    payload = "{'MailAdres':'volkan.yurtseven@hotmail.com', 'Sifre':'......'}" 'json formatında vercem parametreleri, o yüzden setrequestte json yapıcam.
    'şöyle de olabilirdi. send "key1=value1", ama tabi hedef URL'in gelen requesti işleme önemli, ben sadece json kabuledecek bir kurgu yapmıştım
    xhr.Open "POST", url, False
    xhr.setRequestHeader "content-type", "application/json" 'body'mizi nasıl gönderdiğimizi söylemiş oluyoruz
    xhr.send payload 'put ve post'ta body de geçirilebiliyor, ki genelde parametreler oluyor
    sResult = xhr.responseText
    Debug.Print sResult
End Sub</pre>
        <p>Siz de kendi mail adresiniz ve şifrenizle üyelik tarihinizi ve ETK izni verip vermediğinizi görebilirsiniz. Eğer üye olmadıysanız ve olmak istemiyorsanız sırayla alfa ve beta ifadelerini de yazabilirsiniz.</p>
        <h4>Jenerik fonksiyon</h4>
        <p>Gördüğümüz gibi kullandığımız fonksiyon üç aşağı beş yukarı hep aynı gibi, o yüzden farklı durumlar için hep baştan kod yazmak yerine jenerik bi fonksiyon hazırlasak en güzeli. </p>
        <p>Birden fazla sonuç döndürmek için ByRef özelliğinden faydalanıyorum. Ayrıca her zaman otantikasyon veya requestheader bilgisi göndermeyebileceğimiz için bunları optional tanımlıyorum. Credentials(otantikasyon için) bilgisini 2 elemanlı bir variant dizi olarak iletiyorum, request headerlarını ise birden çok key-value ikilisi olabileceği için Dictionary olarak geçiriyorum.</p>
        <pre class="brush:vb">
Function httpGet(ByVal url As String, ByRef header As Variant, ByRef statü As String, Optional ByVal credentials As Variant, Optional ByVal setreqheader As Dictionary)
    Dim request As New MSXML2.XMLHTTP60 
               
    With request
        If IsMissing(credentials) Then
            .Open "GET", url, False
        Else
            .Open "GET", url, False, credentials(0), credentials(1)
        End If
        
        If Not setreqheader Is Nothing Then
            For Each rh In setreqheader.Keys
                .setRequestHeader rh, setreqheader(rh)
            Next rh
        End If
        .send
                       
                
        Debug.Print "Tüm response headerlar, fonksiyon içinde yazılıyor...." & vbNewLine
        Debug.Print .getAllResponseHeaders 'sadece Headerla ilgilenseydik GET yerine HEAD ile göndermemiz yeterliydi
        
        'diğer bilgiler ise çağıran fonksiyona döndürülüyor
        header(0) = .getResponseHeader(header(1))
        statü = .Status & "," & .statusText
        httpGet = .responseText
    End With
End Function
            </pre>
        <p>Bunu kullanacak börnek bir kodumuz aşağıdaki gibi olabilir. (Farklı urller için comment/uncomment yaparak veya uygun eklemeler yaparak deneyebilirsiniz)</p>
    <pre class="brush:vb">
Sub jenerik1()
    Dim url As String, r_sta As String, r_body As String
    Dim header As Variant
    Dim reqheader As New Dictionary
        
    header = Array(vbNullString, "Content-Type")
        
    reqheader.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    reqheader.Add "Accept", "application/json;indent=2" 'servera, clientta sdece json işleyebildiğimizi söylüyoruz
        
    &#39;url = "http://www.example.com"
    &#39;url = "http://ptsv2.com/t/trump/d/1"
    &#39;url = "https://www.excelinefendisi.com/Sitemap.xml"
    url = "https://httpbin.org/get"
    &#39;url = "https://www.excelinefendisi.com/httpapiservice/ResponseveRequestTarget.aspx?Anakonu=VBAMakro&Altkonu=Temeller"
        
    r_text = httpGet(url, header, r_sta, reqheader)
    &#39;r_text = httpGet(url, header, r_sta)
        
    Debug.Print "Sonuçlar yazılıyor..."
    Debug.Print "Header:" & header(0)
    Debug.Print "Sta    Debug.Print r_body = r_text
    Debug.Print "Responesetext:" & r_text
        
End Sub       </pre>

        <p>Bir diğeri de şöyle, bundan otantikasyon yapıyoruz.</p>
        <pre class="brush:vb">
Sub jenerik2()

    Dim url As String, r_sta As String
    Dim header As Variant
    
    header = Array(vbNullString, "Content-Type")    
    
    url = "http://ptsv2.com/t/volkitolki/post"
    cred = Array("volki", "tolki")    


    r_text = httpGet(url, header, r_sta, cred)
    
    Debug.Print "Sonuçlar yazılıyor..."
    Debug.Print "Header:" & header(0)
    Debug.Print "Statü:" & r_sta
    Debug.Print "Responesetext:" & r_text

End Sub
                &nbsp;</pre>
        <p>XMLHTTP ile ilgili ilave bilgi için şu kaynaklara bakabilirsiniz</p>
        <ul>
            <li><a href="https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v=vs.85">https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v=vs.85</a>) </li>
            <li><a href="https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms759148(v=vs.85">https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms759148(v=vs.85</a>) </li>
            <li><a href="https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest">https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest</a> </li>
            <li><a href="https://www.w3schools.com/xml/xml_http.asp">https://www.w3schools.com/xml/xml_http.asp</a> </li>
            <li><a href="http://www.tushar-mehta.com/publish_train/xl_vba_cases/vba_web_pages_services/index.htm">http://www.tushar-mehta.com/publish_train/xl_vba_cases/vba_web_pages_services/index.htm</a> </li>
            <li><a href="http://excelerator.solutions/2017/08/28/excel-http-get-request/">http://excelerator.solutions/2017/08/28/excel-http-get-request/</a></li>
        </ul>
        <h3>IE vs Http</h3>
        <p>Peki, şuana kadar iki ana yöntem gördük. Ne zaman IE, ne zaman http kullanmak gerekiyor sorusunu sorabilirsiniz. Bunun için <a href="https://www.youtube.com/watch?v=R0xpDLzVcuw&amp;t=10s">şuarada</a> güzel bir karşılaştırma bulabilirsiniz. </p>
        <p>Özetle</p>
        <ul>
            <li>IE yöntemi ikinci bir uygulamanın(IE browserı) açılmasına neden olduğu için hem bellek yönetimi hem de süre açısından daha deazavatanjlıdır. Her ne kadar &quot;visible=false&quot; desek de uygulama arka planada açılacaktır. Özetle xmlhttp çok daha hızlıdır.</li>
            <li>IE ile sadece HTML işleyebilirsiniz. XML/json sonuçlarını işleyemezsiniz</li>
            <li>Parametreler sözkonusu ise IE ile bunları da URL&#39;e eklemiş oluruz, ki http de GET metodunu kullanmakla aynı şey. Ancak hassas bilgi geçirmemiz gerektiğinde bunu yapmamalıyız. Zira http konusunu araştırdıysanız görmüşsünüzdür ki, gönderdiğiniz bu bilgiler sunucuya gidene kadar başka bilgisayarladan da geçer ve oralarda depolanıabilir, yani bi güvenlik sorununa neden olur. (Bunun için IE yöntemini de uygun şekilde ayarlamak mümkün ama işleri uzamıtş olursunuz)</li>
            <li>Otantikasyon gereken durumlarda IE yine yetersizdir, aslında bunda da SendKeys gibi yöntemlelre destek alınabilse de süreci uzatmaktan başka işe yaramaz.</li>
            <li>IE artık ölmekte üzere olan bir browser, MS tarafından geliştirilmesi de önemsenmiyor. O yüzden şimdi veya ileride desteklemediği bazı kodlar olabilir.</li>
        </ul>
        <p>
            Peki bu durumda IE&#39;nin pek de avantajlı olduğu bir durum yok gibi görünüyor. Ama öyle değil, onun da kullanılabileceği caseler olabilir.</p>
        <ul>
            <li>İlgili sitenin bir veriyi size sunması için bir javascript kodu çalıştırması gerekebilir. Burda xmlhttp nesnesi işe yaramaz, IE kullanılabilir. (Bu arada VBA WebTools, Web Browser veya Selenium for VBA gibi 3rd party başka kütüphaneler de var, ama biz bunlara hiç girmedik, bunlar da bu konularda iş görür)</li>
            <li>Bahsekonu veri bir frame içindeyse IE terhcih edilir, xmlhttp ile biraz daha <a href="https://stackoverflow.com/questions/48720180/unable-to-parse-some-links-lying-within-an-iframe">teferruatlıdır</a></li>
        </ul>
        </div>
        <h2 class="baslik">Web servisler, API&#39;ler ve Json </h2>
        <div class="konu">
        <p>XML eskiden çok poüler olmakla birlitke yerini artık gittikçe Json&#39;a bırakmakta. Bunlar bazen karşımıza json uzantılı bir dosya olarak çıkabilmekte, bazen de gittiğimiz link bize sonucu json formatında bir output şeklinde vermektedir. Bazen de bir web service&#39;in sonucu olarak json output görebilmekteyiz. Bununla birlikte webservisler daha çok XML çıktı verirken, kısmen json da verebilmektedir. RESTful servisler ise ağırlıklı olarak json çıktı üretirler. Biz, yukarda XML&#39;i bolca gördük. Şimdi öncelikle biraz jsona bakalım, sonra da web servislere.</p>
        <h3>Json</h3>
        <p>Json, gerek VBA gerek diğer programlama dillerinde dictionarylere denk gelen bir veri formatıdır. Bu bazen basit bir dictionary iken bazen çok kompleks formatlarda karşımıza çıkabilmekte. Dictionary of dictionary, dictionary of dicitonary of collection, dictionary of collection of dictionary v.s. Bunları parse etmek için XML&#39;de olduğu gibi hazır bir kütüphanemiz ve sınıfımız yok. Ancak çok güzel hazırlanmış ve oldukça popüler olan bir <a href="https://github.com/VBA-tools/VBA-JSON">modül</a> var, ki bunu kullanabiliriz. Nasıl kurulacağı ilgili linkte anlatılıyor.</p>
            <p>Bu arada <a href="https://codebeautify.org/jsonviewer">https://codebeautify.org/jsonviewer</a> diye bi ste var, elinizdeki karışık ve formatsız json verisini formatlı hale getiriyor. Şimdi benim hazırladığım birkaç json örneği üzerinden farklı kurgulardaki örnekleri inceyelim</p>
            <h4>Dictioanry of dictionary</h4>
            <p>Bu örnekteki json <a href="https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis.json">metnine</a> önce kendiniz browserda bakın. Sonra bu kodu çalıştırın.</p>
        <pre class="brush:vb">
Sub basicjson()

Dim httpObject As New MSXML2.XMLHTTP60
Dim oJSON As Dictionary, innerdict As Dictionary

sURL = "https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis.json"

httpObject.Open "GET", sURL, False
httpObject.send
sGetResult = httpObject.responseText

Set oJSON = JsonConverter.ParseJson(sGetResult) ' dict of dict
  
For Each d In oJSON
    Debug.Print (d)
    Set innerdict = oJSON(d)
    For Each x In innerdict
        Debug.Print innerdict(x)
    Next x

Next d
End Sub</pre>
        <h4>Dictioanry of colleciton of collection </h4>
            <p>Bazen oluşturulan liste <a href="https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis_split.json">bu linkteki</a> gibi olabilir. O zaman kodumuzda küçük bi değişklik yaparız. Zira bu sefer elde ettiğimz nesne, bir dictioanry of collection of collection oluyor.</p>
        <pre class="brush:vb">
Sub basicjson2()

Dim httpObject As New MSXML2.XMLHTTP60
Dim oJSON As Dictionary, coll As Collection

sURL = "https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis_split.json"

httpObject.Open "GET", sURL, False
httpObject.send
sGetResult = httpObject.responseText

Set oJSON = JsonConverter.ParseJson(sGetResult) ' collection of dict
Debug.Print (TypeName(oJSON)) 'dictionary
Set coll = oJSON("data")
For Each Item In coll
    'Debug.Print (TypeName(Item)) 'collection
    Debug.Print Item(1)
Next Item
End Sub</pre>
            <h4>
                Dictionary of collection of dictionary</h4>
            <p>
                Bazen de
                <a href="https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis_table.json">şu şekilde</a> hazırlanmış bir json dosya olabilir. O zaman eldeki yapıya göre de kodumuzu değiştiririz.</p>
        <pre class="brush:vb">
Sub basicjson3()

Dim httpObject As New MSXML2.XMLHTTP60
Dim oJSON As Dictionary, coll As Collection, innerdict As Dictionary

sURL = "https://raw.githubusercontent.com/VolkiTheDreamer/dataset/master/json/indentli_bolgesatis_table.json"

httpObject.Open "GET", sURL, False
httpObject.send
sGetResult = httpObject.responseText

Set oJSON = JsonConverter.ParseJson(sGetResult) ' collection of dict
'Debug.Print (TypeName(oJSON)) 'dictionary
Set coll = oJSON("data")
For Each innerdict In coll
    'Debug.Print (TypeName(innerdict)) 'dictionary
    Debug.Print innerdict("Bolge")
Next innerdict
End Sub            &nbsp;</pre>
        <h4>
            Collection of dictioanry of dictioanry</h4>
            <p>
            Şuana kadarkilerin hepsinde aslında kök yapı olarak bir dictionary vardı, genelde 
                de durum böyledir ancak bazen gittiğimiz yerde dinamik olarak oluşan dosya bir json dönmek yerine &quot;array of json&quot;, veya VBA dilinde söyleyecek olursak kökünde collection olan bir yapı(Ör:collection of dictionary) dönüyor olabilir. O zaman kurgumuzu ona göre yaparız. Dönecek json nesnesini collection olarak tanımlayalım, sonra her bir itemı dictionary gibi kullanalım.
&nbsp;</p>
        <pre class="brush:vb">
Sub basicjson4()

Dim httpObject As New MSXML2.XMLHTTP60
Dim oJSON As Collection, dict As Dictionary

sURL = "https://www.excelinefendisi.com/httpapiservice/ReturnJson.aspx"

httpObject.Open "GET", sURL, False
httpObject.send
sGetResult = httpObject.responseText

Set oJSON = JsonConverter.ParseJson(sGetResult) ' collection of dict
'Debug.Print (TypeName(oJSON)) 'collection
For Each dict In oJSON
    'Debug.Print (TypeName(dict)) 'Dictionary
    For Each key In dict
        Debug.Print dict(key)
    Next key

Next dictEnd Sub   &nbsp;</pre>
        <h4>Diğer json örnekleri</h4>
            <p>Diğer senaryolar , kendi pc&#39;nizdeki bir json dosyayı okumak, exceldeki bir veriyi json olarak kaydetmek olabilir. Bunlar için bu json modülünün sayfasındaki örneklerle stackoverflow gibi sitelere bakabilirsiniz.</p>
            <h3>WebService ve API</h3>
            <p>Webservice ve API kavramlarına ve bunlar arasındaki farklara uzun uzadıya değinmeyeceğim, bunları bildiğinizi varsaycağım, bilmiyorsanız biraz araştıramanızda fayda var. </p>
            <p>Özünde her webservice&nbsp;bir apidir ama her api bi webservice değidlir. API&#39;ler genelde belli bir programlama dilinde kullanmak üzere yazılırlar. Webserviceleri ise doğrudan kullanabilirsiniz(tüketebilirsiniz). Bildiğim kadarıyla VBA için yazılan public bir API yok. Aslında VBA kütüphanelerinin hepsi bir API&#39;dir, hatta VBA&#39;in kendisi de Excel için bir API^&#39;dir, zira Excel aslında c++ ile yazılmış bir program olup, onu kontrol etmek, onun obje modeline ulaşmak için eklenmiş bir API&#39;dir. Public API&#39;lerden kastım spotify, amazon, twitter v.s gibi sitelerin API&#39;leri VBA&#39;de yok demek istiyorum. Ama benim sitem dahil birçok public sitenin web serivisleri VBA ile rahatça tüketilebilir.</p>
            <p>Peki çok teknik detaya girmeden, VBA ile web servisleri nasıl kullanırız ona bakalım. Öncelikle yine XML ve JSON&#39;da olduğu gibi VBA-Toolsun bu konuda da hazırlanmış bir <a href="https://github.com/VBA-tools/VBA-Web">modülü</a> var ama bana biraz karışık geldi, çok kullanamadım.&nbsp;Bir de Micorsoftun <strong>webtoolkit</strong> diye de paketi var, o da çok eskiymiş sanırım, onu da çok kullanmadım. Bana çok da gerekli gelmediler açıkçası, zira eldeki mevcut yöntemlerle bu servisleri tüketebiliriz, en azında bu aşağıda vereceğim benim kendi sitemin servisleri için.</p>
            <h4>Excel&#39;in Yerleşik Fonksiyonları: WEBSERIVCE ve FILTERXML(Office365) </h4>
            <p>Öncelikle Excel&#39;in WEBSERVICE ve FILTERXML fonksiyonlarından bahsetmek isterim. </p>
            <p>WEBSERVICE fonksiyonu ile ya bi webservice linkini çağırabilir ya da doğrudan XML döndüren bi sayfayı çağrıbilirsiniz.(json olanı da getirir ama bu bizim için çok kullanışlı olmaz)</p>
            
            <ul>
                <li><pre class="formul">=WEBSERVICE(&quot;https://www.excelinefendisi.com/httpapiservice/WebService.asmx/GetSorular?tag=sheet&quot;) //xml döndüren bi webservice</pre></li>
                <li><pre class="formul">=WEBSERVICE(&quot;https://www.excelinefendisi.com/Sitemap.xml&quot;)&nbsp; //xml döndüren bi link(webservice değil)</pre></li>
            </ul>
                
            <p>Bunlar bize ilgili çıktıyı XML formatında gösterir. Ama bunu Excel&#39;in hücrelerine satır satır yazdırmak için bir de FILTERXML fonksiyonunu, uygun kısım için XPATH vererek yazarız. Ve bu fonksiyon <a href="../Excel/FormulasMenusuFonksiyonlar_DiziFormulleriveSumproduct.aspx#dinamikdizi">dinamik dizi</a> ürettiği için sonuçlar aşağı doğru dökülür. Yukarıdaki linkler için sırayla şu formülleri yazmak yeterlidir.(A1 ve A2&#39;de WEBSERVICE formülleri olduğunu düşünün)</p>
            
            <ul>
                <li><pre class="formul">=FILTERXML(A1;&quot;/ArrayOfString/string&quot;)</pre>
                </li>
                <li><pre class="formul">=FILTERXML(A2;&quot;//urlset/url/loc&quot;)</pre>
                </li>
            </ul>
                
            <p>Ancak batch işlem yapackasanız yani otomasyon ihtiyacınız varsa, tabiki VBA yine imdada yetişecektir. Şimdi VBA&#39;de bunları nasıl yapıyoruz ona bakalım.</p>
            <h4>VBA ile WebService</h4>
            <p>Ben burada yine, kendi stemdeki webservis örneklerini kullanacağım. Benim servislerim eski tipteki asmx uzantılı webseviselerdir, bunun yerine artık daha çok Web API kullanılıyor. O yüzden çözüm şeklinde farklılıklar olabilir, emin değilim.(Belki bi ara WEB API de oluştururum, o zaman burayı da güncellerim)</p>
            <p>Öncelikle yine uygun library&#39;yi(SOAP) ekleyerek başlıyoruz: <strong>Microsoft Office Soap Type Library</strong>. Reference menüsünde bunu bulamazsanız Browse deyip &quot;<a href="file:///C:/Program%20Files/Common%20Files/Microsoft%20Shared/Office12/MSSOAP30.dll">C:\Program Files\Common Files\Microsoft Shared\Office12\MSSOAP30.dll</a>&quot; adresinden bulmaya çalışın. Web servicler her ne kadar XML dışında (Ör.json) sonuç döndürebilse de SOAP teknolojisi sadece XMLle çalışır.</p>
            <p>Web servicelerinin WSDL denen, hangi metodları hangi parametrelerle kullanabileceğinizi gösteren tanım sayfaları vardır. Benimkine <a href="https://www.excelinefendisi.com/httpapiservice/WebService.asmx?WSDL">buradan</a> ulaşabilirsiniz. Bu karışık bi görünt olabilir ama VBA kodu içinde bunu kullnacağız. İnsan gözü için daha okunaklı olan ve sadece metodların bi listesini gösteren sayfayı görmek için sondaki &quot;?WSDL&quot;i kaldırmayı deneyin. </p>
            <p>Şimdi kodumuza bakalım.</p>
            <pre class="brush:vb">
Sub soap_webservice()
    Dim ws As New MSOSOAPLib30.SoapClient30 'Sadece XMLle çalışır
    Dim xmldoc As New MSXML2.DOMDocument60
    Const c_WSDL_URL  As String = "https://www.excelinefendisi.com/httpapiservice/WebService.asmx?WSDL"

    ws.MSSoapInit c_WSDL_URL

    'ÖNEMLİ: intellisensete bu metodlar doğal olarak görünmez, web servicein WSDL'inden hangi metodların bulunduğuna bakmanız lazım        
    Debug.Print ws.HelloWorld 
    Debug.Print ws.WriteDuyurularAsJsonIndented 'hata verir, çünkü sonuç XML değildir
    Debug.Print ws.GetDuyurularAsJsonIndented 'hata vermez, çünkü sonuç XML içine paketlenmiş jsondır
    Debug.Print ws.GetDuyurularAsJson
    Debug.Print ws.GetSorular("?tag=sheet") 'hata: bu kütüphane arraylerden oluşmuş XML yapılarını(arrayofstring) desteklemiyor. Çözüm, bunu VSTO'da .Net ile yazıp, VBA içine COM obje olarak almak olabilir'
    Debug.Print ws.GetDuyurularAsXMLSimple
    Debug.Print ws.GetDuyurularAsXMLFormatted
    
    'şimdi bunlardan birini XML olarak ele alalım.
    xmldoc.LoadXML (ws.GetDuyurularAsXMLFormatted)
    
    Set Nodes = xmldoc.DocumentElement.ChildNodes(1).FirstChild.ChildNodes
    For Each Node In Nodes
        Debug.Print Node.ChildNodes(1).Text
    Next Node
    
End Sub     </pre>
            <p>Aslında bütün bu işi xmlhttprequest nesnesi ile de yapabilirdik. Mesela yukarıda desteklenmeyen arrayofstring elementi bu yöntemle sorunsuz şekilde elde edilebilir.</p>
            <pre class="brush:vb">
Sub http_webservice()

Dim request As New MSXML2.XMLHTTP60
Dim root As IXMLDOMElement
Dim snodes As IXMLDOMNodeList

url = "https://www.excelinefendisi.com/httpapiservice/WebService.asmx/GetSorular?tag=sheet"
request.Open "GET", url, False
request.send

Set root = request.responseXML.DocumentElement
Set snodes = root.ChildNodes

For Each s In snodes
    Debug.Print s.Text
Next s
End Sub</pre>
        <p>&nbsp;</p>
    </div>
</asp:Content>
