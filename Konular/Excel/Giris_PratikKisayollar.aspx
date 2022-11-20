<%@ Page Title='Giris PratikKisayollar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>

<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>

<h1>Pratik Kısayollar</h1>
<p> Bu bölümde gerek klavye kısayollara, gerek klavye kullanımı dışındaki bazı numaralara değineceğim</p>


<h2 class="baslik">Klavye Kısayolları</h2>
<div class="konu">

<table class="alterantelitable">
<th>Amaç</th>
<th>Kısayol</th>

<tr>
  <td>Sayfalar arasında dolaşmak</td>
  <td>CTRL + PgUp/PgDn</td>
</tr>

<tr>
  <td>Bugünün Tarihini yazmak</td>
  <td>CTRL + SHIFT +, </td>
</tr>

<tr>
  <td>Tüm açık dosyalarda calculation yapmak</td>
  <td>F9</td>
</tr>

<tr>
  <td>Seçili kısmın değerini hesaplayıp göstermek</td>
  <td>Hücre içindeki formül seçilip F9</td>
</tr>

<tr>
<td>Aktif sayfada calculation yapmak</td>
<td>SHIFT+F9</td>
</tr>

<tr>
<td>Sadece belli range için calculation yapmak</td>
<td>VBA ile yapılır. <a href="/Konular/VBAMakro/DortTemelNesne_Range.aspx#Calculation">Burdan </a>bakın.</td>
</tr>

<tr>
<td>Bulunduğun hücrenin <abbr title="Bulunulan hücrenin etrafındaki tüm dolu alandır">CurrentRegion</abbr>'ını seçme</td>
<td>CTRL+ A</td>
</tr>

<tr>
<td>Bulunduğun hücreden <abbr title="Bulunulan hücrenin etrafındaki tüm dolu alandır">CurrentRegion</abbr>'ın uç noktlarına gitmek</td>
<td>CTRL+ Ok tuşları</td>
</tr>

<tr>
<td>Bulunduğun hücreden itibaren belli bir yöne doğru seçim yapmak</td>
<td>SHIFT+Ok tuşları</td>
</tr>

<tr><td>Bulunduğun hücreden itibaren <abbr title="Bulunulan hücrenin etrafındaki tüm dolu alandır">CurrentRegion</abbr> bir ucuna doğru toplu seçim yapmak</td>
<td>CTRL+SHIFT+Ok tuşları</td>
</tr>

<tr>
<td>Bulunduğun hücreden <abbr title="Bulunulan hücrenin etrafındaki tüm dolu alandır">CurrentRegion</abbr>'ın Sağ Aşağı uç noktlasına gitmek</td>
<td>CTRL+END</td>
</tr>

<tr>
<td>Bulunduğun hücreden <abbr title="Bulunulan hücrenin etrafındaki tüm dolu alandır">CurrentRegion</abbr>'ın Sağ Aşağı uç noktlasına kadar seçmek</td>
<td>CTRL+SHIFT+END</td>
</tr>

<tr>
<td>Bulunduğun hücreden  A1 hücresine kadar olan alanı(sol yukarı) seçmek</td>
<td>CTRL+SHIFT+HOME</td>
</tr>

<tr>
<td>Bir hücre içinde veri girerken, aynı hücre içinde yeni bir satır açıp oradan devam etmek</td>
<td>ALT+ENTER</td>
</tr>

<tr>
<td>Veri/Formül girişi yaptığınız hücrede alt hücreye geçmeden giriş tamamlamak </td>
<td>CTRL+ENTER</td>
</tr>

<tr>
<td> Ekranda bir sayfa sağa kaymak.</td>
<td>ALT+PGE DOWN</td>
</tr>

<tr>
<td>AutoFilter'ı aktif/pasif hale getirmek</td>
<td>CTRL+SHIFT+L</td>
</tr>

<tr>
<td>Bulunduğunuz hücrenin satır ve sütununa aynı anda freeze uygulamak/kaldırmak</td>
<td>Alt+W+FF</td>
</tr>

<tr>
<td>VBA editörünü açmak</td>
<td>Alt+F11</td>
</tr>

<tr>
<td>Ribbonu küçültüp/büyütmek</td>
<td>CTRL+F1</td>
</tr>

<tr>
  <td>Üst hücrelerdeki tüm rakamların toplamını almak</td>
  <td>ALT+=</td>
</tr>

<tr>
  <td>Flash Fill uygulamak</td>
  <td>CTRL+E</td>
</tr>

<tr>
  <td>Sadece görünen hücreleri seçmek</td>
  <td>ALT+;</td>
</tr>

</table>

<p>Bunların dışında ayrıca, kendi kısayollarınzı da yaratabilirsiniz, ben mesela CTRL+M kombinasyonuyla PasteSpecial:=Values yapıyorum. Bunun gibi birkaç kısayolum mevcut. Bunları VBA/Makrolar sayfalarında ele alacağım. Gözatmak isterseniz 
<a href="../VBAMakro/Giris_MakroKaydetmeveVBE.aspx">buradan </a>bakabilrsiniz.</p>

<p>Son olarak daha kapsamlı bir liste görmek isterseniz, şu linklere de göz atmak isteyebilirsiniz</p>

<ul>
<li><a href="https://www.shortcutworld.com/en/win/Excel_2016.html">Excel 2016 Shortcuts(En kapsamlısı bu, Exclein tüm versiyonlarına ait sayfaları da mevcut)</a></li>
<li><a href="https://support.office.com/en-us/article/Excel-shortcut-and-function-keys-1798d9d5-842a-42b8-9c99-9b7213f0040f?ui=en-US&rs=en-US&ad=US&fromAR=1">Miscrosoft'un destek sitesi(en son 2007ye göre update etmişler)</a></li>
<li><a href="http://www.asap-utilities.com/excel-tips-shortcuts.php">Excel-tips-shortcuts</a></li>
</ul>
</div>

<h2 class="baslik">Bir hücre grubunu belirli bir rakamla çarpma/bölme v.s</h2>
<div class="konu">
<p>Diyelim ki, bir hücredeki veriyi başka hücrelerde bulunan rakamlarla çarpmak istiyorsunuz. Örneğin elinizde 1000'e bölünmüş rakamlar var, boş bir hücreye 1000 yazıp bunu 1000le çarpmak istediğiniz rakamların bulunduğu alanı seçip sağ tıklayın, <strong>Copy>Paste Special</strong> dedikten sonra Paste alanı altında Value, Operation altından da Multiply'ı seçip OK diyin. Bu kadar basit. </p>

<p>Önce</p>
<img src="/images/excelquickislem1.jpg">
<p>İşlemi yapıyoruz</p>
<img src="/images/excelquickislem2.jpg">
<p>Sonra</p>
<img src="/images/excelquickislem3.jpg">

<p>Tabi bunun bir başka(hızlı ve pratik) yolu da <a href="/Excelent.aspx">Excelent</a> altında bulunan Hızlı İşlem menüsündeki bu görevi gören makroyu kullanmaktır.</p>
</div>

<h2 class="baslik" id="ozelbul">Özel bulma</h2>
<div class="konu">
<p>Diyelim ki, bi workbooktaki formülleri bulmak ve bunlarla ilgili çeşitli formatlama işlemi yapmak istiyorsunuz. Bunun birkaç yolu var.</p>
<ol>
<li><span class=" keywordler">Formulas Menüsü>Formula Auditing>Show Formula</span>: Burada sadece formülleri gösterir, onları sizin adınıza seçmez</li>
<li>Excel 2013 ile gelen <span class=" keywordler">ISFORMULA(hucre)</span> formülü ile formül olan hücreleri tespit edip sonra bunları filtreyebilirsiniz.</li>
<li>Benim önereceğim yöntem ise <span class=" keywordler">Home menüsü>Find&Select>Formulas</span> düğmesidir. Buna tıkladığınızda tüm formüllü hücreler seçilir. Bundan sonrasında bu hücrelere ne yapacağınız size kalmış.</li>
</ol>
</p>
<img src="/images/kisayol_ozelbul.jpg"  alt="" />

<p>Find&Select menüsü altında başka seçenekler de var, bunlarda Go To Speacial'a gidip detaylı arama/seçme işlemleri yapabiliyorsunuz.</p>

<img src="/images/kisayol_find_gotospecial.jpg"  alt="" />

</div>

</asp:Content>
