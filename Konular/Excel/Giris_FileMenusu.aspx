<%@ Page Title='Giris FileMenusu' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>

<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>

<h1>File Menüsü</h1>
<p>File menüsü, işimizi yaparken rahat çalışmamızı sağlayan Seçenekler menüsünü ve MIS/Raporlama v.s alanlarında çalışan kişilerin arada bir de olsa ihtiyacı olan birkaç önemli özelliği içerir.</p>

<p>Bunları iki grupta toplayabiliriz.
<ul>
  <li>Dosya koruma</li>
  <li>Options(Seçenekler)</li>
</ul></p>

Aşağıdaki bölümlerde bunları detayını bulabilirsiniz.
<h2 class='baslik'>Dosya Koruma</h2>
<div class='konu'>

<h2>Koruma ve Erişim Kısıtlama</h2>
<h3>Dosya şifreleme</h3>
<p>Diyelim ki, çok gizli bilgileri içeren bir dosyanız var, ve maille bir yere göndermeniz gerekiyor, ancak bir şekilde bu mail başkalarının eline geçebilir diye de rahat değilsiniz. Bu durumda, Info alt menüsündeki <span class=" keywordler">Encrypt with Password</span> butonu ile dosyamıza bir şifre koyarız ve maille göndereceğimiz kişiye de telefonla şifreyi söyleriz. Bu kadar basit.

<img src="/images/FileEncrypt.jpg" alt="File Encrypt"/>
</p>

<h3>Erişim sınırlama</h3>
<p>Şimdi de diyelim ki, oluşturduğunuz bir dosyayı bir alıcı grubuyla paylaşmak istiyorsunuz ancak, kimsenin bunu print etmesini istemiyorsunuz, veya üzerinde bir değişiklik yapmalarını da istemiyorsunuz. Böyle durumlarda yine Info alt menüsündeki <strong>Restrict Access</strong> butonunu kullanabilirsiniz. </p>

<img src="/images/FileRestrictPermission.jpg" alt="File Restriction"/>
</div>




<h2 class="baslik">Options(Seçenekler)</h2>
<div class="konu">
<p>
Options alt menüsüne tıklandığında, default olarak seçilmiş özelliklere çok değinmeyeceğim, bunlar zaten hayatınızda olduğu için farkında olmadan kullandığınız özellikler olabilir. Seçilmemiş seçeneklerden, özellikle hız kazandıracak özelliklerden ise bahsetmekte fayda var. Kısaca bi bakalım:
</p>

<ul>
<li>
<p><span class=" keywordler">General>User Interface Options>Show Mini Toolbar on selection</span> seçeneği: Belirli bir alanı seçtiğinizde seçimin hemen bitiminde bir kutucuk belirir, ona tıkladığınızda bu seçimle ilgili neler yapabileceğinizi gösteren bir kutucuk daha açılır. İster toplam alır, ister conditional format uygularsınız, v.s. Aslında tüm bunları seçimi yaptıktan hemen sonra Ribbondan da yapabilirsiniz, ancak bu küçük kutucuk sayesinde amacınıza daha kısa sürede ulaşmış olursunuz.</p>

<img src="/images/FileOptionstoolbar.jpg" alt="File Options"/>
</li>

<li>
<p>Bir diğer önemli seçenek de Özet(Pivot) Tablolarla çalışırken karşımıza çıkıyor. <span class=" keywordler">Formulas>Working with formulas>Use GetPivotData functions for PivotTable references</span> seçeneği: Bu seçenek işaretliyse bunu kaldırmanızı tavisye ederim. Zira genelde baş belası olmaktan başka işe yaramamaktadır. Bu özelliği kullanan var mıdır bilmiyorum ama ben sık sık özet tablolardaki belirli hücrelerden beslenen formül yazma gereği duyarım ve formülümün oldukça sade görünmesini isterim. Aşağıda resimlerden ne demek istediğimi anlayacaksınız.</p>

<img src="/images/FileOptionsFormulaPivot2.jpg" alt="Seçenek işaretli hali" />
<p class="ortala"><strong>Resim1. Seçenek işaretli hali</strong></p>

<p>Gördüğünüz gibi formülü bu haliyle aşağı indirmeniz pek de kolay olmayacaktır. Bu yüzden bu seçeneği iptal etmemiz hayrımıza olacaktır. Ben açıkçası GetPivotData formülünü kullanma ihtiyacını çok fazla duymadım. Yine de bi gözatmak isterseniz, <a href="http://www.contextures.com/xlPivot06.html">şu linkte</a> örnekler var. Bu arada olur da kullanmanız gerekirse bu menüye girmeden bu özelliği aktive edebiliyorsunuz. Bunu <a href="/Konular/Excel/Insertmenusu_pivottable.aspx">Özet Tablolar</a> konusunda detaylıca göreceğiz.
</p>

<img src="/images/FileOptionsFormulaPivot1.jpg" alt="Seçenek işaretsiz hali" />
<p class="ortala"><strong>Resim2. Seçenek işaretsiz hali</strong></p>


</li>

<li>
<p>Yine önemli bir seçenek de <span class=" keywordler">Advanced>When calculating this workbook>Update links to other documents</span> seçeneği: Bunu duruma göre açık duruma göre kapalı bırakmanız gerekebilir. Ben genelde şöyle yapıyorum; eğer birkaç tane yoğun formül içeren dosyayı ve bunların linkli olduğu dosyayı açacaksam bu seçeneği geçici olarak kapatırım, onun dışında hep işaretli durur.</p>
</li>

<li>
Options'ta Ribbon ve Add-in/Makro ayarları da bulunuyor, ancak bunlara ayrı sayfalarda değiniyor olacağım.
</li>

</ul>


</div>

</asp:Content>
