<%@ Page Title='Giris Ribbon' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='5'></asp:Label></td></tr></table></div>

<h1>Ribbon</h1>
<p>2007 versiyonuyla hayatımıza giren Ribbonla ilgili olarak söyleyebileceğim iki konu var. Biri Ribbon'un özelleştirilmesi, diğeri de Quick Access Toolbar.
</p>

<h2 class="baslik">Ribbonun özelleştirilmesi</h2>
<div class="konu">
<p>Ribbonu özelleştirmenin üç tür yolu vardır. Bunu ya, sağ tıklayıp Custom Group ekleyerek yapabilirsiniz, ya <a href="/Konular/VSTO/Giris_Konular.aspx">VSTO</a> ile ya da XML tabanlı Custom UI Editorü ile.</p>

<p>Bu sitede  Custom UI Editor detayına girilmeyecek, VSTO konusuna yan menüden detaylıca bakabilirsiniz. Bu sayfada sadece Custom Tab/Group'tan bahsedeceğim. </p>

<p>Öncelikle neden Custom Tab/Group ihtiyacımız olur sorusunun cevabını bulalım. Diyelim ki, Excelin yerleşik menülerinin çoğunu kullanmıyorsunuz ve çok sık kullandığınız bazı butonlar var, bunları bir yerde gruplamak isteyebilirsiniz. Mesela sadece formatingle ilgili butonları tek bir grupta, data işleriyle ilgili butonları başka bir grupta toplamak isteyebilirsiniz. Gelin şimdi bunu nasıl yapacağımıza bakalım.</p>

<ul>
<li>Ribbon'a sağ tıklayıp Customize Ribbon diyin</li>
<li>New Tab diyin, ve Rename diyerek sekmenin adını değiştirin</li>
<li>New Group diyin, ve Rename diyerek grubun adını değiştirin</li>
<li>İhtiyaç duymadığınız menüleri kaldırın(Home ve Data)</li>
</ul>

<img src="/images/GirisRibbon.jpg" alt="Ribbon"/>

<p>Bütün bu adımlardan sonra yeni menümüz aşağıdaki gibi olacaktır</p>

<img src="/images/GirisRibbon2.jpg" class="zoomla" alt="Ribbon Özel"/>
<p>Ribbona veya aşağıdaki QAT'ye makro düğmesi atama hakkında bilgi için 
<a href="../VBAMakro/Giris_YerlesimveErisim.aspx">buraya</a> tıklayınız.</p>
	<p>Custom UI Editor ile ribbon tasarlamak istiyorsanız, aşağıdaki 
	bağlantıları incelemek isteyebilirsiniz.</p>
	<ul>
		<li><a href="https://www.rondebruin.nl/win/s2/win001.htm">
		https://www.rondebruin.nl/win/s2/win001.htm</a></li>
		<li><a href="https://www.rondebruin.nl/win/s2/win003.htm">
		https://www.rondebruin.nl/win/s2/win003.htm</a></li>
		<li><a href="https://www.contextures.com/excelribbonaddcustomtab.html">
		https://www.contextures.com/excelribbonaddcustomtab.html</a></li>
		<li><a href="https://powerspreadsheets.com/custom-excel-ribbon/">
		https://powerspreadsheets.com/custom-excel-ribbon/</a></li>
	</ul>
</div>


<h2 class="baslik">Quick Access Toolbar(QAT)</h2>
<div class="konu">
<p>QAT ise Ribbonda menüler arasında dolaşmakla uğraşmak yerine istediğiniz araca hızlıca ulaşmanızı sağlar. Buraya en sık kullandığınız yerel araçları koyabileceğiniz gibi makro atadığınız düğmeleri de koyabilirsiniz.</p>

<p>Aşağıda, benim iş bilgisayarımda kullandığım QAT'ı görebilirsiniz.</p>
<img src="/images/excelqat.jpg"  alt="QAT"  class="zoomla" />

</div>

</asp:Content>
