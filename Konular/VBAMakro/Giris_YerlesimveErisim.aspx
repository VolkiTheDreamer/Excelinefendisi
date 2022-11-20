<%@ Page Title='Giriş YerlesimveErisim' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'>
<table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>

<h1>Yerleşim ve Erişim</h1>

<p>Yazdığm kodlar nerede gözükecek, bunları nasıl ve nerede çalıştıracağız gibi sorularınızın olması muhtemeldir. Bu konuyu, bu sorularınıza cevap için hazırladım.</p>
        <h3>Kodların yazılacağı yerler</h3>
<p>İlk örneklerde, kodlarımızı hep Personal.xlsb içindeki Modüller içine yazdık. Peki başka yere kod yazamaz mıyız? Tabiki yazarız.</p>
	<ul>
		<li>Mesela, Dosyaların kendisine ait bir kod bölümü var, VBE editöründe 
		ThisWorkbook içine gider. Bu konu
		<a href="Olaylar_WorkbookOlaylari.aspx">şurada</a> ele alınacaktır</li>
		<li>Sayfaların da kendine ait kodları olabilir. Bu konu
		<a href="Olaylar_WorksheetOlaylari.aspx">şurada</a> ele alınacaktır</li>
		<li>Bir butona tıklandığında bir kod çalışmasın sağlayabiliriz ancak bu da 
		Modül seviyesinde ele alınır.</li>
		<li>Bir UserForm oluşturulabilir(bu konu ayrıca
		<a href="Formlar_Konular.aspx">burada</a> ele alınacak)</li>
	</ul>
    <h3>Kodların çalıştırma yöntemleri</h3>
	<p>Yazdığınız kodların çalışmasını sağlamanın da birkaç yolu var. Yukardaki 
	maddelerle bağlantılı olarak;</p>
	<ul>
		<li>VBE açıkken ve bir prosedürün içindeyken F5 ile</li>
		<li>WB veya WS ile ilgili bir olay gerçekleştiğinde kendiliğinden devreye 
		girecek event bazlı kodlar</li>
		<li>Ribbona veya QuickAccesbara atadığınız butonlara tıkladığınızda 
		çalışacak kodlar</li>
		<li>Add-in olarak hazırladığınız kodlar</li>
		<li>Sayfa üzerinde bir butona bastığınızda çalışacak kodlar</li>
		<li>"Macros" dialog kutusu (Alt+F8)</li>
		<li>Kısayol(Shortcut) atadığınız kodlar</li>
	</ul>
	<p>Bunların hepsini yeri geldikçe göreceğiz, burada sadece Ribbon'a ve 
	QAT'ye düğme nasıl eklenir ona bakacağız.</p>


	<h3>Ribbon veya QAT'ye makro düğmesi atama</h3>
	<p>İşlemler her ikisi için de aynı olacağı için ben sadece QAT üzerinden 
	anlatacağım.</p>
	<p>QAT'a sağ tıklayarak özelleştir diyelim. Sonra menüden "Macros"u seçip, 
	aşağıdan da istediğimiz makroyu Add düğmesine tıklarayak QAT'de istediğimiz 
	yere alalım.</p>
	<p><img src="../../images/QATbuton1.jpg">&nbsp;</p>
	<p>Düğmemiz eklendikten sonra Modify tuşuna basarak ikonu ve makronun 
	görünen ismini istediğimiz gibi değiştirebiliriz.</p>
	<p><img src="/images/QATbuton2.jpg" height="315" width="281"></p>
	<p>OK dedikten sonra düğmemizin QAT'ye eklendiğini görürüz. Personal.xlsb 
	üzerindeki bir makroyu eklemişsek, ki genelde öyle yaparız, bu makro tüm 
	dosyalarda çalışır halde olacaktır.</p>
	<p><img src="/images/QATbuton3.jpg" height="84" width="670"></p>
</asp:Content>
