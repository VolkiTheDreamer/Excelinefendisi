<%@ Page Title='Giris DosyaUzantilari' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>


<h1>Dosya Uzantıları</h1>
<p>Excel'in desteklediği birçok dosya türü olmasına rağmen burada sadece üçünden bahsedeceğim. Bu arada dosya türü ile dosya uzantısı genelde birebir örtüşmekle birlikte bazen bir uzantı birden fazla türü kapsayabilmektedir, ama bu detaya burada girmeyeceğim.</p>

<p>Bizim için önemli olan dosya türleri şunlardır.</p>

<ul>
<li>Standart format:<strong>xlsx</strong></li>
<li>Makrolu format:<strong>xlsm</strong></li>
<li>Binary(ikili) format:<strong>xlsb</strong></li>
</ul>

<h3>Standart Format(*.xlsx)</h3>
<p>Excel, 2007 versiyonunudan itibaren XML destekli dosya tipine geçmiştir. XML hakkında küçük bir araştırma yapmanızı tavsiye ederim. Bu format farklı platformlar arasında standart veri taşıma formatı(başka standartlar da var, şimdilik en yaygını bu) olarak adlandırılabilir. Bu uzantılı dosyalar Excel'in eski standart uzantısı olan xls uzantısına göre daha küçük hacimde yer tutar, çünkü arka planda bir ZIP sıkıştırma sistemi çalışır. </p>

<p>Eskiden(2007 öncesi) internetten indirdiğiniz, veya bir şekilde size gelmiş ama içeriğinin ne olduğunu bilmediğiniz bir Excel dosyası güvenlik problemi teşkil edebiliyordu, çünkü xls uzantılı dosyalar içinde kötü amaçlı makrolar kaydedilebilirdi. Artık bu korkuya yer yok. Çünkü xlsx uzantılı bir dosya içinde makro yer alamaz. Dolayısıyla gerek bu siteden, gerek başka sitelerden indireceğiniz xlsx uzantılı dosyaları güvenle açabailirsiniz.</p>

<h3>Makrolu Format(*.xlsm)</h3>
<p>Az önce güvenlik nedenlerinden ötürü, xlsx uzntılı bir dosya içine makro kaydedemez, VBA kodu yazamazsınız demiştik. Bunun için dosyanızı standart makro uzantısı olan xlsm yapmanız gerekmektedir. Bu uzantı türü de XML tabanlıdır. </p>

<h3>Binary(İkili) Format(*.xlsb)</h3>
<p>xlsb uzantısı, diğer iki format gibi XML tabanlı olmayıp, binary formattadır ve bu sayede bilgiyi daha küçük hacimde tutar. Bu dosya türü, özellikle büyük hacimli dosyalarda kullanışlıdır. Lokalde kullandığınız, web üzerinden bağlantısı olmayan veya bir şekilde başka platformlara gönderilme durumu olmayan büyük hacimli dosyalarınızı xlsb olarak kaydetmenizi önerebilirim. 
Hatta, <span class="keywordler">File&gt;Options&gt;Save</span> menüsünden varsayılan 
dosya kaydetme uzantısını xlsb olarak belirleyin.</p> 

<p>xlsb uzantılı dosyalar da xlsm gibi makroları desteklemektedir. Makrolar bölümünde göreceğiniz üzere "Personal" dosyasının uzantı seçiminde xlsm'ye göre tercih edilmektedir.</p>
	<p>Ama yukarıda belirttiğim gibi, bu uzantı tipinin nerede sıkıntı 
	yaratacağını öngöremeyebilirsiniz, o yüzden genelgeçer bir çözüm haline 
	getirmeden önce test etmenizde fayda var. Ör: Birçok telefonda bu uzantılı 
	dosyalar açılamıyor. Birçok programlama dili bu uzantılı dosyaları 
	okuyamıyor. Örnekler çoğaltılabilir.</p>

<h3>Karşılaştırma</h3>
<a href="http://stackoverflow.com/questions/7821632/when-should-the-xlsm-or-xlsb-formats-be-used">Stackoverflow</a> sitesinde bir soruya verilen cevapta bir karşılaştırma yer alıyor. Buna göre .xlsx uzantılı dosyalar xlsb uzantılılara göre 4 kat daha uzun sürede açılıyor, 2 kat daha yavaş kaydoluyor, ve 1,5 kat daha çok yer kaplıyor. (Detaylı tablo aşağıda olup karşılaştırmayı yapan kişinin PC donanımına ve dosya büyüklüğüne göre bu verilerin değişeceği aşikardır, ancak oran çok fazla değişmeyecektir)

</br>
</br>

<table class="alterantelitable">
<th>Karşılaştırma Konusu</th>
<th>xlsx</th>
<th>xlsb</th>


<tr>
<td>Dosya açılma süresi</td>
<td>165sn</td>
<td>43sn</td>

</tr>

<tr>
<td>Dosya kaydetme süresi</td>
<td>115sn</td>
<td>61sn</td>

</tr>

<tr>
<td>Dosya boyutu</td>
<td>91MB</td>
<td>65MB</td>

</tr>
</table>


</asp:Content>
