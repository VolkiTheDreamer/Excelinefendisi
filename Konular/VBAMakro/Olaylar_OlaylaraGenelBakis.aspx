<%@ Page Title='Olaylara Genel Bakış' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Olaylar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>
<h1>Olaylara Genel Bakış</h1><div class='konu'>
	<p> 	Eventlerle, biz Excele "şu olduğunda şu kodu çalıştır" demiş oluruz. 
	Burdaki "şu 
	kod" dediğimiz prosedürlere <strong>Event Handler </strong>prosedürleri denir.&nbsp;Bu 
	prosedürleri yazmak temel olarak normal bir prosedür yazmaya benzer fakat iki 
	farklı yönü vardır.</p>
	<ul>
		<li>Bunlar, kendileriyle ilişkili olan nesnenin modülün(Workbook, 
		worksheet) içine yazılır, 
		yani standart bir modül içine yazılmazlar. Bu kuralın 3 istisnası vardır<ul>
			<li>Standart bir modül içine konan nesne olmayan olaylar(<strong>OnTime</strong> 
			and <strong>OnKey</strong>)</li>
			<li>Application eventleri</li>
			<li>Class modüllerine yazılan Chart(sayfa içindeki gömülü 
			olanlardan) eventleri</li>
		</ul>
		</li>
		<li>Özel yazım syntaxları vardır.<ul>
			<li>Nesne adını takiben "_" işareti</li>
			<li>Eventin adı</li>
			<li>(Varsa argümanlar)</li>
			<li>Ör:Workbook_Open(), Worksheet_Change(ByVal Target As Range)</li>
		</ul>
		</li>
	</ul>
	<p> 	Tabi bunların özel yazım syntaxı var diyorum ama genelde bunları elle 
	yazmayız. Mesela Projects penceresinde <strong>ThisWorkbook </strong>seçiliyken(çift tıkla 
	seçilmesi gerekir) aşağıdaki resimde göründüğü gibi soldaki nesneler 
	ComboBox'ına tıklanıp <strong>General </strong>olan seçimi,</p>
	<p> 	<img src="../../images/event2.jpg"></p>
	<p> 	aşağıdaki gibi Workbook yapınca otomatikman 
	aşağıdaki prosedür oluşacaktır. Sağdaki prosedür ComboBoxına tıklandığında 
	da diğer eventleri görebilirsiniz.</p>
	<p> 	<img src="../../images/event1.jpg"></p>
		<h3> 	Kategoriler</h3>
		<p> 	Olayları 8 kategoriye ayırabiliriz. </p>
		<ul>
			<li><a href="Olaylar_WorkbookOlaylari.aspx">Workbook olayları</a></li>
			<li><a href="Olaylar_WorksheetOlaylari.aspx">Worksheet olayları</a></li>
			<li><a href="Olaylar_Grafikolaylari.aspx">Chart olayları</a></li>
			<li><a href="Formlar_Temeller.aspx">UserForm olayları</a></li>
			<li><a href="../Excel/DeveloperMenusu_Kontroller.aspx">Worksheet Kontrol olayları(Form ve ActiveX)</a></li>
			<li><a href="Olaylar_ApplicationOlaylari.aspx">Application olayları</a></li>
			<li><a href="Olaylar_OzelOlaylar.aspx">Özel olaylar</a></li>
			<li>Belli bir nesneyle ilişkili olmayan olaylar(Bunlar 
			Application.OnTime ve Application.OnKey gibi belirli bir zamanda 
			veya bir tuşa basıldığında meydana gelen olaylardır. İkisi için de
			<a href="DortTemelNesne_Application.aspx">Application nesnesiyle 
			ilgili sayfaya</a> bakabilirsiniz)</li>
		</ul>
	</div>
</asp:Content>
