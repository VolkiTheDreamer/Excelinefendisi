<%@ Page Title='Olaylar Userformolaylari' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Olaylar'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>
<h1 >Grafik Olayları</h1>

	<p> Bu bölümde her ne kadar grafiklerle ilgili VBA kod bilgisi çok gerekmese 
	de ortalama bir genel grafik bilgisinin gerekli olduğu aşikardır. O yüzden 
	eğer ihtiyaç duyuyorsanız grafiklerle ilgili genel Excel bilgisine
	<a href="../Excel/InsertMenusu_Grafikler.aspx">buradan</a>, kod bilgisine 
	ise <a href="Ileriseviyekonular_ObjelerDunyasi.aspx#chart">buradan</a> 
	alabilirsiniz. Sonrasında buradan tekrar devam de edebilirisiniz. </p>
		<p> Grafik olayları grafiğinize birşey olduğunda tetiklenirler. Bizim 
		burada ele alacağımız grafikler normal bir Worksheet içindeki gömülü 
		duran grafiklerdir. Bir Sheet türü olan Chart Sheetlerindeki 
		grafiklerin ise aşağıdaki resimdeki göreceğiniz üzere tıpkı Worksheet eventleri gibi olayları 
		vardır. O yüzden onlar bu sayfada kapsam dışılar.</p>
		<p> 
		<img src="../../images/vbachartevent1.jpg"></p>
		<p> Yanlış anlaşılma olmasın, Gömülü grafiklerin olayları da aslında 
		Chart sheetlerinki gibidir. Burada onları farklı olarak ele almak istememizin 
		sebebi, onlara ulaşım şeklimizin farklı oluşudur. Gömülü grafiklerin 
		eventlerini yakalamak için class yaratmamız gerekir.</p>
	<h2>
	Ne zaman ihtiyaç duyarız?</h2>
	<p>
	Grafik olayları oluşturduğunuz dosyalara interaktivite eklemenizi 
	sağlayarak onların daha kolay kullanımını sağlrlar. Özellikle drilldown ve 
	drillup işlemlerinde kullanışlı olabilirler. Ayrıca grafik üzerindeki bir noktaya tıkladıktan sonra o noktayla 
	ilgili detay bir bilgi baloncuğu göstermek gibi şeyler de yapabilirsiniz.</p>
	<p>
	Açıkçasını söylemek gerekirse şimdiye kadar çok kullanmadım ama kullanımının 
	faydalı olacağını düşündüğüm için konular arasına aldım. Kendim bir kullanım 
	imkanı yaratan kadar sizlere faydalı olacağını düşündüğüm birkaç link 
	vermekle yetineceğim.</p>
	<ul>
		<li>Büyük üstadlardan Jon Peltierin
		<a href="https://peltiertech.com/chart-events-microsoft-excel/">sayfası</a> 
		grafik olaylarıyla ilgili oldukça fazla miktarda bilgi içeriyor</li>
		<li>Mouse_move ile ilgili
		<a href="http://www.databison.com/interactive-chart-in-vba-using-mouse-move-event/">
		şu sitede</a> güzel bir örnek var </li>
		<li>
		<a href="https://blog.sverrirs.com/2017/02/excel-vba-chart-events.html">
		Bu sitede</a> de güzel örnekler bulabilirsiniz</li>
	</ul>
	<h2> Tanımlama şekli</h2>
	<p> Yukarıdaki linklerde detaylarını görebileceksiniz gerçi ama ben yine de 
	temel olarak nasıl bir işlem yapmanız gerektiğini anlatmak isterim. Aslında 
	<a href="Olaylar_ApplicationOlaylari.aspx">Application olayları</a> 
	tanımlamaktan bi farkı yok.</p>
	<p> Örnek dosyayı <a href="../../Ornek_dosyalar/Makrolar/chartevent.xlsm">
	buradan</a> indirebilirsiniz. Bu dosyanın ilk sayfası normal bir grafik 
	sayfası olup bunun event kodu aşağıdaki gibidir.</p>
	<pre class="brush:vb">Private Sub Chart_Activate()
    MsgBox "Bir sheet olan chart sayfası seçildi"
End Sub</pre>
	<p>Esas özel kodun yazıldığı kısım ise worksheetteki gömülü grafik içindir. 
	O da aşağıdaki gibidir. Bu kod ThisWorkbook modülüne yazılır.</p>
		<pre class="brush:vb">
Public WithEvents CHT As Chart

Private Sub Workbook_Open()
    Set CHT = Worksheets(1).ChartObjects(1).Chart
End Sub

Private Sub CHT_Activate()
    MsgBox "CHT: TypeName: " & TypeName(CHT) & vbCrLf & _
        "CHT Name: '" & CHT.Name & "'" & vbCrLf & _
        "CHT Parent TypeName: " & TypeName(CHT.Parent) & vbCrLf & _
        "CHT Parent Name: " & CHT.Parent.Name
End Sub
</pre>
		
</asp:Content>
