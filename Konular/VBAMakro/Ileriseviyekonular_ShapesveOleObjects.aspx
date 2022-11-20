<%@ Page Title='Shapes ve Ole Objects' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='İleri seviye konular'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>
<h1>Shape(s)'ler ve OleObject(s)'ler</h1>
	<p>Geldik, en kafa karıştıran konulardan birine. Bu bölümde, 
	<strong>Shape</strong>(sayfa içi nesneler)&nbsp;ve onların bir alt grubu olan 
	<strong>OleObject </strong>kavramlarına değineceğiz. </p>

	<p>Özellikle, <strong>dashboard</strong> tasarlayanlar, çok fazla <strong>
	ActiveX</strong> veya form kontrolü kullananlar, Slicer/Chart gibi 
	nesnelerin görünürlük ve konum bilgilerini dinamik olarak yönetmek 
	isteyenler için önemli olmakla birlikte çalışmalarına görsellik katmak 
	isteyen herkesin bilmesinde fayda olan bir konudur.</p>
	<h2 class="baslik">Giriş</h2>
		<div class="konu">
	<p>Burada bu nesnelerin sayfa düzeni, sitil 
	özellikleri, görünürlük gibi "şekle özgü" özelliklerine bakacağız. 
	Özellikle kafanızın karışmaması gerek nokta şudur: AutoShape dışındaki şekillerin nesne 
	modelindeki özelliklerine burada girmeyeceğiz. Örneğin bir Slicer'la 
	ilgili seçim yapma, filtreleri kaldırma gibi özelliklerden ziyade Slicer'ın 
	sayfanın neresinde konumlandırılacağı, gösterilip gösterilmeyeceği gibi 
	özellikler kapsamımızda olacaktır. Bu nesnelerin, nesne modeline ait 
	konuları kendilerine ait sayafalarda bulunacaktır.
	<a href="Ileriseviyekonular_PivotTableChartveSlicernesneleri.aspx">Slicer ve 
	Chart</a>,&nbsp;
	<a href="Formlar_Kontroller.aspx">Form kontrolleri</a> gibi.</p>
	<p>Şimdi, elimizde aşağıdaki şekilleri içeren bir sayfa olduğunu düşünün. 
	İsterseniz bunları içeren dosyayı
	<a href="../../Ornek_dosyalar/Makrolar/shapeler.xlsm">buradan</a> 
	indirebilirsiniz.</p>
	<p><img src="../../images/vba_shapeler.jpg" alt="Shape on worksheet" ></p>
	<p>Burada hemen her türden nesne(Grafik, Slicer, TextBox, Konuşma balonu, 
	Resim, gömülü pdf dosyası, Form Buton, ActiveX button) var. PivotTable ve 
	Table dışında buradaki herşey bir <strong>Shape </strong>nesnesidir ve 
	bunlar doğal olarak <strong>Shapes</strong> collection'ının bir üyesidir.
	</p>
	<p>Aslında bir range'e yayılmamış olan herşey bir Shape'tir diye 
	düşünebilirsiniz. Bu bağlamda, bir range bölgesi olan Table ve 
	PivotTable'lar Shape olmamakta. Bunlar sırasıyla <strong>ListObjects</strong> ve 
	<strong>PivotTables</strong> 
	collectionlarının üyeleridir.</p>
	<p>
	<strong>OleObject</strong>'ler de Shape'lerin bir alt türüdür, <strong>yani 
	her OleObject aynı zamanda bir Shape'tir.</strong> Bu yukarıdaki örnekte; 
	ActiveX commandbutonu,&nbsp; gömülü pdf dosyası ve linkli powerpoint sunumu. 
	Zaten bunlara tıkladığımızda fonksiyon çubuğunda ya
	<strong class="keywordler">EMBED</strong> ile başlayan bir ifade görürüz: 
	=EMBED("Forms.CommandButton.1";"") ve =EMBED("Acrobat Document";"") ya da 
	Link adresi, "=PowerPoint.Slide.12|'C:\Users\Volkan\Videos\Movavi Screen 
	Capture Studio\Udemy Kurslar\2-ileri vba-makro\dosyalar\İLERİ EXCEL 
	VBA(MAKRO) EĞİTİMİ - Giriş.pptx'!'!265'"</p>
	<p>
	Şimdi bunlara yakından bakalım.</p>
	</div>
	<h2 class="baslik">Shapes Collection'ı ve Shape Nesnesi</h2>
	<div class="konu">
		<p>Yukarıda belirttiğimiz gibi, bir Range'ten ziyade sayfa üzerinde ayrı 
		bir nesne gibi duran ve mouse ile seçildiğinde köşelerinde ve kenar 
		ortalarında küçük yuvarlaklar çıkaran herşey Shape'tir.</p>
		<p>Shape'lerle ilgili önemli birkaç özellik/metod aşağıdaki gibidir.</p>
		<pre class="brush:vb">Activesheet.Shapes.SelectAll 'sayfadaki tüm shapeleri seçer
Activesheet.Shapes.Count 'sayfadaki shapelerin sayısını verir
Activesheet.Shapes.Addxxx 'xxx yerine şekil tipi gelir, Ör:AddOLEObject</pre>
		<h3>Shape'lere erişim</h3>
		<p>Tüm diğer collection(workbooks, worksheets v.s) tiplerinde olduğu 
		gibi shape'lerde de indeks(indeks no veya isim) ile tekil shape'lere 
		ulaşabiliyor ve sonra bunlara ait özellikler veya metodları 
		kullanabiliyoruz.</p>
		<pre class="brush:vb">Activesheet.Shapes(i).Delete 'i. şekili siler
ActiveSheet.Shapes("Button 1").Visible = msoFalse 'Button 1 isimli shape grünmez yapar</pre>
		<p>Tüm shapelerde dolaşmak için aşağıdaki gibi bir döngü kullanabiliriz.</p>
		<pre class="brush:vb">
Sub shapelerde_dolaş()
Dim şekil As Shape

For Each şekil In ActiveSheet.Shapes
    Debug.Print şekil.Name
Next şekil
End Sub</pre>
		<p><span class="keywordler">Type</span> property'si ile de ilgili şeklin 
		tipinin enumeration değeri döner. Bunlara ait değerleri
		<a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa432678(v=office.12)">
		buradan</a> görebilirsiniz. Aşağıda en sık kullanılanlara ait bir 
		tabloyu bulabilirsiniz.</p>
		<p>
		<table class="alterantelitable">
			
			<tr >
				<th  >Name</th>
				<th >Value</th>
				<th >Description</th>
				<th >Örnek</th>
			</tr>
			<tr >
				<td >msoShapeTypeMixed</td>
				<td >-2</td>
				<td>Mixed shape type</td>
				<td>TimeLine</td>
			</tr>
			<tr >
				<td >msoAutoShape</td>
				<td >1</td>
				<td>AutoShape</td>
				<td>Diktörtgen</td>
			</tr>
			<tr >
				<td >msoChart</td>
				<td >3</td>
				<td>Chart</td>
				<td>Grafik</td>
			</tr>
			<tr >
				<td >msoEmbeddedOLEObject</td>
				<td >7</td>
				<td>Embedded OLE object</td>
				<td>Gömü PDF</td>
			</tr>
			<tr >
				<td >msoLinkedOLEObject</td>
				<td >10</td>
				<td>Linked OLE object</td>
				<td>Linkli PowerPoint</td>
			</tr>
			<tr >
				<td >msoOLEControlObject</td>
				<td >12</td>
				<td>OLE control object</td>
				<td>ActiveX CommandButton</td>
			</tr>
			<tr >
				<td >msoPicture</td>
				<td >13</td>
				<td>Picture</td>
				<td>Resim</td>
			</tr>
			<tr >
				<td >msoTextBox</td>
				<td >17</td>
				<td>Text box</td>
				<td></td>
			</tr>
			<tr >
				<td >msoSlicer</td>
				<td >25</td>
				<td>Slicer</td>
				<td>Slicer</td>
			</tr>
		</table>
		</p>
		<p>Bu Type özelliği, özellikle If ile kontrol yapıp sadece belirli 
		şekillerle işlem yapmak istediğimizde kullanışlıdır. Mesela aşağıdaki 
		kod ile sayfadaki tüm grafikleri silebiliriz.</p>
		<pre class="brush:vb">
Sub shapleerde_dolaş()
Dim şekil As Shape

For Each şekil In ActiveSheet.Shapes
    If şekil.Type = 3 Then 'veya msoChart
        şekil.Delete
    End If
Next şekil
End Sub</pre>
		<p>
		Aşağıdaki kod ile de grafik ve comment dışındaki tüm şekilleri 
		siliyoruz.</p>
		<pre class="brush:vb">
If şekil.Type &lt;&gt; msoChart And şekil.Type &lt;&gt; msoComment Then şekil.Delete
</pre>
	</div>
	<h2 class="baslik">OleObjects Collection'ı ve OleObject Nesnesi</h2>
		<div class="konu">
	<p>Insert menüsünden "Object" olarak eklenenler ve Developer menüsünden 
	eklenen <strong>ActiveX</strong> objeleri <strong>OleObject</strong> olarak görünür. 
	ActiveX dışı OleObjeler 
	sayfaya gömülü veya linkli olurlar. <span class="keywordler">OLEType</span> özelliği ile ilgili nesnenin linkli mi yoksa gömülü mü 
	olduğu tespit edilebilir. Alacağı değerler şöyledir: </p>
			<ul>
				<li>Linkli:0(xlOLELink), 
	</li>
				<li>Gömülü:1(xlOLEEmbed) ve </li>
				<li>ActiveX Kontrol:2(xlOleControl).</li>
			</ul>
	<h3>Nesnelere Erişim</h3>
	<p>Bunlarda da tekil objelere yine indeks ile ulaşırız.</p>
	<pre class="brush:vb">Activesheet.OLEObjects("ListBox1").Delete 'veya indexno</pre>
	<p>Yine döngüyle tüm oleobjeleri dolaşalım:</p>
		<pre class="brush:vb">
Sub oleobjelerdedolaş()
Dim oleo As OLEObject

For Each oleo In ActiveSheet.OLEObjects
    Debug.Print oleo.Name
Next oleo
End Sub	</pre>
		<p>
		Collectionlarla toplu işlemler de gerçekleştirebiliriz.</p>
		<pre class="brush:vb">Activesheet.OLEObjects.Visible = False 'hepsini gizler</pre>
		<p>Sayfaya dinamik olarak OleObje eklemek de mümkündür.</p>
		<pre class="brush:vb">Worksheets(1).OLEObjects.Add FileName:="arcade.gif" 'gömülü gif dosyası
Worksheets(1).OLEObjects.Add ClassType:="Forms.ListBox.1" 'ActiveX kontrolü</pre>
		<h3>
		Nesnelerin
		Özelliklerine erişim</h3>
		<h4>
		Gömülü/Linkli nesneler</h4>
		<p>
		Gömülü/Linkli öğelerde daha çok Visible ve Top/Left gibi konum 
		özellikleriyle ilgilineceğiz. </p>
		<p>
		Gömülü bir pdf dosyasındaki özelliklere erişime bir bakalım.</p>
		<pre class="brush:vb">ActiveSheet.OLEObjects("Ole1_embed_PDF").Visible = True</pre>
		<p>
		Bu özelliklerin bir kısmına Shapes collection'ı üzerinden de 
		ulaşabiliriz. Ne de olsa tüm OleObjectler aynı zamanda bir Shape'tir. 
		Enabled gibi bazı öellikler Shape class'ında bulunmadığı için bunları 
		kullanamayız, dolayısıyla mecburen OleObject nesnesini kullanırız.</p>
		<pre class="brush:vb">ActiveSheet.Shapes("Ole1_embed_PDF").Visible = msoFalse 'False yerine msoFalse
'ama
ActiveSheet.OLEObjects("Ole1_embed_PDF").Enabled = True</pre>
		<p>
		Özellikler dışında bir de Activate ve Verb gibi metodlarla da 
		ilgilenebiliriz. Mesela bir pdf dokümanına tıklandığında onu açtıracak 
		kodu aşağıdaki gibi yazabiliriz.</p>
		<pre class="brush:vb">Sub Ole1_embed_PDF_Click()
    ActiveSheet.OLEObjects("Ole1_embed_PDF").Verb xlOpen
End Sub</pre>
		<h4>
		ActiveX kontrolleri</h4>
		<p>
		ActiveX kontrollerinde ise durum biraz karışıktır. Bu case'de
		OleObjectler, içinde bu kontrolü barındıran bir sarmalayıcı objeden 
		oluşur. Bazı durumlarda sarmalayıcının özelliklerini kullanmak bazen 
		içteki nesneyi bazen de ikisini kullanmak gerekebilir. Yine bunların bir 
		kısmına Shapes collection'ı üzerinden de ulaşabiliriz. </p>
		<p>
		Bu iç kısımdaki esas objeye ulaşmak için OLEObject nesnesinin
		<span class="keywordler">Object</span> 
		propertysi kullanılır, veya buna Shape nesnesi üzerinden ulaşıyorsak 
		önce Oleobject'yi elde etmek için OLEFormat'ı, sonra da bunun Object 
		property'si kullanılır. Özetle aşağıdaki ifadeler özdeştir: </p>
		<pre class="brush:vb">Shapes("Ole1_embed_PDF").OLEFormat.Object ile OLEObjects("Ole1_embed_PDF") 'sarmalayıcı oleobject nesnesi
Shapes("Ole1_embed_PDF").OLEFormat.Object.Object ile OLEObjects("Ole1_embed_PDF").Object 'iç kısımdaki kontrol</pre>
</div>
		<h2 class="baslik">Shape ve OleObject bir arada</h2>
	<div class="konu">	
	<p>Tüm OleObject'lerin aynı zamanda bir Shape de olduğunu söylemiştk. Peki bir 
		Shape döngüsü içindeyken veya bir şekilde elimizde bir shape nesnesi 
		varken bunların OleObect özelliklerine erişmek istersek ne 
		yaparız? Öncelikle <span class="keywordler">OLEFormat </span>özelliğini kullanırız. Bu 
	bize bir OleFormat nesnesi döndürür, bu nesnenin de 
		<span class="keywordler">Object </span>özelliğini kullanarak OleObject 
	nesnesine erişiriz. Bu nesne bir ActiveX nesnesi ise ilk Object property'siyle 
	sarmalayıcı nesneye erişmiş oluruz. OleObjectin sarmaladığı 
		içteki esas kontrole ulaşmak için ise bir <span class="keywordler">Object 
		</span>propertysi daha 
		kullanırız. Evet çok karışık oldu, farkındayım, şimdi hemen kodlara 
	bakalım, sonra açıklamayı tekrar okuyalım.</p>
		<p>Mesela yukarıdaki pdf dokümanının Click eventine aşağıdaki kodu da 
		yazabilirdik.</p>
		<pre class="brush:vb">Sub Ole1_embed_PDF_Click()
    ActiveSheet.Shapes("Ole1_embed_PDF").OLEFormat.Object.Verb xlOpen
End Sub</pre>
		<p>Veya tüm Shapelerde dolaşırken OleObject olanların isim bilgisini 
		öğrenmek için aşağıdaki kodu yazabiliriz.</p>
		<pre class="brush:vb">For Each şekil In ActiveSheet.Shapes
    If TypeName(şekil.OLEFormat.Object) = "OLEObject" Then
        Debug.Print şekil.OLEFormat.Object.Name
    End If
Next şekil</pre>
		<p>İç kısımdaki objeye ulaştıktan sonra bu sefer onun özelliklerine 
		erişebilrsiniz. Malesef adından anlaşılacğı üzere bu bir Object olduğu 
		için intellisense çıkmamaktadır. İnstellisenseten yaralanmak isterseniz 
		ilgili tipte bir nesne tanımlamanız gerekir. Aşağıdaki örnekte sayfada 
		duran iki adet CommandButtonun Caption özelliklerine ulaşıyorum, tabi 
		bunların Caption özelliğine sahip olduklarını bildiğimiz için 
		intellisense çıkmamış olsa bile ezberden yazabiliyoruz.</p>
		<pre class="brush:vb">Sub oledetay()
Dim şekil As Shape
Dim oo As OLEObject

For Each şekil In ActiveSheet.Shapes
    Set oo = şekil.OLEFormat.Object
    Debug.Print oo.Name, oo.OLEType, TypeName(oo), TypeName(oo.Object), oo.Object.Caption, oo.Left
Next şekil

End Sub</pre>
		<p>Sonuç:</p>
		<pre>CommandButton1 2 OLEObject CommandButton Düğme1 66 
CommandButton2 2 OLEObject CommandButton Düğme2 150,6 </pre>
		<p>Dikkat ettiyseniz sarmalayıcının Type'ı OLEObject iken iç nesnelerin 
		CommandButton çıkıyor.</p>
		<pre class="brush:vb">&nbsp;</pre>
		<h3>Özelliklere erişim yöntemleri</h3>
		<p>Şimdi buraya kadar öğrendiklerimizi yice pekiştirmek adına 3 tür 
		erişim şekline bakalım.</p>
		<h4>Gömülü nesne(PDF dokümanı, Word veya PowerPoint dokümanı v.s) için 
		Visible özelliği</h4>
		<h5>Shape üzerinden</h5>
		<pre class="brush:vb">ActiveSheet.Shapes("Ole1_embed_PDF").Visible = msoFalse</pre>
		<h5>OleObject üzerinden</h5>
		<p>Buna da istersek OleObject nesnesi üzerinden doğrudan veya Shape 
		üzerinden dolaylı olarak erişebiliriz.</p>
		<pre class="brush:vb">ActiveSheet.OLEObjects("Ole1_embed_PDF").Visible = True 'Doğrudan
'veya
ActiveSheet.Shapes("Ole1_embed_PDF").OLEFormat.Object.Visible = True 'Dolaylı</pre>
		<h4>ActiveX kontrolü ile Visibile özelliğine erişim</h4>
		<h5>Shape üzerinden</h5>
		<pre class="brush:vb">ActiveSheet.Shapes("Ole2_embed_ActiveXcmdbuton").Visible = msoFalse</pre>
		<h5>OleObject üzerinden</h5>
		<p>Buna da istersek OleObject nesnesi üzerinden doğrudan veya Shape 
		üzerinden dolaylı olarak erişebiliriz.</p>
		<pre class="brush:vb">ActiveSheet.OLEObjects("Ole2_embed_ActiveXcmdbuton").Visible = True 'Doğrudan
'veya
ActiveSheet.Shapes("Ole2_embed_ActiveXcmdbuton").OLEFormat.Object.Visible = True 'Dolaylı</pre>
		<p>NOT:<strong>Visible</strong> özelliği <strong>sadece sarmalayıcı obje</strong> için sözkonusudur, 
		dolayısıyla içteki Obje'ye ulaşarak bu nesnenin Visible özelliğine değer 
		atayamayız.</p>
		<h4>ActiveX kontrolü ile Enabled özelliğine erişim</h4>
		<h5>Shape üzerinden</h5>
		<p>Shape'in bu özelliği bulunmamaktadır.</p>
		<h5>OleObject üzerinden</h5>
		<p>Buna da istersek OleObject nesnesi üzerinden doğrudan veya Shape 
		üzerinden dolaylı olarak erişebiliriz.</p>
		<pre class="brush:vb">
ActiveSheet.OLEObjects("Ole2_embed_ActiveXcmdbuton").Enabled = False 'Sarmalayıcı nesneye doğrudan
'veya
ActiveSheet.Shapes("Ole2_embed_ActiveXcmdbuton").OLEFormat.Object.Enabled = False 'Sarmalayıcı nesneye dolaylı
</pre>
		<p>
		Ayrıca Enabled özelliği, içteki esas kontrol için de bulunduğu için 
		içteki nesne üzerinden de ulaşabiliriz.</p>
		<pre class="brush:vb">
ActiveSheet.OLEObjects("Ole2_embed_ActiveXcmdbuton").Object.Enabled = True 'İç nesneye doğrudan
'veya
ActiveSheet.Shapes("Ole2_embed_ActiveXcmdbuton").OLEFormat.Object.Object.Enabled = True 'İç nesneye dolaylı
</pre>
		<h4>

		ActiveX kontrolü ile FontSize özelliğine erişim</h4>
		<h5>Shape üzerinden</h5>
		<p>Shape'in bu özelliği bulunmamaktadır.</p>
		<h5>OleObject üzerinden</h5>
		<p>

		Yukarıdaki Visible özelliğinde 
		karşılaştığımız durumun tersine Font'un Size bilgisi gibi bilgilere sadece 
		<strong>iç 
		nesne </strong>aracılığı ile ulaşılmaktadır. Dış nesneden bu özelliğe erişim 
		yoktur.</p>
		<pre class="brush:vb">
ActiveSheet.OLEObjects("Ole2_embed_ActiveXcmdbuton").Object.Font.Size = 11 'Doğrudan
'veya
ActiveSheet.Shapes("Ole2_embed_ActiveXcmdbuton").OLEFormat.Object.Object.Font.Size = 14 'Dolaylı</pre>
		<h3>Özelliklere erişimle ilgili örnek</h3>
		<p>Örnek dosyamız üzerindeki tüm shape'lerin çeşitli bilgilerini 
		aşağıdaki gibi yazdırmak için bir kod yazalım. </p>
		<p><img src="../../images/vba_shapeolobje2.jpg" alt="" class="zoomla" ></p>
		<p>Başlıkları manuel yazdığımızı düşünecek olursak, L2 hücresine 
		konumlanıp aşağıdaki kou çalıştırınca bu çıktıyı elde ederiz.</p>
		<pre class="brush:vb">
Sub shapleerde_dolaş()
Dim şekil As Shape
Dim dict As New Dictionary
'OLEType,FormControlType ve AutoShapeType'lar da dictionary yapılarak isimleri yazdırılabilir
'OleType basit: (0,1,2:Linkli, Embedded, Control)
'https://docs.microsoft.com/en-us/office/vba/api/excel.xlformcontrol,
'https://docs.microsoft.com/en-us/office/vba/api/office.msoautoshapetype

dict.Add 1, "msoAutoShape"
dict.Add 2, "msoCallout"
dict.Add 20, "msoCanvas"
dict.Add 3, "msoChart"
dict.Add 4, "msoComment"
dict.Add 27, "msoContentApp"
dict.Add 21, "msoDiagram"
dict.Add 7, "msoEmbeddedOLEObject"
dict.Add 8, "msoFormControl"
dict.Add 5, "msoFreeform"
dict.Add 28, "msoGraphic"
dict.Add 6, "msoGroup"
dict.Add 24, "msoIgxGraphic"
dict.Add 22, "msoInk"
dict.Add 23, "msoInkComment"
dict.Add 9, "msoLine"
dict.Add 29, "msoLinkedGraphic"
dict.Add 10, "msoLinkedOLEObject"
dict.Add 11, "msoLinkedPicture"
dict.Add 16, "msoMedia"
dict.Add 12, "msoOLEControlObject"
dict.Add 13, "msoPicture"
dict.Add 14, "msoPlaceholder"
dict.Add 18, "msoScriptAnchor"
dict.Add -2, "msoShapeTypeMixed"
dict.Add 19, "msoTable"
dict.Add 17, "msoTextBox"
dict.Add 15, "msoTextEffect"
dict.Add 26, "msoWebVideo"
dict.Add 25, "msoSlicer"


For Each şekil In ActiveSheet.Shapes
    'ilk 3ü ole
    If şekil.Type = msoEmbeddedOLEObject Then
        Dizi = Array(şekil.Type, dict(şekil.Type), şekil.Name, şekil.OLEFormat.progID, şekil.OLEFormat.Object.Name, TypeName(şekil.OLEFormat.Object), "TypeName(şekil.OLEFormat.Object.Object)", şekil.OLEFormat.Object.OLEType)
    ElseIf şekil.Type = msoLinkedOLEObject Then
        Dizi = Array(şekil.Type, dict(şekil.Type), şekil.Name, "ole ama linkte N/A", şekil.OLEFormat.Object.Name, TypeName(şekil.OLEFormat.Object), "TypeName(şekil.OLEFormat.Object.Object)", şekil.OLEFormat.Object.OLEType)
    ElseIf şekil.Type = msoOLEControlObject Then
        Dizi = Array(şekil.Type, dict(şekil.Type), şekil.Name, şekil.OLEFormat.progID, şekil.OLEFormat.Object.Name, TypeName(şekil.OLEFormat.Object), TypeName(şekil.OLEFormat.Object.Object), şekil.OLEFormat.Object.OLEType)
    'form butonu
    ElseIf şekil.Type = msoFormControl Then
        Dizi = Array(şekil.Type, dict(şekil.Type), şekil.Name, "N/A", "N/A", "N/A", "N/A", şekil.FormControlType)
    'diğer hepsi
    Else
        Dizi = Array(şekil.Type, dict(şekil.Type), şekil.Name, "N/A", "N/A", "N/A", "N/A", şekil.AutoShapeType)
    End If
    Range(ActiveCell, ActiveCell.Offset(0, 7)).Value = Dizi
    ActiveCell.Offset(1, 0).Select
Next şekil
End Sub		
		</pre>
	</div>
	</asp:Content>
