<%@ Page Title='DigerAraclaruzerinemakaleler MsAccess' Language='C#' MasterPageFile='~/MasterPage.master' 
AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
	<div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='Fasulye'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Diğer Araçlar üzerine makaleler'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<script src="<%= Page.ResolveClientUrl("~/syntaxhighlighter_3.0.83/scripts/shBrushSql.js" ) %>"> </script>

<h1>MS Access</h1>
<p>Tek cümle:Access'siz bir MIS ekibi düşünemiyorum. </p>
	<p>Excelin yetersiz 
kaldığı yerlerde Accessin gücü hayat kurtarıcı olmaktadır. Buna biraz da VBA 
bilgisi de eklediniz mi harikalar yaratabilirsiniz.</p>
	<p>Mesela büyük verilerle çalışıyorsanız, bunu Excelde lookupla veya Sumifs 
	tarzı bir fonksiyonla işleme tabi tutuyorsanız dakikalarca beklemeyi göze 
	almanız gerekirken Accesste saniyler içinde lookupınızı yapabilirsiniz.(Excelde 
	hızlılookup yöntemi ile tek boyutlu olması kaydıyla kısa sürede lookup 
	işlemi yapabiliyorsunuz) </p>
	<p>Lafı çok uzatmadan konulara geçelim.</p>

<h2 class="baslik">Veritabanını Otomatik Sıkıştırma</h2>
<div class="konu">
<p>Access'te dosya boyutları zaman içinde şiştikçe şişer, bunun için arada bir Compact&Repair yapmakta fayda var. 
</p>
<p>Bunu manuel yapabileceğiniz gibi, şu ayarlama ile de her Accessten çıkışta otomatik olarak bu işlemin olmasını sağlayabilirsiniz.</p>

<img src="/images/access_compact.jpg" class="zoomla" alt="Access compact" width="480" and height="360"/>

<p><strong>NOT</strong>:Dosya boyutunun şişme sebebi şudur:Dosyadan bazı verileri silip başka veriler eklersiniz, ancak arka planda verinin diskte durduğu yer tam anlamıyla silinmez ve dosya boyutuna katkıda bulunmaya devam eder.</p>
</div>



<h2 class='baslik'>Database sıkıştırmanın Excel üzerinden schedule edilmesi</h2>
<div class='konu'>
<h2>Doğrudan Excel manipülasyonu</h2>
<p>Şimdi bu örnekte ise, compact işlemini bi takvime bağlamak istiyorsunuz, yani belli saatlerde/günlerde belli 
Access dosyalarının sıkıştırılmasını istiyorsunuz diyelim.</p>

<p> Neden böyle birşey isteyebilirsiniz, çünkü bir önceki yöntemde veritabanı dosyasından çıkış yaptığınızda compact işlemini hemen o anda yapmaya çalışır ve dosya boyutu büyükse sizi birkaç dakika bekletebilir. 
Bunun yerine sizin ofiste olmadığınız gece saatleri veya haftasonunda bu işlemin yapılması çok daha 
verimli olurdu.</p>

<p>Şimdi bunun için nasıl bir Excel VBA kodu yazmak lazım, ona bakalım. Tabi bu örnekte schedule işlemini görmüyoruz, bu işlem nasıl yapılır görmek için <a href="/Konular/VBAMAkro/DortTemelNesne_Application.aspx#ontime">buraya </a>tıklayabilirsiniz.
</p>

<pre class="brush:vb">
Sub accessleri_compact()

On Error GoTo hata
Dim app As Object
Dim DBler As New Collection 'dosyalara tek tek aynı işlemi yapmamak için bir collectiona atayacağız

Set app = CreateObject("Access.Application")

DBler.Add ("C:\Paylaşım\HG Takip\Aylık Gelişmeler 2016.accdb")
DBler.Add ("C:\Paylaşım\HG Takip\Aylık Gelişmeler - Miy 2016.accdb")
'buraya istendiği kadar dosya eklenebilir

'şimdi de collection içinde geziniyor ve her eleman için aynı işlemi yapıyoruz
For Each d In DBler
    cmp = Left(d, Len(d) - 6) & "_cmp.accdb"
    okmi = app.CompactRepair(d, cmp, False) 'boolean döndürdüğü için böyle yapıyoruz

    If FileLen(d) = FileLen(cmp) Then 'eğer compact sonucunda dosya daha da küçülmediyse, boşuna işlem yapmaya gerek yok,
                                      'sadece yeni üretilen dosyayı silelim, böylece dosyamızın son erişim tarihini de değiştirmemiş oluruz
        Kill cmp
    Else
        Kill d 'orjinal dosyayı siliyoruz
        Name cmp As d 'Kompakt edilen dosyayı orjinal ismi ile rename ediyoruz
    End If

Next d

Set app = Nothing
Exit Sub
hata:
    'Hata durumunda vermek istediğinzi mesjaı veya alacğaınız aksiyonu buraya yazarsınız
End Sub</pre>

<p>Neden Excel üzerinden schedule ettik de Access üzerinden etmedik. Accessten de edilebilridi 
tabi ancak benim Excel'im her zaman açık, Access'i ise sadece ihtiyaç duydukça açıyorum. Schedule işleminin gerçekleşmesi için ise ilgili uygulamanın açık olması gerekir. O yüzden Excel'i tercih ettim. Zaten Excel üzerinden schedule ettiğim daha birçok işlem var, bu da bir yenisi olmuş oldu.</p>

<h2>Dolaylı Excel manipülasyonu</h2>
<p>Bu bölümde de dolaylı yoldan Excel ile VT sıkıştırma nasıl ona bakalım.</p>

<p>Dolaylıdan kasıt şu, schedule işlemini yine Excelde yapıyor olacağız, ancak sıkıştırma işlemini 
Excel komutu olarak değil, bunun yerine Access içinde bir makro hazırlayıp, o makroyu çalıştırıcağız.</p>

<pre class="brush:vb">
Sub gelisimdb()

On Error GoTo hata
Dim mydb As Object

Set mydb = GetObject("C:\Paylaşım\HG Takip\GelişimDB.accdb")
mydb.Application.DoCmd.runmacro "MakeTablelar" 'compact işini bu makro yapar-ys
mydb.Application.Quit
Set mydb = Nothing

Exit Sub

hata:
'hata durumunda yapılacklar-ys
End Sub
</pre>


<p>Yine bu örnekte de Doğrudan yöntemde olduğu gibi <span class=" keywordler">Options>Current Database>Compact on close</span> seçeneğine göre bir avantaj var. </p>

</div>

<h2 class='baslik'>Tablo ve Sorgu gibi nesnelerin isimlerini bulma</h2>
<div class='konu'>
	<p>Bazen Access'te tabloların, sorguların ve diğer nesnelerin adına toplu 
	şekilde ihtiyaç 
	duyarız. Bunlar için aşağıdaki sorguları çalıştırıabilirsiniz. Bunun için 
	New Query yapıp, hiçbir tablo eklemeden, sağ alt köşedeki SQL düğmesiyle SQL 
	moduna geçin ve direkt bu kodu yapıştırın.</p>

<pre class="brush:sql">
SELECT MsysObjects.Name, MsysObjects.DateCreate, MsysObjects.DateUpdate
FROM MsysObjects
WHERE (((Left$([Name],1))&lt;&gt;"~") AND ((Left$([Name],4))&lt;&gt;"Msys") AND ((MsysObjects.Type)=1))
ORDER BY MsysObjects.Name;&nbsp;</pre>
<p>Sorguları aşağıdaki elde edebilirsiniz</p>
<pre class="brush:sql">
SELECT MsysObjects.Name, MsysObjects.DateCreate, MsysObjects.DateUpdate
FROM MsysObjects
WHERE (((Left$([Name],1))<>"~") AND ((MsysObjects.Type)=5))
ORDER BY MsysObjects.Name;
</pre>
<p>Makroları da aşağıdaki elde edebilirsiniz</p>
<pre class="brush:sql">
SELECT MSysObjects.Name
FROM MsysObjects
WHERE (Left$([Name],1)<>"~") AND 
(MSysObjects.Type)= -32766
ORDER BY MSysObjects.Name;
</pre>

<p>Bazen karmaşık tablo ve sorgu yapısını sadeleştirmek için Navigation 
Pane'deki görünümü Object Type'dan özel bir Grup görünümüne dönüştürebilirsiniz. İşte burdaki grup id'lerine 
de ihtiyaç 
duyabilirsiniz. Neden ihtiyaç duyabileceğinizle ilgili örnek ise 
bir alt konuda bulunuyor.</p>
<pre class="brush:sql">
SELECT MSysNavPaneGroups.Name AS GroupName, MSysNavPaneGroupToObjects.GroupID
FROM MSysNavPaneGroups INNER JOIN (MSysNavPaneGroupToObjects INNER JOIN MSysObjects ON MSysNavPaneGroupToObjects.ObjectID = MSysObjects.Id) ON MSysNavPaneGroups.Id = MSysNavPaneGroupToObjects.GroupID
GROUP BY MSysNavPaneGroups.Name, MSysNavPaneGroupToObjects.GroupID, MSysNavPaneGroups.GroupCategoryID
HAVING (((MSysNavPaneGroups.GroupCategoryID)=3))
ORDER BY MSysNavPaneGroups.Name;
</pre>


</div>

<h2 class='baslik'>Import işlemini VBA ile otomatikleştirmek</h2>
<div class='konu'>
<p>Bazen, mesela büyük tabloların arka arkaya importlanması gereken durumlarda 
bunları manuel yapmak yerine bir butona atamak ve işten çıkarken butona basmak 
veya schedule edilmiş bir saatte kendiliğinden çalışacak şekilde ayarlamak iyi 
bir fikir olmaktadır. Bunun için aşağıdaki gibi bir kod işinizi halledecektir. </p>
	<p>Bu kodda ilaveten, eğer klasik Object Type görünümünde çalışmak yerine 
	Group görünümünde çalışıyorsanız yeni importladığınız tabloyu olması gereken 
	gruba taşıyan bir fonksiyon da bulunmaktadır, zira yeni importlanan tablo en 
	aşağıya "isimsiz gruba" düşmektedir.</p>
	<p>Önce import koduna bakalım.</p>

<pre class="brush:vb">
Sub import_gerc()

Dim blnHasFieldNames As Boolean, blnEXCEL As Boolean, blnReadOnly As Boolean
Dim lngCount As Long
Dim objExcel As Object, objWorkbook As Object
Dim colWorksheets As Collection
Dim strPathFile As String
Dim strPassword As String
Dim tbldef As TableDef, tbldefs As TableDefs
Dim grupid As Integer
grupid = 1334 'import edeceğim tabloların bulunduğu grubun id numarası

On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set objExcel = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

blnHasFieldNames = True
strPathFile = "C:\Paylaşım\HG Takip\Netice Gerçekleşen.xlsx"

strPassword = vbNullString

blnReadOnly = True 'Excel'i readonly açarız, olur da o sırada başkaları bu excelle ilgili işlem yapar diye

' excel dosyasını açar ve sheetleri bir collectiona atarız
Set colWorksheets = New Collection
Set objWorkbook = objExcel.Workbooks.Open(strPathFile, , blnReadOnly, , _
      strPassword)
For lngCount = 1 To objWorkbook.Worksheets.Count
      colWorksheets.Add objWorkbook.Worksheets(lngCount).Name
Next lngCount

' exceli save etmeden kaparız
objWorkbook.Close False
Set objWorkbook = Nothing
If blnEXCEL = True Then objExcel.Quit
Set objExcel = Nothing

'önce varolan importlu tabloları silelim
Set tbldefs = CurrentDb.TableDefs
For Each tbldef In tbldefs
    If Left(tbldef.Name, 11) = "import_gerç" Then
        DoCmd.DeleteObject acTable, tbldef.Name
    End If
Next tbldef


' Importa başlayalım
For lngCount = colWorksheets.Count To 1 Step -1
      DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, _
            "import_gerç_" & colWorksheets(lngCount), strPathFile, blnHasFieldNames, _
            colWorksheets(lngCount) & "$"
      Call grubatasi("import_gerç_" & colWorksheets(lngCount), grupid) 'eğer grup görünümünde çalışıyorsa yeni eklenen tablo uncategorized içine düşeceği
      'için onu burdan alıp olması gereken gruba atan fonksiyonu çağırırız
Next lngCount


Set colWorksheets = Nothing

End Sub
</pre>

<p>Aşağıdaki fonskiyonu uzun aramalar sonucunda 
<a href="https://social.msdn.microsoft.com/Forums/en-US/7bd1d1ca-7363-4e56-9421-c889ceb66f67/how-to-group-newly-imported-tables-using-vba-in-ms-access?forum=accessdev">burda
</a>buldum ve ihtiyacıma göre modifiye ettim. Bu arada bu kod bazı hatalar 
verdi. mesela recordsetleri deklare kısmında başlarına <span class="keywordler">
DAO </span>koymam gerekti. Bir de <span class="keywordler">myID=rs1.Fields(1)</span> olan yer
<span class="keywordler">myID=rs1![id]</span> şeklindeydi, onu da bu şekilde 
değiştirdim. Tablo adını ve grup id'sini parametre olarak verdim.</p>

<pre class="brush:vb">
Function grubatasi(tablo As String, grid As Integer)

    Dim rs1 As DAO.Recordset
    Dim rs2 As DAO.Recordset2
    
    Dim myID, myPos As Integer

    DoCmd.SetWarnings False
'Get the ObjectID for the dta_SalesAnalysis table
    Set rs1 = CurrentDb.OpenRecordset("SELECT MSysObjects.Name, MSysObjects.ID FROM MSysObjects WHERE MSysObjects.Name='" & tablo & "'")
    rs1.MoveFirst
    myID = rs1.Fields(1)
'Get the LastPosition number from the MSysPaneGroupToObjects table, and increment by 1
    Set rs2 = CurrentDb.OpenRecordset("SELECT MSysNavPaneGroupToObjects.GroupID, " & _
    "Max(MSysNavPaneGroupToObjects.Position) AS LastPos " & _
    "FROM MSysNavPaneGroupToObjects " & _
    "GROUP BY MSysNavPaneGroupToObjects.GroupID " & _
    "HAVING (((MSysNavPaneGroupToObjects.GroupID) = " & grid & ")) " & _
    "ORDER BY Max(MSysNavPaneGroupToObjects.Position);")
    rs2.MoveFirst
    myPos = rs2![LastPos] + 1

'Insert the new record into the MSysNavPaneGroupToObjects table
    DoCmd.RunSQL "INSERT INTO MSysNavPaneGroupToObjects ( Flags, GroupID, Icon, ObjectID, [Position] ) " & _
                 "SELECT 0 AS xFlag, " & grid & " AS xGroup, 0 AS xIcon, " & myID & " AS xObjectID, " & myPos & " AS xPos;"

    Set rs1 = Nothing
    Set rs2 = Nothing

    DoCmd.SetWarnings True


End Function
</pre>

</div>


<h2 class='baslik'>Zamanlanmış(Schedule edilmiş) rutinler belirlemek</h2>
<div class='konu'>
<p>Yukarda Excel'den bir görevin(sıkıştırma) nasıl schedule edildiğini görmüştük. 
Bu işi Excelle yapmıştık çünkü Excelimizin her zaman açık olduğunu, Accessin ise her zaman açık olmadığını belirtmiştik, en azından benim dünyamda böyle.</p>

<p>Peki şimdi de şöyle bir senaryomuz olsun: Diyelim ki, network üzerinde veya kendi PC'nizde shared(paylaşılmış) olarak bulunan bir Access dosyanız var, ve çeşitli kullanıcılar bu dosyayı zaman zaman açıyorlar. Ancak bu kullanıcılar malesef bazen dosyayı açık bırakıp işten öyle çıkıyorlar. Bununla beraber sizin yukarda yaptığımız gibi schedule edilmiş bir sıkıştırma işleminiz varsa bu açık dosya üzerinde işlem gerçekleşmeyecektir, çünkü MSDN'nin söylediğine göre sıkıştırma işlemi dosyanın exclusively açılmasını gerektirir, ancak bir dosya başkasında açıksa exclusively açılamaz. Bu durumda yapılması gereken iş basittir: Dosyanın 
scheduled sıkıştırma öncesinde yine scheduled bir şekilde kapanmasını sağlamak. Bunun için yapılması gereken iş de basittir:</p>

Dosyanın, açıldığında bir anaform ile açılıdığını varsayalım(eğer böyle değilse dosyayı doğrudan 
	Timer işlevini sağlayacak Form ile başlatabilirsiniz, tabi gizli olarak)
<ul>
<li>Bu anaformun Load eventine başka bir formu açmasını söylemek, adı frmTimer olsun</li>
<li>frmTimer formunu gizlemek</li>
<li>frmTimer'ın Load eventine TimeInterval girip her 1 saatte Timer eventinin çalışmasını sağlamak</li>
<li>Timer eventine de kapanmasını istediğiniz saate gelip gelmediğini kontrol edeceğiniz bir kod yazmak(aşağıdaki örnekte gece yarısından hemen önce kapatılacaktır)</li>
</ul>

<p>Kodlarımız ise aşağıdaki gibi olacak.</p>
	<h5>Ana form kodu</h5>

<pre class="brush:vb">
Private Sub Form_Load()
    DoCmd.OpenForm "Form2"
    Forms!Form2.Visible = False
End Sub
</pre>
	<h5>Timer'lı formu kodu</h5>
<pre class="brush:vb">
Private Sub Form_Load()
    Me.TimerInterval = 3600000
End Sub
Private Sub Form_Timer()
    If TimeValue(Now()) >= #10:59:00 PM# Then 'saatlik kontrol ettiğimiz için 11:59 sorun olabilir, o yüzden 10:59 yaptım. Bi düşünün bakalım neden?
        Application.Quit
    End If
End Sub
</pre>

<p><strong>NOT</strong>:TimerInterval bilgisini biz VBA kodu olarak girdik ancak design moddayken Formun Properties'ine gelip orada Timer Interval alanına da girebilirdik.</p>
<p>Evet schedule işlemi de bu kadar basit. Tabi illa kapatma işlemi yapmak zorunda değilsiniz, schedule etmek istediğiniz işlem her ne ise onu da yapabilirsiniz. Bazı işlemler dosyanın excluesively açık olmasını gerektirmez, bu nedenle uygulamayı kapatmanız gerekmez, doğrudan işleme geçebilrsiniz.</p>

	<a href="../../Ornek_dosyalar/Diğer/access_auto_kapa.accdb">Örnek dosyayı burdan indirebilirsiniz</a>
</div>

</asp:Content>
