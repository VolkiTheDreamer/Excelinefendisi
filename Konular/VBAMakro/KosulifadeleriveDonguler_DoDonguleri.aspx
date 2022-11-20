<%@ Page Title='KosulifadeleriveDonguler LoopDonguleri' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Koşul ifadeleri ve Döngüler'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div>

<h1>Do ve While Döngüleri</h1>
<p> For döngülerini incelerken gördük ki, bu döngüler genellikle çevrelenen kod parçasının 
kaç kez çalıştırılacağının bilindiği durumlarda kullanılmaktaydı. </p>
	<p> Üst limitin bilinmediği diğer durumlarda ise çoğunlukla <strong>Do</strong> ve 
	<strong>While</strong> 
	döngüleri kullanıır. While ifadesi ile genelde "<strong>Do While</strong>" 
	kalıbı içinde geçer ve bu 
	aslında While döngüsü değil bir Do döngüsüdür. <strong>While Wend</strong> döngüsü artık çok 
	kullanılan bir döngü olmayıp bu bölümdeki&nbsp; ve site genelindeki 
	kodlarımızı Do döngüleriyle halletmeye çalışacağız. En son kısa bir While 
	Wend döngü örneği de yapıp konuyu bitireceğiz.</p>
	<p> Do döngüleri de For döngüleri gibi dizi ve dizimsilerde bol miktarda 
	kullanılır, ayrınılı bilgi için<a href="KosulifadeleriveDonguler_ForDonguleri.aspx"> 
	For Döngüleri</a> ve <a href="DizilerveDizimsiYapilar_Konular.aspx">Diziler</a> 
	bölümlerine bakabilirsiniz.</p>
	<h2> Do Döngüleri</h2>
	<p> Do döngülerinin 2 ana, 4 alt tipi vardır.</p>
	<ul>
		<li><strong>Do While</strong><ul>
			<li><span class="keywordler">Do While Şart şuysa.......Loop</span>:Şart gerçekleşmezse çevrelenmiş 
			kod hiç çalışmayabilir</li>
			<li><span class="keywordler">Do ....... Loop While Şart şuysa</span>:Şart gerçekleşse de 
			gerçekleşmese de <strong>en az 1 kez çalışır</strong></li>
		</ul>
		</li>
		<li><strong>Do Until</strong><ul>
			<li><span class="keywordler">Do Until Olay......Loop</span>:Olay olana(şart 
			gerçekleşene) kadar çalışır, döngüye 
			girildiğinde olay zaten olmuşsa(şart gerçekleşmişse) çevrelenmiş kod hiç çalışmaz</li>
			<li><span class="keywordler">Do ....... Loop Until Olay</span>:Olay olana(şart 
			gerçekleşene) kadar çalışır; Çevrelenmiş 
			kod <strong>en az 1 kez çalışır</strong> </li>
		</ul>
		</li>
	</ul>
	<p>Gördüğünüz üzere, <strong>şart eğer Do satırındaysa, kod hiç çalışmayabiliyor, ama 
	Loop satırındaysa en az 1 kez çalışır.</strong></p>
	<p>Hemen bir örnek yapalım. Kullanıcıdan karesi alınacak bir sayı girmesini 
	isteyelim ve eğer kullanıcı geçerli bir sayı yerine başka birşey mesela bir 
	harf girerse başa dönmesini sağlayalım. Bunu bir If ve GoTo ile de 
	yapabilirdik ama bu sefer Do ile yapacağız.</p>
<pre class="brush:vb">
Sub dowhile1()
'bu en az 1 kere çalışır, ve sayı girene kadar bize aynı soruyu sorar 
Do
  sayı = InputBox("Karesi alınacak bir sayı girin")
Loop While Not IsNumeric(sayı)
 
MsgBox sayı & " sayısının karesi şudur:" & sayı * sayı
 
End Sub

Sub dowhile2()
'Bu ise soruyu hiç sormaz. Çünkü önce şart kontrolünü yapar, 
'şart sağlanmadığı için döngüye girmeden çıkar 
'Şart sağlanmaz çünkü say değişkeni varianttır ve 
'Variantlara henüz değer atanmadıysa 0 değerini alır, yani IsNumeric sorugus True'dur
'başında da bir Not operatörü olduğu için şart sağlanmaz.

Do While Not IsNumeric(sayı)
  sayı = InputBox("Karesi alınacak bir sayı girin")
Loop
 
MsgBox sayı & " sayısının karesi şudur:" & sayı * sayı
 
End Sub</pre>
	<p>
	Aşağıda ise Do Until yapısına ait 2 örnek bulunuyor.</p>
	<pre class="brush:vb">
Sub dountil1()
'sayfa sayısı 5 olana kadar işlem yapar. 5ten çoksa siler azsa ekler.
Application.DisplayAlerts = False
If Sheets.Count < 5 Then
    Do Until Sheets.Count = 5
        Sheets.Add After:=Sheets(Sheets.Count)
    Loop
ElseIf Sheets.Count > 5 Then
    Do Until Sheets.Count = 5
        Sheets(Sheets.Count).Delete
    Loop
End If
End Sub

Sub dountil2()
'en az bir kere çalışır, taki sayfa sayısı 1 olana kadar
Do
  Sheets(2).Delete 'Her defasında hep 2.sayfa silinir, ta ki tek sayafana kalana kadar
Loop Until Sheets.Count = 1
End Sub
</pre>
	<h3>While mı Until mi?</h3>
	<p>Hangi durumda hangisini kullanmalıyız?</p>
	<p>İkisini de her durumda kullanabilirsiniz. Tek farkı, birini diğerinin 
	tersi mantıkla yazmak. Geliştiriciler neden böyle bir ayrıma girmişler emin 
	değilim ama sanırım konuşma dilindeki kullanım terchilerimizi gözönüde 
	bulundurmuş olabilirler. Mesela çocuğumuza "yemeğin bitene kadar masadan 
	kalkmak yok" da diyebilriz, "yemeğin bitmediği sürece masadan ayrılmak yok" 
	da, ikisi de aynı kapıya çıkar. VBA'de de durum pek farklı değil: En alt satıra 
	inene kadar çalış da 
	diyebiliriz, "aktif satır no&lt;son satır no" olduğu sürece çalış da.</p>
<h3>Döngüden Çıkış</h3>
<p>For döngülerinde olduğu gibi Do Looplarından çıkış için de Exit ifadesini 
	kullanırız. Mesela aşağıdaki döngüde, boş bir hücreye rastlanıldığında 
	döngüden çıkılır.</p>
<pre class="brush:vb">
i=1
Do While i &lt; 1000
  If IsEmpty(Cells(i,1)) Then 
     Exit Do
  End If
  i = i + 1
Loop </pre>



	<h3>İçiçe Loop</h3>
	<p>Diyelim ki bir bölgenin şubelerini hacimlerine göre boy sırasına 
	dizdiniz, buna göre her bölgenin en yüksek hacimli rakamını yazdırmak 
	istiyorsunuz, yani bir nevi Excelde varolmayan(2016da MAXIFS geldi) bir 
	MAXIF fonksiyonunu Sub prosedür olarak ele alacağız ve bunun tersi olan 
	MINIF'i.</p>
	<p>Tabiki bunları bir UDF ile yapmak daha şık olackatır ancak içiçe Loop 
	örneği görmek adına bu örnek faydalı olacaktır.</p>
	<pre class="brush:vb">
Sub maxif()
'maxı yazacağın yere gel, ordayken çalıştır ve liste sıralı olsun
 
Set kriter = Application.InputBox("ana değişken kriterinin olduğu sütundan bir hücre seç", Type:=8)
Set rakam = Application.InputBox("maksimumun arandığı sütunu seç", Type:=8)
 
ks = ActiveCell.Column - kriter.Column
rs = ActiveCell.Column - rakam.Column
 
 
Do
  
    Maks = ActiveCell.Offset(0, -rs).Value
    Set ilkyer = ActiveCell
   
    Do While ActiveCell.Offset(0, -ks).Value = ActiveCell.Offset(1, -ks).Value
       
        If ActiveCell.Offset(1, -rs).Value > Maks Then Maks = ActiveCell.Offset(1, -rs)
        ActiveCell.Offset(1, 0).Select
               
    Loop
   
    Set sonyer = ActiveCell
    Range(ilkyer, sonyer).Value = Maks
    ActiveCell.Offset(1, 0).Select
       
Loop Until ActiveCell.Offset(0, -1).Value = ""
 
End Sub
 
Sub minif()
'mini yazacağın yere gel, ordayken çalışıtır ve liste sıralı olsun
 
Set kriter = Application.InputBox("ana değişken kriterinin olduğu sütundan bir hücre seç", Type:=8)
Set rakam = Application.InputBox("minimumun arandığı sütunu seç", Type:=8)
 
ks = ActiveCell.Column - kriter.Column
rs = ActiveCell.Column - rakam.Column
 
 
Do
  
    Mini = ActiveCell.Offset(0, -rs).Value
    Set ilkyer = ActiveCell
   
    Do While ActiveCell.Offset(0, -ks).Value = ActiveCell.Offset(1, -ks).Value
       
        If ActiveCell.Offset(1, -rs).Value <= Mini Then Mini = ActiveCell.Offset(1, -rs)
        ActiveCell.Offset(1, 0).Select
               
    Loop
   
    Set sonyer = ActiveCell
    Range(ilkyer, sonyer).Value = Mini
    ActiveCell.Offset(1, 0).Select
       
Loop Until ActiveCell.Offset(0, -1).Value = ""
 
End Sub	</pre>
	<h2>While Wend döngüsü</h2>
	<p>Başta da söylediğimiz gibi bu döngü tipi artık pek kullanılmıyor, sadece 
	eski yazılmış bir kod karşımıza geldiğinde bilelim diye hala var. Ama 
	MSDN'nin bize söylediği gibi, Do Döngüsü daha fonksiyonel ve esnektir ve 
	yeni kodlarda bu kullanılmaldır. </p>
	<p>While döngülerinde sadece başta şart kontrolü yapılır, yani en az 1 kere 
	çalışma olayı bunda seçime bağlı değildir. Bu da demektir ki döngü içindeki kodumuzun hiç çalışmaması 
	mümküdür. Ayrıca Do ve For dönülerindeki gibi döngüden çıkış 
	ifadesi(Exit) de yoktur.</p>
	<p>Aşağıdaki örnekte 1'den 10 kadar olan sayıların toplamını alıyoruz.</p>
	<pre class="brush:vb">Sub topla()
While sayı &lt;= 10
   toplam = toplam + sayı
   sayı = sayı + 1
Wend

MsgBox toplam
End Sub
&nbsp;</pre>

<h2 class="baslik">Çeşitli Örnekler</h2>
<div class="konu">
<h4 class="baslik">Bir hücre grubuna harfleri yazdırma</h4>
<div>
<pre class="brush:vb">Sub harfyaz()
Dim i As Integer
i = 1
Do While i &lt; 27
  Cells(i, 1) = Chr(i + 64) 'A harfinin Ascii kodu 65tir
  i = i + 1
Loop
End Sub</pre>
</div>
<h4 class="baslik">İlk sayfa hariç tüm sayfaları silme</h4>
<div>
<pre class="brush:vb">Do
  Sheets(2).Delete 'Her defasında hep 2.sayfa silinir, ta ki tek sayafana kalana kadar
Loop Until Sheets.Count = 1</pre>
</div>
	<h4 class="baslik">Database'den tek tek tarihler için data çekip başka sayfada altalta 
	getirme</h4>
<div>
	<p>Diyelim ki, bir nedenle tarihsel datasını tek tek çekmeniz gereken bir sorgu var. 
	Ör:1 Ocak için ayrı, 2 Ocak için ayrı,.... v.s Bunu 
	bir SQL ortamında yapmanız ağır bir iş gelebilir. Neyse ki VBA ve döngüler 
	sayesinde bu ağır iş bile kolaylaşabilir.</p>
	<pre class="brush:vb">Sub dowhileornek()

Dim t As Date
t = "01.01.2016"

Do
'SQL stringini oluşturuyorum
s1 = "Select * from PARG.ARG_TNETICE_MUSTERI_GER" &amp; Chr(13)
s2 = "where BAGLI_URUN_ID in (165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184)"
s3 = "and TARIH=" &amp; t &amp; " and KPI_ID=1001 and RAPOR_TURU=2"
s4 = "and musteri_id in (Select ESLESEN_TARAF_ID from PDWH.DWH_TSAKLAMA where YONLENDIREN_SUBE &lt;&gt; 1234 and TARIH=t)" &amp; Chr(13)


With ActiveWorkbook.Connections("DWH").ODBCConnection
.BackgroundQuery = True
.CommandText = Join$(Array(s1 &amp; s2 &amp; s3 &amp; s4))
'Diğer kodlar
End With

ActiveWorkbook.Connections("Query from PDWH_USR").Refresh
t = t + 1 'Tarihi 1 artırıyorum

ActiveCell.CurrentRegion.Select
Selection.Copy
Sheets(2).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste 'ilk çalışan sorgu sonucunu 2.sayfaya ekleyip bir alt satıra iniyorum

Loop Until t = "15.02.2016" '15 şubata kadar yani 45 kez sorgu çalışacak
'SQL ortamınıda bunu yapmak çok yorucu olabilirdi

End Sub</pre>
</div>
<h4 class="baslik">Belirli adette sayfası olan dosya yaratma</h4>
<div>
	<p>Diyelim ki, kodumuzun bir yerinde 5 sayfalı bir workbook yaratmamız 
	gerekiyor. Normalde Excelin default ayarı yeni workbook açılıdığında 3 adet 
	safyası olması yönündedir ama biz bu default ayarı değiştirmiş olabiliriz, 
	mesela 1 veya 7 yapmış olabiliriz. O&nbsp; yüzden ilk olarak bir If bloğu 
	ile kontrol ederiz ve eğer 5 ten büyükse fazlalıkları sildirip, küçükse 5e 
	tamamlayan bir kod yazarız. Sayfa sayımız zaten 5se kod hiçbirşey yapmaz.</p>
	
	<pre class="brush:vb">
Sub sayfaekleveyasil()
Application.DisplayAlerts = False 'sayfa silerken uyarmasın
If Sheets.Count &lt; 5 Then
	Do Until Sheets.Count = 5
		Sheets.Add After:=Sheets(Sheets.Count)
	Loop
ElseIf Sheets.Count &gt; 5 Then
	Do Until Sheets.Count = 5
		Sheets(Sheets.Count).Delete
	Loop
End If
End Sub</pre>
    </div>
</div>
</asp:Content>
