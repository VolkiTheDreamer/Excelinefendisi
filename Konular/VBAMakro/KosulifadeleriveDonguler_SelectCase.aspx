<%@ Page Title='KosulifadeleriveDonguler SelectCase' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Koşul ifadeleri ve Döngüler'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>
<h1>Select Case ve Diğer koşullu yapılar</h1>
<p>If yapısı tek koşul yapısı değildir. <span class="keywordler">Select Case</span> 
yapısı başta olmak üzere üç ayrı ifade daha vardır. Bu bölümde de bunlara bakacağız.
</p>
	<h2 class="baslik">Select Case</h2>
	<div class="konu">
	<p>Kontrol edilen şey en az 3 farklı değer için kontrol ediliyorsa <strong>If Then Else </strong>yerine
	<strong>Select Case </strong>yapısı tercih edilebilir; zira okunuruluk 
	açısından böylesi daha iyi olacaktır. Genel kullanım şekli şöyledir.</p>
		<pre class="brush:vb">Select Case SorgulananŞey
   Case şuysa
      şunu yap
   Case buysa 
      bunu yap
   Case ....
      ....
   Case Else 'Diğer durumlarda da burayı işlet
      ....
End Select</pre>
		<p>Örneğimizi kompleks durumları da içerecek şekilde yapalım.</p>
		<pre class="brush:vb">Dim sayı as Integer
sayı=Inputbox("Bir sayı girin")

Select Case sayı
   Case Is &lt; 0
      MsgBox "Negatif bir sayı girdiniz"
   Case 1 To 9
      MsgBox "Pozitif bir rakam girdiniz" '0-9 arasındaki herşey rakamdır, 9un üstündekiler rakam değil sayıdır
   Case Is &gt; 0 
      MsgBox "Pozitif bir sayı girdiniz"
   Case 0
      MsgBox "0 girdiniz" '0 da bir rakamdır ama biz onu ayrı ele allaım dedik
   Case Else 
      MsgBox "Lütfen geçerli bir sayı giriniz"
End Select</pre>
		<p>Gördüğünüz gibi bir değer aralığı için <strong>x To y</strong> 
		şeklinde, birşeyden küçük/büyük için ise <strong>Is &gt; 0</strong> 
		şeklinde yazıyoruz. Birkaç değeri yanyana yazacaksa da "," ile 
		birbirinden ayrırız. Dosya formatına göre uzantı belirlenen aşağıdaki 
		örnekte de bu durumu görüyoruz.</p>
		<pre class="brush:vb">
Select Case ActiveWorkbook.FileFormat
	Case "-4143", "-4158", 6, 56 'normal xls, txt, csv veya Excel2007deki 97-2003 xls'i mi
		FileExtStr = ".xls"
	Case 50, 51, 52 'xlsx, xlsb veya xlsm ise
		FileExtStr = ".xlsx"
	Case Else
		MsgBox "Bu dosya formatı bu makronun çalışması için uygun değil. xls, xlsx, xlsb, xlsm, txt veya csv dosyalarıyla çalışmalısınız"
		Exit Sub
End Select		</pre>
		<p><strong>NOT</strong>:En az 3 değer için sorgulama olması durumunda 
		Select Case tercih edilmeli dedik, ancak sorgulanan şey her defasında 
		değişiyorsa yine If Then Else kullanılmalıdır. Ör:</p>
		<pre class="brush:vb">Sub ifmiselectmi()
If a &gt; 10 Then
'şunları yap
ElseIf b &gt; 20 Then
'şunları yap
ElseIf c &gt; 50 Then
'şunları yap
ElseIf d &gt; 100 Then
'şunları yap
Else
'şunları yap
End If
End Sub</pre>
	</div>
	
	<h2 class="baslik">Diğer iki koşul yapısı</h2>
	<div class="konu">
	<h3>Choose</h3>
		<p>1'den başlayıp artarak giden bir sırada ilerleyen bir değer grubu 
		varsa, ve bir değişkene değer atamak istiyorsanız, bunlar için If veya 
		Select Case yapısı kullanmak yerine <span class="keywordler">Choose
		</span>ifadesinin kullanmak çok daha pratik olacaktr.&nbsp;Genel Syntax'ı 
		şöyledir.</p>
		<p><span class="keywordler">Syntax:Choose(değişken,değer1,değer2,......,değerx)</span></p>
		<p>Bu örnekte değişken=1 değerini alıyorsa dönen değer Değer1, değişken=2 
		ise dönen değer=Değer2 olur ve bu şekilde ilerler.</p>
		<p>Aylar için bu özellik rahatlıkla kullanılabilir. Örneğin, ay 
		isimlerinden oluşan bir klasör grubunuz var diyelim. Kullanıcıya ay ismini 
		girdirmek yerine ay numarasını isteyerek istediğiniz ay adını elde 
		edebilirsiniz.</p>
		<pre class="brush:vb">Sub chooseornek()
Dim ayno As Integer
Dim ayadı As String

ayno = Application.InputBox("ay no giriniz", Type:=1)
ayadı = Choose(ayno, "Ocak", "Şubat", , , , , "Aralık")
'kodların devamı

Debug.Print ayadı

End Sub</pre>
		<p>Not:Eğer eleman sayısından fazla bir index girilirse <strong>null</strong> 
		döndürür. Bu örnekte 13 girmek gibi. Tabi ayadı değişkenimiz String 
		tanımlandığı için kod hata alır. Ancak String yerine variant 
		tanımlanırsa Immediate Window'a Null yazdığını görebilirsiniz.</p>
		<h3><span style="font-size: 1em">Switch</span></h3>
		<p><strong>If Else End If</strong>'in tek satır karşılığı nasıl <strong>IIF</strong> ise, Select Case'in de 
		tek satır versiyonu <span class="keywordler">Switch</span>tir. Genel kullanımı şekli şöyledir:</p>
		<p>x=Switch(Şart 1,Sonuç 1,Şart 2,Sonuç 2,....,Şart n,Sonuç n)</p>
		<pre class="brush:vb">Sub switchornek()
kanalkodu = 1
kanaladı = Switch(kanalkodu = 1, "Şube", kanalkodu = 8, "İnternet", kanalkodu = 16, "Mobil")
MsgBox kanaladı

End Sub</pre>
<p>Okunurluğu artırmak adına şu şekilde de düzenlenebilir:</p>
<pre class="brush:vb">
Sub switchornek()
  kanalkodu = 1
  kanaladı = Switch(kanalkodu = 1, "Şube", _
  kanalkodu = 8, "İnternet", _
  kanalkodu = 16, "Mobil")
  MsgBox kanaladı
End Sub</pre>
</div>

</asp:Content>
