<%@ Page Title='Diğer Fonksiyonlar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Fonksiyonlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='4'></asp:Label></td></tr></table></div><h1>Diğer Fonksiyonlar</h1><p>
	VBA'de kullanılan birçok fonksiyon var elbette. Bunların bir kısmını önceki 3 
	bölümde gördük. Bunun dışında farklı bölümlerde ele aldığımız(<strong>
	</strong>Interaktivite bölümüdeki <strong>InputBox</strong> gibi) ve sonraki 
	bölümlerede ele alacağımız(Dosya işlemlerindeki <strong>ChDir</strong> gibi) 
	fonksiyonlar da var.</p>
	<p>Bizim bu bölümde ele alacağımız fonksiyonlar da tıpkı bu yukarıda 
	bahsettiklerim(<strong>metinsel,numerik ve tarihsel</strong>) gibi kendi başına kullanılan önemli fonksiyonlar olacak. 
	Bunlar da öncekiler gibi çeşitli bağımsız Modüller(Information ve 
	Interactive modülleri) içinde bulunuyor, Class 
	modüllerinde değil. O yüzden bunları da tıpkı diğerleri gibi, önlerinde bir nesne olmadan kullanacağız.</p>
	<p>Aşağıdaki görselden de farkedileceği üzere Modüllerin ikonu Class 
	Modülllerinden farklıdır.(Bunlar başka dillerdeki Static/Shared sınıflara 
	benzerler)</p>
	<p><img src="/images/vbafuncidger.jpg"></p>
	
	<h3>Fonksiyonlar</h3>
	<p><span class="keywordler">IsArray, 
	IsDate,IsEmpty,IsError,IsNull,IsNumeric,IsObjcet</span>: Bunlar parametere 
	olarak aldıkları ifadelerin sırayla dizi mi, tarihsel ifade mi, boş mu, hata 
	mı,null mı, sayısal mı obje mi olduğunu döndürür. Çoğunun kullanımı farklı 
	yerlerde gösterildiği için burada ayrıca detaya girmiyorum.</p>
	<p><span class="keywordler">TypeName</span>:Değişkenin tipini verir.(String, 
	Integer, Range vs). Genelde rutin kodlarımız içinde bulunmak yerine 
	birşeyleri kontrol ederken test amaçlı kullanılır.</p>
	<p><span class="keywordler">VarType</span>: Değişkenin tip numarasını verir.<span> 
	Alacağı değerler şöyledir. 0:empty, 1:null, 2:int, 3:long,....7:Date, 
	8:string, 9:object, 11:boolean, 12:Variant (sadece variant arraylerde), 
	8192:Array(normal değer + 8192). Bu da TypeName gibi genelde test amaçlı 
	kullanılır.</span></p>
	<p>Bunların hepsini bir arada ele alındığı bir örneğe
	<a href="../Fasulye/NeNeredeNasil_NullNothingEmptyveIlkdegeratama.aspx#DebugOrnekliKod">
	buradan</a> ulaşabilirsiniz.</p>
	<p><span class="keywordler">Environ</span>:İşletim sistemiyle ilgili bilgiler verir. 
	Ya bir indeksle ya da ifade ile kullanılır. Tüm indekslerin değerlerini 
	aşağıdaki kod ile bulabilirsiniz.</p>
	<pre class="brush:vb">
Sub env()
For i = 1 To 46
    Debug.Print "i:" & i & ":" & Environ(i)
Next i
End Sub	

'İfade kullanımı da şöyledir
Debug.Print Environ("USERNAME")
</pre>
	<p>Ben şahsen bunlardan özellikle <strong>COMPUTERNAME</strong> ve <strong>USERNAME'</strong>i 
	sıklıkla kullanma ihtiyacı duyuyorum. Mesela ortak kullanılan bir dosya var 
	diyelim, bu bende açıldığında farklı bir işleve sahip olsun başklarında 
	açıldığnda farklı işleve sahip olsun istiyorsam, bunu şöyle hallederim:</p>
	<pre class="brush:vb">
If Environ("USERNAME")=12345 Then 'kullanıcı adımın 12345 olduğunu varsayın
    'diğer kodlar
Else
    'Exit Sub
End Sub</pre>
	<p>Birden fazla bilgisayarla çalıyorsam ve sadece birinde açılan dosyada işlem 
	olsun istersem de şu kod işimi görür:</p>
	<pre class="brush:vb">
If Environ("COMPUTERNAME")="A12345" Then 'kullanıcı adımın 12345 olduğunu varsayın
    'diğer kodlar
Else 'B12345 ve L12345'te birşey yapmadan çıkar
    'Exit Sub
End Sub</pre>
	<p><span class="keywordler">CreateObject ve GetObject</span>:Bunlar objeler bölümüde ele alınıyor.</p>
	<p><span class="keywordler">SendKeys</span>:Klavyeden belli tuş vey tuş 
	kombinasyonlarının basılması taklidini yapar.</p>
	<pre class="brush:vb">
SendKeys "^{F2}" 'Ctrl+F2 kombinasyonuna basılmış sayar</pre>
	<p>Tüm kullanılabilcekek parametereler
	<a href="https://msdn.microsoft.com/en-us/vba/excel-vba/articles/application-sendkeys-method-excel">
	burada</a> bulunmaktadır.</p>
	<p><span class="keywordler">Shell</span>:Belirli bir programı açar. Hesap 
	makinesi, Windows Explorer en yaygın olanlarıdır.Mesela aşağıdaki örnekte 
	diyelim ki bir dosyayı parçalara ayırdınız, dosyaların bölündüğü yer de Böl 
	klasörü olsun. Kullanıcıya en son bir mesaj verip, ilgili klasörün açılması 
	sağlanır.</p>
	<pre class="brush:vb">
Sub shellornek()

'kod bloğu
'
'
MsgBox "İşlem tamam. Dosyları görmek için tıkayınız"
Call Shell("explorer.exe" & " " & "C:\böl", vbNormalFocus)
'veya Shell "explorer.exe" & " " & "C:\böl", vbNormalFocus

End Sub
</pre>
		<p>&nbsp;</p>
</asp:Content>
