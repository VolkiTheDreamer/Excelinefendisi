<%@ Page Title='Değer Tipi Sorgulama' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Fasulye'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Ne Nerede Nasıl'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='8'></asp:Label></td></tr></table></div>

<h1>Değer Tipi Sorgulama</h1>
<p><!--Bir hücrenin, değişkenin, alanın veritipinin ne olduğunu sorgulama konusu da fakrlı platformlarda benzer şekillerde ifade edilebildiği için karışıklığa nedne olabilmektedir. Bu bölümde bunları ele aldım.</p>

<h2>Sayısal mı?</h2>
<p>
<strong>Excel</strong>:IsNumber()<br>
<strong>VBA</strong>:IsNumeric()<br>
<strong>Vb.Net</strong>:IsNumeric()<br>
<strong>C#</strong>:C#'ta malsef Vb.Netteki gibi Isnmeric yok, bunun  yerine bir extension metod yazmak gerekir

<pre class="brush: csharp">public static class Extension
{
    public static bool IsNumeric(this string s)
    {
        float output;
        return float.TryParse(s, out output);
    }
}</pre>
<strong>SQL(SQL-Server)</strong>:IsNumeric()<br>
<strong>SQL(Oracle)</strong>:LENGTH(TRIM(TRANSLATE(string1, ' +-.0123456789', ' ')))
</p>
-->
</asp:Content>
