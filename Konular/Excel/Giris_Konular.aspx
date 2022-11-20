<%@ Page Title='Giris Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Excel Giriş</h1>
<p>Sitemin bu bölümde Excelin Tarihçesi, File Menüsü, işlerinizi kolaylaştıracak kısayol tuşları ve ribbon nesnesine(<a href="/Konular/VBAMakro/Giris_ExcelNesneModeli.aspx">nesne mi, o da ne diyorsanız böyle buyrun</a>) değinmeye çalışacağım. Anasayfada belirttiğim üzere bu sitenin amacı okuyucuya Exceli baştan itibaren anlatmak değildir, Ör:Hücre renklendirmesi, yazıcı ayarları v.s gibi temel konulara girmeyeceğim. Onun yerine ortalama 
Excel bilgisi olan kişilere işlerini daha verimli nasıl yaparlar bunları anlatmaya çalışacağım.</p>
	<p>Şunu da belirtmek isterim ki, çoğu durumda Excel'de birşeyi yapmanın 
	birden fazla yolu vardır. Bazı yöntemler uzun sürer ama bi seferlik 
	yaparsınız, sonrasında çok uğraşmazsınız, üstelik zarif de görünürler; bazı 
	yöntemler ise çok şık değildir ama sizi çok hızlı sonuca götürür. Yönteme, 
	o anki ihtiyaca göre siz karar vereceksiniz. Düzenli kullanacağınız bir 
	raporsa onu bir seferlik yapıp ama tam yapmak çoğu durum için geçerli yol 
	olacakken, sizden çok acil bir cevap isteniyorsa, gerekiyorsa birden fazla 
	yardımcı kolon açıp, fazladan formüller yazıp, gereksiz(tabi o an için gerekli) 
	özet tablolar oluşturup hızlıca istediğiniz sonuca ulaşabilirsiniz. </p>



<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
WHERE (t.AltKonu='Giriş') and (a.AnaKonu='Excel') order by KonuSiraNo"></asp:SqlDataSource>

</asp:Content>

