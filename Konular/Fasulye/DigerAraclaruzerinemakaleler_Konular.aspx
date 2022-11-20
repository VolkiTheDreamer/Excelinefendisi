<%@ Page Title='DigerAraclaruzerinemakaleler Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Diğer Araçlar</h1>
<p>Bu bölümde analitik işlerinizde kullanacağınız birkaç önemli araçla ilgili yazılarım bulunacak. Bir nevi blog tarzında bir yapılanması olacak. Burdaki diğer 3 ana konu gibi belirli bir konu bütünlüğü olmayacak. Umarım burdaki bilgiler de işinize yarayacaktor.</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Diğer Araçlar üzerine makaleler') and (a.AnaKonu='Fasulye') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
