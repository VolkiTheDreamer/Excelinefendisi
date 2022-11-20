<%@ Page Title='FormulasMenusu2 Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Formulas Menüsü-Diğer Bilgiler</h1>
<p></p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Formulas Menüsü(Diğer)') and (a.AnaKonu='Excel') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
