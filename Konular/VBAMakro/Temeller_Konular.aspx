<%@ Page Title='Temeller Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Temeller</h1>
<p>Bu bölümde, bir temel oluşturması adına Visual Basic terminolojisiyle tanışacağız. Bu bölüm çok önemli, zira sonraki sayfalarda görüp okuyacağınız birçok kavramın havada kalmaması gerekiyor. Akabinde veri tiplerinin neler olduğuna ve operatörlere bakacağız. Programla kullanıcı arasındaki interaktivitenin nasıl sağlandığını göreceğiz. Son olarak da Personal.xlsb dışında alternatif kod çalıştırma alanlarına bakacağız.</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Temeller') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
