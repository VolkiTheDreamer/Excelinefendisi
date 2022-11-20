<%@ Page Title='DataMenusu Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
   <%--     
    <%
        Page sayfa = HttpContext.Current.CurrentHandler as Page;
        string sayfaadres=sayfa.Request.Url.AbsoluteUri.ToString();
        string ana = MyStatik.GetNameFromAdres(sayfaadres, MyStatik.NameTypeForAdres.Anakonu);
        string alt = MyStatik.GetNameFromAdres(sayfaadres, MyStatik.NameTypeForAdres.Altkonu);
        %>--%>
<h1>Data Menüsü</h1>
<p>
    
</p>
<div id="headerDatalist">Konular</div>
    <%--<p><%="selam" + sayfaadres + "___"+ana+"_____"+alt%></p>--%>
  <%--  <% string komut = @"SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
WHERE (t.AltKonu='Data Menüsü') and (a.AnaKonu='" + ana + "') order by KonuSiraNo"; %>
    <p><%=komut %></p>--%>
  
<asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1" >         
    <ItemTemplate>
        <asp:Hyperlink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
        </asp:Hyperlink>                                            
      </ItemTemplate>
    </asp:DataList>

<asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Data Menüsü') and (a.AnaKonu='Excel') order by KonuSiraNo"></asp:SqlDataSource>
  <%--  <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand=""+<%=komut %>+""></asp:SqlDataSource>--%>
</asp:Content>
  