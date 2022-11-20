<%@ Page Title='Olaylar Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Olaylar(Eventler)</h1>
<p> 	Kodların nasıl çalıştırılacağı ile ilgili bölümde görmüştük ki, çoğu zaman 
kodları F5 tuşu ile manuel çalıştırırız. Ancak kodların kendiliğinden çalışmasını sağlamanın 
da bir yolu var. İşte <strong>olaylar</strong>, bu kendinden çalıştırma 
işlemini yapmaktadır. </p>
<p>Olaylar konusuna ben çok önem veriyorum. Bu konuyu ve <strong>Application.OnTime</strong> 
konusunu iyi anladığınızda işyerinizde 2-3 ek kişi çalışıyormuş gibi bir 
verimlilik sağlarsınız.</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Olaylar') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
