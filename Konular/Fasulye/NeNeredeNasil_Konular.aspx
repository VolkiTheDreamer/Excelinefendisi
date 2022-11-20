<%@ Page Title='NeNeredeNasil Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Ne Nerede Nasıl</h1>
<p>Böyle bir konuya grme fikri nerden çıktı diye düşünüyor olabilirsiniz. Aslında cevap basit, bir şekilde aynı anda birok tool ile çalışıyorsanız bir süre sonra bi ifade bi toolda nasıl diğerinde nasıl diye karışmaya başlıyor, bunların derli toplu bir listesi olsa çok güzel olur diyordum, o yüzden böyle bir referans listesi hazırlamayı düşndüm, yani bu sayfa aslında bir nevi kendime hatırlatma olarak yapmaya başladım.</p>

<p>En sık karıştırılabilen konuları aşağıda bulabilirsiniz. Umarım sizlere de faydası olabilir</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Ne Nerede Nasıl') and (a.AnaKonu='Fasulye') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
