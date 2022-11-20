<%@ Page Title='Giris Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Makrolar ve VBA'ya Giriş</h1>

<p>Bu ilk bölümde makroların ne işe yaradığından ve çalışmaya başlamadan önce nasıl bir organizasyon(dosya ayarları) yapacağımızdan bahsedeceğiz. Sonrasında VBE denen geliştirme ortamını(IDE) tanıyacağız ve Makro kaydetme aracından nasıl faydalanacağımızı göreceğiz. Son olarak da makroları daha iyi anlamak için Excel Nesne Modelini yakından tanıyacağız.</p>

<p>Konuların öyle bir özelliği var ki, birini anlatırken henüz anlatmadığımız temel bir yapıyı o konu içinde kullanmamız gerekebilir. O an için detayını çok anlamanıza gerek yoktur, yeri gelince detaylarıyla zaten öğrenmiş olacaksınız. Mesela koşullu yapılar olan <strong>IF blokları</strong> ve döngüsel yapılar olan <strong>For-Next</strong> yapıları sık sık kullanılır. Ancak örneklerimiz daha çok Excel'in temeli olan hücreler ve hücrelerin bulunduğu sayfalar üzerinden olacağı için hücreleri anlattığımız <strong>Range</strong> nesnesini ve sayfaları anlattığımız <strong>Worksheet</strong> nesnesini bu yapılardan daha önce anlattım. Eğer ki <strong>IF ve For..Next</strong> konusunu daha öne çekseydim bu sefer Range nesnesinin nasıl işlediğini merak edecektiniz. O yüzden tercihimi bu yönde kullandım. Ancak istediğiniz zaman ilgili sayfalara gidip ara bir bilgi edinme yöntemini de kullanabilirsiniz</p>

<p>
Şimdi hepsinden önce Makro nedirle başlayacağız. İkinci bölümde temellere ineceğiz. Üçüncü bölümden itibaren hız kazanacağız.</p>

<p> 
İlk 2 bölümü sağlam temel atma adına oldukça önemli buluyorum, ancak ilk başlarda anlamayacağınız yerler olabilir, o yüzden kitabın ortalarına gelindiğinde tekrar tekrar bakılması gereken bölümlerdir. 
</p>
	<p><span class="dikkat">DİKKAT:</span>Tüm konuları içeren, hatta siteye 
	dahil etmeyip sadece videosunu çektiğin örnekleri
	<a href="https://www.udemy.com/user/volkan-yurtseven/">Udemy</a> 
	eğitimlerimde bulabilirsiniz.</p>

<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Giriş') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
