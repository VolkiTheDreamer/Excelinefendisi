<%@ Page Title='DebuggingveHataYonetimi Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

	<h1>Debugging ve Hata Yönetimi</h1>
<p>Debuggingi bilmediğim dönemlerde kodun bi yere kadar çalışmasını istediğim 
zamanlarda oraya bir MsgBox koyardım, MsgBox çıkınca da Ctrl+Break ile kodu 
durdururdum. Debugging araçları işte bu tür uzun ve kafa yaran süreçlerden 
kurtulmanızı, kodunuzu daha sağlıklı test etmenizi sağlar.</p>
	<p>Hata Yakalama ise, özellikle başkalarının kullanımına açık olacak şekilde 
	kodlama yaptığınızda(Ör:Bir add-in) beklenmeyen durumlarda 
	kodunuzun uygun reaksiyonu vermesini sağlar. Doğru bir hata yönetimi 
	uygulanmazsa kodunuz beklenmedik sonuçlara neden olabilir, yapılan işlem 
	yarıda kalabilir, kullanıcının güveni zedelenebilir.</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Debugging ve Hata Yönetimi') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
