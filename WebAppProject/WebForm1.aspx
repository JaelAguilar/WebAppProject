<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="WebAppProject.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Importar datos de Excel a la página web</h1>
        <asp:DropDownList ID="importTableSelector" runat="server">
            <asp:ListItem Value="A.1">A.1 Presupuesto Global</asp:ListItem>
            <asp:ListItem>A.1.1</asp:ListItem>
            <asp:ListItem></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="importExcel" runat="server" Text="Import" />
        <h1>Generación de reportes</h1>
        <asp:ListBox ID="ListBox1" runat="server" SelectionMode="Multiple">
            <asp:ListItem Value="A.1">A.1 Presupuesto Global</asp:ListItem>
            <asp:ListItem Value="A.2">A.2 Presupuesto Global</asp:ListItem>
            <asp:ListItem Value="A.3">A.3 Presupuesto Global</asp:ListItem>
        </asp:ListBox><br />
        <asp:Label ID="Secretary" runat="server" Text="Secretaría: ">
            <asp:TextBox ID="exportSecretary" runat="server"></asp:TextBox>
        </asp:Label><br />

        <asp:Label ID="Directory" runat="server" Text="Dirección: ">
            <asp:TextBox ID="exportDirectory" runat="server"></asp:TextBox>
        </asp:Label>

        <asp:Button ID="generateReport" runat="server" Text="Generar Reporte" />
    </form>
</body>
</html>
