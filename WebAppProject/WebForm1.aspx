<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="WebAppProject.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Importar datos de Excel a la página web</h1>
        <asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Value="A.1">A.1 Presupuesto Global</asp:ListItem>
            <asp:ListItem>A.1.1</asp:ListItem>
            <asp:ListItem></asp:ListItem>
        </asp:DropDownList>
<asp:Button ID="importExcel" runat="server" Text="Import" />
        <asp:Button ID="generateReport" runat="server" Text="Generar Reporte" />
    </form>
</body>
</html>
