<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="WebAppProject.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Importar datos de Excel a la página web</h1>
    <input type="file" id="inputFile" />
    <button onclick="importExcelData()">Importar archivo de Excel</button>
        <asp:Button ID="importExcel" runat="server" Text="Import" />
    </form>
</body>
</html>
