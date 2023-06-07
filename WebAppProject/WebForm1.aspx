<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="WebAppProject.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Proyecto Integrador</title>
    <link href="Content/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server" class="m-xl-4 p-xl-3 border rounded">

        <h1>Importar datos de Excel a la página web</h1>
        <p>Escoja la base de datos que desea actualizar, de click en Importar y elija el archivo que desee</p>
        <div class="form-group">
            <asp:Label Text="Base de datos" runat="server" AssociatedControlID="importTableSelector" />

            <asp:DropDownList ID="importTableSelector" runat="server" class="form-control">
                <asp:ListItem Value="A.1">A.1 Presupuesto Global</asp:ListItem>
                <asp:ListItem Value="A.1.1">A.1.1 Presupuesto de Ingresos Autorizado Global por Capítulo</asp:ListItem>
                <asp:ListItem Value="A.1.2">A.1.2 Presupuesto de Egresos Autorizado Global por Capítulo</asp:ListItem>
                <asp:ListItem Value="A.1.3">A.1.3 Presupuesto de Egresos por Secretaría Global</asp:ListItem>
                <asp:ListItem Value="A.1.4">A.1.4 Presupuesto de Egresos por Segretaría</asp:ListItem>
                <asp:ListItem Value="A.2">A.2 Reporte de Presupuesto Operativo</asp:ListItem>
                <asp:ListItem Value="A.3">A.3 Estados Financieros</asp:ListItem>
                <asp:ListItem Value="A.4">A.4 Asignación del Fondo Único de Operación</asp:ListItem>
                <asp:ListItem Value="A.4.1a">A.4.1a Arqueo del Fondo Técnico de Operación</asp:ListItem>
                <asp:ListItem Value="A.4.1b">A.4.1b Arqueo del Fondo Técnico de Operación</asp:ListItem>
                <asp:ListItem Value="A.5">A.5 Rel. de Citas. Bancarias, Inversiones, etc.</asp:ListItem>
                <asp:ListItem Value="A.5.1">A.5.1 Detalle de las Cuentas de Cheques</asp:ListItem>
                <asp:ListItem Value="A.5.1.1">A.5.1.1 Conciliación las Cuentas de Cheques</asp:ListItem>
                <asp:ListItem Value="A.5.2">A.5.2 Detalles de las Cuentas de Inversión</asp:ListItem>
                <asp:ListItem Value="A.5.2.1">A.5.2.1 Conciliación de Cuentas de Inversión</asp:ListItem>
                <asp:ListItem Value="A.6">A.6 Rel. de Citas x Cobrar</asp:ListItem>
                <asp:ListItem Value="A.6.2">A.6.2 Reporte de Saldos del Sistema de Ingresos</asp:ListItem>
                <asp:ListItem Value="A.7">A.7 Rel. de cuentas x Pagar</asp:ListItem>
                <asp:ListItem Value="A.7.1">A.7.1 Relación de Saldos con Proveedores y Contratistas</asp:ListItem>
                <asp:ListItem Value="A.7.2">A.7.2 Relación de Saldos con Acreedores Diversos</asp:ListItem>
                <asp:ListItem Value="A.7.3">A.7.3 Rel. de documentos x Pagar</asp:ListItem>
                <asp:ListItem Value="A.8">A.8 Rel. de Cheques Pendientes de Entregar</asp:ListItem>
                <asp:ListItem Value="A.9">A.9 Rel. Pol. De Fianza que Garantizan un Crédito Fiscal</asp:ListItem>
                <asp:ListItem Value="A.10">A.10 Estado que guarda la cuenta pública</asp:ListItem>
                <asp:ListItem Value="B.1">B.1 Relación de Personal de Base, Por Honorarios Asimilables</asp:ListItem>
                <asp:ListItem Value="B.2">B.2 Relación de Personal con Licencia, Permiso, Comisión, In</asp:ListItem>
                <asp:ListItem Value="B.3">B.3 Relación de Turnos y Cantidad de Personas Asignadas</asp:ListItem>
                <asp:ListItem Value="B.4">B.4 Relación de Personal Jubilado y Pensionado</asp:ListItem>
                <%--<asp:ListItem Value="C.1">C.1 Equipo, Herramienta y Accesorios y-o Listado de Patrimo</asp:ListItem>
                <asp:ListItem Value="C.1.1">C.1.1 Equipo, Herramienta y Accesorios y Condiciones de Bienes Muebles</asp:ListItem>
                <asp:ListItem Value="C.1.2">C.1.2 Equipo, Herramienta y Accesorios y-o del Estado</asp:ListItem>
                <asp:ListItem Value="C.1.3">C.1.3 Relación de Bienes Muebles Enser Menor</asp:ListItem>
                <asp:ListItem Value="C.1.4">C.1.4 RRelación de Bienes Muebles en Comodato</asp:ListItem>
                <asp:ListItem Value="C.1.5">C.1.5 Relación de Bienes Intangibles</asp:ListItem>--%>
                <asp:ListItem Value="C.2">C.2 Relación de Equipo de Transporte, Maquinaria y Combustible</asp:ListItem>
                <asp:ListItem Value="C.3">C.3 Relación de Leyes, Reglamentos, Manuales, Libros y Publicaciones</asp:ListItem>
                <asp:ListItem Value="C.3.1">C.3.1 Relación de Manuales de Organización y Proceso</asp:ListItem>
                <asp:ListItem Value="C.4">C.4 Relación de Papelería Oficial en Stock</asp:ListItem>
                <asp:ListItem Value="C.5">C.5 Inventario de Almacenes</asp:ListItem>
                <asp:ListItem Value="C.6">C.6 Relación de Bienes Muebles Propiedad de Terceros</asp:ListItem>
                <asp:ListItem Value="C.7">C.7 Relación de Armamento Municipal y del Estado</asp:ListItem>
                <asp:ListItem Value="C.8">C.8 Relación de Cd's y Cassettes de Audio y Video</asp:ListItem>
                <asp:ListItem Value="C.9">C.9 Relación de LIbros de Propiedad del Gobierno del Estado</asp:ListItem>
                <asp:ListItem Value="C.10">C.10 Relación de Equinos y Caninos</asp:ListItem>
                <asp:ListItem Value="C.11">C.11 Relación de Bienes Inmuebles Propiedad del Municipio en trámite de incorporación</asp:ListItem>
                <asp:ListItem Value="C.11.1">C.11.1 Relación de Bienes Inmuebles del Municipio Acreditados y/o Incorporados</asp:ListItem>
                <asp:ListItem Value="C.11.2">C.11.2 Relación de Bienes Inmuebles en Comodato</asp:ListItem>
                <asp:ListItem Value="D.1">D.1 Padrón de Proveedores y Contratistas</asp:ListItem>
                <asp:ListItem Value="D.2">D.2 Relación de Obras Terminadas y en Proceso</asp:ListItem>
                <asp:ListItem Value="D.3">D.3 Relación de Programas</asp:ListItem>
                <asp:ListItem Value="D.4">D.4 Relación de Contratos Financiaddos con Recursos Estatales</asp:ListItem>
                <asp:ListItem Value="D.5">D.5 Relación de Contratos Financiados con Recursos Federales</asp:ListItem>
                <asp:ListItem Value="D.6">D.6 Relación de Contratos Financiados con Recursos Propios</asp:ListItem>
                <asp:ListItem Value="D.7">D.7 Expedientes de Obras y Ubicación</asp:ListItem>
                <asp:ListItem Value="D.8">D.8 Relación de Comités de Obra Pública Formados</asp:ListItem>
                <asp:ListItem Value="E.1">E.1 Relación de Amparos, Juicios Contenciosos, Asuntos Penal</asp:ListItem>
                <asp:ListItem Value="E.2">E.2 Acuerdos, Contratos y Convenios Vigentes</asp:ListItem>
                <asp:ListItem Value="E.3">E.3 Consejos, Comités, Fideicomisos, Patronatos, Asociaciones</asp:ListItem>
                <asp:ListItem Value="E.4">E.4 Relación de Delegados Municipales</asp:ListItem>
                <asp:ListItem Value="E.5">E.5 Bienes Embargados Decomisados</asp:ListItem>
                <asp:ListItem Value="E.6">E.6 Relación de Inmuebles Desafectados</asp:ListItem>
                <asp:ListItem Value="E.7">E.7 Rel. de Regularización de Colonias</asp:ListItem>
                <asp:ListItem Value="E.8">E.8 Relación de Actas de Cabildo y Ubicación</asp:ListItem>
                <asp:ListItem Value="E.9">E.9 Informe y Documentación Relativa a los Asuntos en Trámite de las Comisiones del Ayuntamiento</asp:ListItem>
                <asp:ListItem Value="E.10">E.10 Beneficiarios de los Programas Federales y Estatales</asp:ListItem>
                <asp:ListItem Value="F.1">F.1 Relación de Papelería oficial en uso y en archivo muerto</asp:ListItem>
                <asp:ListItem Value="F.1.1">F.1.1 Relación de Expedientes y Actas en Archivo</asp:ListItem>
                <asp:ListItem Value="F.2">F.2 Archivo de Planos</asp:ListItem>
                <asp:ListItem Value="F.3">F.3 Relacioń de Asuntos en Trámite y Proyectos</asp:ListItem>
                <asp:ListItem Value="F.4">F.4 Relación de Sellos Autorizados</asp:ListItem>
                <asp:ListItem Value="I">I. Informe de Actividades</asp:ListItem>
                <asp:ListItem Value="II">II. Organigrama</asp:ListItem>
                <asp:ListItem Value="III">III. Funciones Generales</asp:ListItem>
                <asp:ListItem Value="IV">IV. Relación de Anexos Aplicables</asp:ListItem>
                <asp:ListItem Value="V">V. Plan Municipal de Desarrollo</asp:ListItem>
                <%--<asp:ListItem Value="ActaER">Acta de Entrega y Recepción</asp:ListItem>--%>
            </asp:DropDownList>
        </div>
        <br />

        <div class="flex-center-container">
            <asp:Button ID="importExcel" runat="server" Text="Importar" type="button" class="btn btn-dark text-center" />
        </div>
        <br />
        <br />

        <hr />

        <h1>Generación de reportes</h1>
        <p>Seleccione el reporte que desea generar. Deje presionado CTRL mientras selecciona para obtener más de un reporte.</p>
        <div class="form-group">
            <asp:Label Text="Reporte" runat="server" AssociatedControlID="ListBoxReport" />
            <asp:ListBox ID="ListBoxReport" runat="server" SelectionMode="Multiple" CssClass="form-control">
                <asp:ListItem Value="A.1">A.1 Presupuesto Global</asp:ListItem>
                <asp:ListItem Value="A.1.1">A.1.1 Presupuesto de Ingresos Autorizado Global por Capítulo</asp:ListItem>
                <asp:ListItem Value="A.1.2">A.1.2 Presupuesto de Egresos Autorizado Global por Capítulo</asp:ListItem>
                <%--<asp:ListItem Value="A.1.3">A.1.3 Presupuesto de Egresos por Secretaría Global</asp:ListItem>--%>
                <asp:ListItem Value="A.1.4">A.1.4 Presupuesto de Egresos por Segretaría</asp:ListItem>
                <asp:ListItem Value="A.2">A.2 Reporte de Presupuesto Operativo</asp:ListItem>
                <asp:ListItem Value="A.3">A.3 Estados Financieros</asp:ListItem>
                <asp:ListItem Value="A.4">A.4 Asignación del Fondo Único de Operación</asp:ListItem>
                <asp:ListItem Value="A.5">A.5 Rel. de Citas. Bancarias, Inversiones, etc.</asp:ListItem>
                <asp:ListItem Value="A.5.1">A.5.1 Detalle de las Cuentas de Cheques</asp:ListItem>
                <asp:ListItem Value="A.5.1.1">A.5.1.1 Conciliación las Cuentas de Cheques</asp:ListItem>
                <asp:ListItem Value="A.5.2">A.5.2 Detalles de las Cuentas de Inversión</asp:ListItem>
                <asp:ListItem Value="A.5.2.1">A.5.2.1 Conciliación de Cuentas de Inversión</asp:ListItem>
                <asp:ListItem Value="A.6">A.6 Rel. de Citas x Cobrar</asp:ListItem>
                <asp:ListItem Value="A.6.2">A.6.2 Reporte de Saldos del Sistema de Ingresos</asp:ListItem>
                <asp:ListItem Value="A.7">A.7 Rel. de cuentas x Pagar</asp:ListItem>
                <asp:ListItem Value="A.7.1">A.7.1 Relación de Saldos con Proveedores y Contratistas</asp:ListItem>
                <asp:ListItem Value="A.7.2">A.7.2 Relación de Saldos con Acreedores Diversos</asp:ListItem>
                <asp:ListItem Value="A.7.3">A.7.3 Rel. de documentos x Pagar</asp:ListItem>
                <asp:ListItem Value="A.8">A.8 Rel. de Cheques Pendientes de Entregar</asp:ListItem>
                <asp:ListItem Value="A.9">A.9 Rel. Pol. De Fianza que Garantizan un Crédito Fiscal</asp:ListItem>
                <asp:ListItem Value="A.10">A.10 Estado que guarda la cuenta pública</asp:ListItem>
                <asp:ListItem Value="B.1">B.1 Relación de Personal de Base, Por Honorarios Asimilables</asp:ListItem>
                <asp:ListItem Value="B.2">B.2 Relación de Personal con Licencia, Permiso, Comisión, In</asp:ListItem>
                <asp:ListItem Value="B.3">B.3 Relación de Turnos y Cantidad de Personas Asignadas</asp:ListItem>
                <asp:ListItem Value="B.4">B.4 Relación de Personal Jubilado y Pensionado</asp:ListItem>
                <%--<asp:ListItem Value="C.1">C.1 Equipo, Herramienta y Accesorios y-o Listado de Patrimo</asp:ListItem>
                <asp:ListItem Value="C.1.1">C.1.1 Equipo, Herramienta y Accesorios y Condiciones de Bienes Muebles</asp:ListItem>
                <asp:ListItem Value="C.1.2">C.1.2 Equipo, Herramienta y Accesorios y-o del Estado</asp:ListItem>
                <asp:ListItem Value="C.1.3">C.1.3 Relación de Bienes Muebles Enser Menor</asp:ListItem>
                <asp:ListItem Value="C.1.4">C.1.4 RRelación de Bienes Muebles en Comodato</asp:ListItem>
                <asp:ListItem Value="C.1.5">C.1.5 Relación de Bienes Intangibles</asp:ListItem>--%>
                <asp:ListItem Value="C.2">C.2 Relación de Equipo de Transporte, Maquinaria y Combustible</asp:ListItem>
                <asp:ListItem Value="C.3">C.3 Relación de Leyes, Reglamentos, Manuales, Libros y Publicaciones</asp:ListItem>
                <asp:ListItem Value="C.3.1">C.3.1 Relación de Manuales de Organización y Proceso</asp:ListItem>
                <asp:ListItem Value="C.4">C.4 Relación de Papelería Oficial en Stock</asp:ListItem>
                <asp:ListItem Value="C.5">C.5 Inventario de Almacenes</asp:ListItem>
                <asp:ListItem Value="C.6">C.6 Relación de Bienes Muebles Propiedad de Terceros</asp:ListItem>
                <asp:ListItem Value="C.7">C.7 Relación de Armamento Municipal y del Estado</asp:ListItem>
                <asp:ListItem Value="C.8">C.8 Relación de Cd's y Cassettes de Audio y Video</asp:ListItem>
                <asp:ListItem Value="C.9">C.9 Relación de LIbros de Propiedad del Gobierno del Estado</asp:ListItem>
                <asp:ListItem Value="C.10">C.10 Relación de Equinos y Caninos</asp:ListItem>
                <asp:ListItem Value="C.11">C.11 Relación de Bienes Inmuebles Propiedad del Municipio en trámite de incorporación</asp:ListItem>
                <asp:ListItem Value="C.11.1">C.11.1 Relación de Bienes Inmuebles del Municipio Acreditados y/o Incorporados</asp:ListItem>
                <asp:ListItem Value="C.11.2">C.11.2 Relación de Bienes Inmuebles en Comodato</asp:ListItem>
                <asp:ListItem Value="D.1">D.1 Padrón de Proveedores y Contratistas</asp:ListItem>
                <asp:ListItem Value="D.2">D.2 Relación de Obras Terminadas y en Proceso</asp:ListItem>
                <asp:ListItem Value="D.3">D.3 Relación de Programas</asp:ListItem>
                <asp:ListItem Value="D.4">D.4 Relación de Contratos Financiaddos con Recursos Estatales</asp:ListItem>
                <asp:ListItem Value="D.5">D.5 Relación de Contratos Financiados con Recursos Federales</asp:ListItem>
                <asp:ListItem Value="D.6">D.6 Relación de Contratos Financiados con Recursos Propios</asp:ListItem>
                <asp:ListItem Value="D.7">D.7 Expedientes de Obras y Ubicación</asp:ListItem>
                <asp:ListItem Value="D.8">D.8 Relación de Comités de Obra Pública Formados</asp:ListItem>
                <asp:ListItem Value="E.1">E.1 Relación de Amparos, Juicios Contenciosos, Asuntos Penal</asp:ListItem>
                <asp:ListItem Value="E.2">E.2 Acuerdos, Contratos y Convenios Vigentes</asp:ListItem>
                <asp:ListItem Value="E.3">E.3 Consejos, Comités, Fideicomisos, Patronatos, Asociaciones</asp:ListItem>
                <asp:ListItem Value="E.4">E.4 Relación de Delegados Municipales</asp:ListItem>
                <asp:ListItem Value="E.5">E.5 Bienes Embargados Decomisados</asp:ListItem>
                <asp:ListItem Value="E.6">E.6 Relación de Inmuebles Desafectados</asp:ListItem>
                <asp:ListItem Value="E.7">E.7 Rel. de Regularización de Colonias</asp:ListItem>
                <asp:ListItem Value="E.8">E.8 Relación de Actas de Cabildo y Ubicación</asp:ListItem>
                <asp:ListItem Value="E.9">E.9 Informe y Documentación Relativa a los Asuntos en Trámite de las Comisiones del Ayuntamiento</asp:ListItem>
                <asp:ListItem Value="E.10">E.10 Beneficiarios de los Programas Federales y Estatales</asp:ListItem>
                <asp:ListItem Value="F.1">F.1 Relación de Papelería oficial en uso y en archivo muerto</asp:ListItem>
                <asp:ListItem Value="F.1.1">F.1.1 Relación de Expedientes y Actas en Archivo</asp:ListItem>
                <asp:ListItem Value="F.2">F.2 Archivo de Planos</asp:ListItem>
                <asp:ListItem Value="F.3">F.3 Relacioń de Asuntos en Trámite y Proyectos</asp:ListItem>
                <asp:ListItem Value="F.4">F.4 Relación de Sellos Autorizados</asp:ListItem>
                <asp:ListItem Value="I">I. Informe de Actividades</asp:ListItem>
                <asp:ListItem Value="II">II. Organigrama</asp:ListItem>
                <asp:ListItem Value="III">III. Funciones Generales</asp:ListItem>
                <asp:ListItem Value="IV">IV. Relación de Anexos Aplicables</asp:ListItem>
                <asp:ListItem Value="V">V. Plan Municipal de Desarrollo</asp:ListItem>
                <%--<asp:ListItem Value="ActaER">Acta de Entrega y Recepción</asp:ListItem>--%>
            </asp:ListBox>
        </div>
        <br />


        <div class="row">

            <div class="col">
                <div class="reportInputs">
                    <asp:Label ID="Secretary" runat="server" Text="Secretaría: " AssociatedControlID="exportSecretary"></asp:Label>
                    <asp:TextBox ID="exportSecretary" runat="server" CssClass="form-control"></asp:TextBox>
                </div>
            </div>
            <div class="col">
                <div class="reportInputs">
                    <asp:Label ID="Directory" runat="server" Text="Dirección: " AssociatedControlID="exportDirectory"></asp:Label>
                    <asp:TextBox ID="exportDirectory" runat="server" CssClass="form-control"></asp:TextBox>
                </div>
            </div>
        </div>

        <br />
        <br />
        <div class="flex-center-container">
            <asp:Button ID="generateReport" runat="server" Text="Generar Reporte" type="button" class="btn btn-dark" />
        </div>

    </form>
</body>
</html>
