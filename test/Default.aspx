<%@ Page Title="CONDUENT TEST - Ana Karen" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="test._Default" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <h3>EJERCICIO PRACTICO Conduent</h3>
    <h3>1.- Se debe cargar a memoria el archivo en formato Excel llamado
    Calificaciones.xlsx y se debe de mostrar todo el contenido en la vista de la
    aplicación/pagina web.</h3>

     <div>
        <table>
            <tr>
                <td>Seleccionar archivo: </td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
                <td>
                    <asp:Button ID="btnImport" runat="server" Text="Import Data" class="button" OnClick="btnImport_Click"/>
                </td>
            </tr>
        </table>
        <div>
            <br />
            <asp:Label ID="LblMessage" runat="server" Font-Bold ="true" />
            <br />
            <asp:GridView ID="GvData" runat="server" AutoGenerateColumns="false">
                <EmptyDataTemplate>
                    <div style="padding:10px">
                        Data not found!
                    </div>
                </EmptyDataTemplate>
                <Columns>
                    <asp:BoundField HeaderText="Alumno ID" DataField="AlumnoId" />
                    <asp:BoundField HeaderText="Nombres" DataField="Nombres" />
                    <asp:BoundField HeaderText="Apellido Materno" DataField="ApellidoMaterno" />
                    <asp:BoundField HeaderText="Apellido Paterno" DataField="ApellidoPaterno" />
                    <asp:BoundField HeaderText="Fecha Nacimiento" DataField="FechaNacimiento" />
                    <asp:BoundField HeaderText="Grado" DataField="Grado" />
                    <asp:BoundField HeaderText="Grupo" DataField="Grupo" />
                    <asp:BoundField HeaderText="Calificacion" DataField="Calificacion" />
                </Columns>
            </asp:GridView>
        <table width="100%">
        <tr>
            <td width ="50%" valing ="top">
                <asp:GridView ID="GvDataGrades" runat="server" AutoGenerateColumns="false" Width="95%">
                    <Columns>
                        <asp:BoundField HeaderText ="AlumnoId" DataField="Alumno ID" />
                        <asp:BoundField HeaderText ="Nombres" DataField="Nombres" />
                        <asp:BoundField HeaderText ="ApellidoMaterno" DataField="Apellido Materno" />
                        <asp:BoundField HeaderText ="ApellidoPaterno" DataField="Apellido Paterno" />
                        <asp:BoundField HeaderText ="FechaNacimiento" DataField="Fecha Nacimiento" />
                        <asp:BoundField HeaderText ="Grado" DataField="Grado" />
                        <asp:BoundField HeaderText ="Grupo" DataField="Grupo" />
                        <asp:BoundField HeaderText ="Calificacion" DataField="Calificacion" />


                    </Columns>
                </asp:GridView>
            </td>

            <h3>Grafica de barras</h3>
            <td width ="50%" valing ="top">
                <asp:Chart ID="Chart1" runat="server" BorderlineWidth="0" Width="550px">
                    <Series>
                        <asp:Series Name="Series1" XValueMember="AlumnoId" YValueMembers="Nombres"
                            LegendText="Nombres" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series2" XValueMember="AlumnoId" YValueMembers="ApellidoMaterno"
                            LegendText="ApellidoMaterno" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series3" XValueMember="AlumnoId" YValueMembers="ApellidoPaterno"
                            LegendText="ApellidoPaterno" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series4" XValueMember="AlumnoId" YValueMembers="FechaNacimiento"
                            LegendText="FechaNacimiento" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series5" XValueMember="AlumnoId" YValueMembers="Grado"
                            LegendText="Grado" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series6" XValueMember="AlumnoId" YValueMembers="Grupo"
                            LegendText="Grupo" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>

                        <asp:Series Name="Series7" XValueMember="AlumnoId" YValueMembers="Calificacion"
                            LegendText="Calificacion" IsValueShownAsLabel="false" ChartArea="ChartArea1"
                            MarkerBorderColor="#DBDBDB"></asp:Series>
                    </Series>
                    <Legends>
                        <asp:Legend Title="Alumno" />
                    </Legends>
                    <Titles>
                        <asp:Title Docking="Bottom" Text="Gráfica sobre las califiaciónes" />
                    </Titles>
                    <ChartAreas>
                        <asp:ChartArea Name="ChartArea1"></asp:ChartArea>
                    </ChartAreas>
                </asp:Chart>
            </td>
        </tr>
    </table>
   
            
            <asp:Button ID="BtnChart" runat="server" Text="Grafica" OnClick="BtnChart_Click" />
        </div>
    </div>

    <h3>3.- Seleccionar al alumno con mejor calificación: </h3>
    <div>
        <td>
            <asp:Button ID="BtnMax" runat="server" Text="Mejor califiación" OnClick="BtnMax_Click" />
        </td>
        <td>
            <asp:GridView ID="GridGrades" runat="server"></asp:GridView>
        </td>
    </div>

</asp:Content>
