<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplicationDemo._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Export data to excel</h1>
    </div>

    <div class="row">
        <div class="col-md-4">
           <asp:Button ID="btnExport" Text="Export to Excel" runat="server" OnClick="btnExport_Click" />
        </div>
        
    </div>

</asp:Content>
