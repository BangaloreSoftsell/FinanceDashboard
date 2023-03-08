<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="FinanceDashboard.aspx.cs" Inherits="FinDash._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <%--</div>--%>
           

    <div class="row" >
        
        <div class="col-md-4">
            <%--<h1>Data Loader</h1>--%>
            <br />
            <asp:Button ID="DataLoader" runat="server" Text="Data Loader" OnClick="DataLoader_Click" class="btn btn-primary btn-lg" Height="43px" Width="175px" />            
        </div>
        <div class="col-md-4">
            <%--<h1>Data Transformer</h1>--%>
            <br />           
            <asp:Button ID="DataTransformer" runat="server" Text="Data Transformer" OnClick="DataTransformer_Click" class="btn btn-primary btn-lg" Height="43px" Width="189px" />            
        </div>
        <div class="col-md-4">
            <%--<h1>Data Display</h1>--%>
            <br />
            <asp:Button ID="DataDisplay" runat="server" Text="Data Display" OnClick="DataDisplay_Click" class="btn btn-primary btn-lg" Height="43px" Width="189px" />
        </div>
       
    </div>

        <%--</div>--%>

</asp:Content>
