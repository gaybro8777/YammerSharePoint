<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.js"></script>
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Yammer Embed Configuration App
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

	Instructions: 
    <ul> 
        <li>Go to any web part page in the host web.</li> 
        <li>Edit the page so you can add an app part.</li> 
        <li>Add the Yammer Embed app part to your page.</li> 
        <li>Customize the app part with the custom properties.</li> 
    </ul> 

	Based On:
	<ul> 
        <li><a href="https://msdn.microsoft.com/en-us/library/office/fp179921.aspx">How to: Create app parts to install with your app for SharePoint</a></li> 
    </ul> 

    Special thanks to:
    <ul> 
        <li><a href="http://blogs.msdn.com/b/alimaz/archive/2013/02/23/building-a-sharepoint-2013-app-to-embed-yammer-feed-on-an-office-365-sharepoint-site-using-yammer-connect.aspx">Building a SharePoint 2013 App to embed Yammer feed on an Office 365 SharePoint site using Yammer Connect</a></li> 
    </ul> 

    
	References:
    <ul> 
		<li><a href="https://www.yammer.com/widget/configure">Yammer Embed Configuration Tool</a></li>
		<li><a href="https://developer.yammer.com/v1.0/docs/feed-events">Yammer Feed Events</a></li>
        <li><a href="https://code.msdn.microsoft.com/officeapps/SharePoint-2013-Resize-app-594acc88#content">SharePoint 2013: Resize app parts dynamically in apps for SharePoint</a></li> 
		<li><a href="https://code.msdn.microsoft.com/SharePoint-2013-Display-be8dac16#content">SharePoint 2013: Display app web content in the host web using an app part</a></li> 
		<li><a href="https://msdn.microsoft.com/en-us/library/office/dn481208.aspx">Client Web Part Definition Schema</a></li> 
    </ul> 
</asp:Content>

