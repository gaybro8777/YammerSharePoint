﻿<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 

    <!-- Allow cross domain iframing: https://msdn.microsoft.com/en-us/library/office/fp179921.aspx-->
    <WebPartPages:AllowFraming runat="server" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Yammer Embed JS Library -->
    <script type="text/javascript" src="https://c64.assets-yammer.com/assets/platform_embed.js"></script>

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>

    <script type="text/javascript">
    	Pxlml.yammerEmbed.draw();
    </script>
