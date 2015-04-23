﻿<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 

    <!-- Allow cross domain iframing: https://msdn.microsoft.com/en-us/library/office/fp179921.aspx-->
    <WebPartPages:AllowFraming runat="server" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Yammer Embed JS Library -->
    <script type="text/javascript" src="https://c64.assets-yammer.com/assets/platform_social_buttons.min.js"></script>

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>


	<!--
	NOTE: Yammer Share as an App does not work well at this time.
		the Yammer function yam.platform.yammerShare() does not allow for the input of options such as a url
		yammerShare will use the url of the app web and not the true window url.
		Provided here as an example, but in production will likely not work as intended as the "Shared" Url will be incorrect.
	-->
    <script type="text/javascript">
    	Pxlml.yammerSocial.drawShare();
    </script>
