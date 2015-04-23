<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>



    <!-- Allow cross domain iframing: https://msdn.microsoft.com/en-us/library/office/fp179921.aspx-->
    <WebPartPages:AllowFraming runat="server" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Yammer Embed JS Library -->
    <script type="text/javascript" src="https://c64.assets-yammer.com/assets/platform_embed.js"></script>

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>

    <div id="embedded-feed" style="height:400px;width:300px;"></div>
    
    <script type="text/javascript">
    	yam.connect.embedFeed({
            container: '#embedded-feed',
            network: gNetwork,
            feedType: gFeedType,
            feedId: gFeedId,
            config: {
                use_sso: true
            }
        });

        yam.on('/embed/feed/loadingCompleted', function (context) {
        	console.log('on load complete');
        	ResizeAppPart();
        	console.log('resize complete');
        });

        
    </script>

<!--
    yam.getLoginStatus(
            function (response) {
                if (!response.authResponse) {
                    yam.connect.embedFeed({
                        container: '#embdded-feed',
                        network: sNetwork,
                        feedType: sFeedType,
                        feedId: sFeedId
                    });
                }
            });

            network: 'pixelmill.com',
            feedType: 'group',                // can be 'group', 'topic', or 'user'          
            feedId: '4797953',                     // feed ID from the instructions above
            config: {
                defaultGroupId: 3257958      // specify default group id to post to 
            }
-->
