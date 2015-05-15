Yammer and SharePoint
============

Yammer components in the form or App and WebParts for SharePoint Online and SharePoint On prem.

##Visual Studio 2013 Solutions

###SPYammerEmbedApp

SharePoint AppParts encapsulating Yammer Embed configuration into an easy to use App for SharePoint Online and SharePoint 2013 app. Current for Yammer Embed as of 4/22/2015.

Includes Yammer Embed, Yammer Embed Commenting (Open Graph Object), Yammer Like, Yammer Follow. Yammer Share does not currently work because until to get iframe's parent url.

####SPYammerEmbedApp Usage notes

When you first open the SPYammerEmbed VS Solution, you may be asked to login to Office365. Cancel this and work offline. Once you have the solution open, open the SPYammerEmbed project in the Solution Explored. You will need to update the "Site URL" property of the SPYammerEmbed project to point to your dev site url.

Thanks to Brad Morgan for the following tip:
If you continue to be prompted for login credentials, you may need to make the following update to your registry. The change helps your workstation properly resolve urls.

1. Click on Start -> Run and type regedit.
2. Locate the key
3. HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa
4. Right click on this key and choose New > DWord Value
5. Name this one "DisableLoopbackCheck"
6. Double-click then on it and type the value “1”

###SPYammerEmbedWebParts

SharePoint WebParts encapsulating Yammer Embed configuration into an easy to use Webparts for SharePoint 2013. Current for Yammer Embed as of 4/22/2015.

Includes Yammer Embed, Yammer Embed Commenting (Open Graph Object), Yammer Like, Yammer Follow and Yammer Share.

*Implemented as a full trust solution so does require central admin access / powershell access*