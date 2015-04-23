using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

namespace YammerEmbedWebpart.Webparts.YammerCommenting
{
	/* Yammer commenting documentation: https://developer.yammer.com/v1.0/docs/commenting */
	/*
	 * Yammer commenting is based on open-graph Yammer Embed.
	 * Since the Yammer Embed Webpart could be configured to allow for commenting, the value for this webpart is to pre-configure specific options for the end user
	 * making it easier to add comments to a give page.
	 * 
	 * Privacy not included in this webpart. Commenting allows for a setting, "private" to be set to true. This will use a  "users" object to determine who has the 
	 * authorization to review the comments. A method to provide a list of users who has access to the OG object would be required.
	 */
	[ToolboxItemAttribute(false)]
	public class YammerCommenting : WebPart
	{
		protected List<string> errorArray = new List<string>();

		#region Custom Webpart properties
		//types from https://developer.yammer.com/v1.0/docs/schema
		public enum yammerOpenGraphTypes { Page, Place, Person, Department, Team, Project, Folder, File, Document, Image, Audio, Video, Company }

		#region container settings
		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Unique wrapper container selector"),
		WebDescription("An unique selector for this instance, i.e. #embedded-feed. Leave blank for default.")]
		[Category("Yammer Embed Wrapper")]
		public String WrapperId { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Container Width"),
		WebDescription("Width in pixels for container. \"0\" for default.")]
		[Category("Yammer Embed Wrapper")]
		public int WrapperWidth { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Container Height"),
		WebDescription("Height in pixels for container. \"0\" for default.")]
		[Category("Yammer Embed Wrapper")]
		public int WrapperHeight { get; set; }
		#endregion

		#region Embed JavaScript Settings
		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Network Permalink (network)"),
		WebDescription("Your Network, i.e. microsoft.com")]
		[Category("Yammer Embed Settings")]
		public String NetworkPermalink { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Show Header? (config.header)"),
		WebDescription("Display Header in Feed?")]
		[Category("Yammer Embed Settings")]
		public bool ShowHeader { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Hide Network Name Header? (config.hideNetworkName)"),
		WebDescription("Hide Network Name?")]
		[Category("Yammer Embed Settings")]
		public bool HideNetworkName { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Prompt Text (config.promptText)"),
		WebDescription("Comment prompt text, leave blank for default")]
		[Category("Yammer Embed Settings")]
		public String PromptText { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Show Footer? (config.footer)"),
		WebDescription("Display Footer in Feed?")]
		[Category("Yammer Embed Settings")]
		public bool ShowFooter { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Use Single Sign On (SSO)? (config.use_sso)"),
		WebDescription("Use Single Sign On?")]
		[Category("Yammer Embed Settings")]
		public bool UseSSO { get; set; }
		#endregion

		#region Open Graph Object Properties
		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Open Graph Type (objectProperties.type)"),
		WebDescription("Type of Open Graph Object to retrieve")]
		[Category("Commenting Settings")]
		public yammerOpenGraphTypes OpenGraphType { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Open Graph Url (objectProperties.url)"),
		WebDescription("Url of Open Graph Object to retrieve")]
		[Category("Commenting Settings")]
		public String OpenGraphUrl { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Show Open Graph Preview? (config.showOpenGraphPreview)"),
		WebDescription("Show the preview for an Open Graph Object if available?")]
		[Category("Commenting Settings")]
		public bool ShowOpenGraphPreview { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Open Graph Title (objectProperties.title)"),
		WebDescription("Title of Open Graph Object")]
		[Category("Commenting Settings")]
		public String OpenGraphTitle { get; set; }

		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Open Graph Image Url (objectProperties.image)"),
		WebDescription("Image Url of Open Graph Object preview")]
		[Category("Commenting Settings")]
		public String OpenGraphImageUrl { get; set; }
		#endregion

		#endregion

		protected override void CreateChildControls()
		{
			if (this.WebPartManager.DisplayMode.Name == "Design")
			{
				//this.ChromeType = PartChromeType.TitleOnly;
				this.AllowEdit = true;
			}
			else
			{
				//this.ChromeType = PartChromeType.None;
				this.AllowEdit = false;
			}
			this.AllowHide = false;

			String sID = this.ClientID;

			PlaceHolder displayPh = new PlaceHolder();
			PlaceHolder editPh = new PlaceHolder();

			displayPh.ID = "displayPanel";
			displayPh.Visible = false;
			displayPh.EnableViewState = false;

			editPh.ID = "editPanel";
			editPh.Visible = false;
			editPh.EnableViewState = false;

			this.Controls.Add(new LiteralControl("<div class=\"yammerEmbedWpWrapper\">"));

			//if in edit mode then show edit panel
			if (this.WebPartManager.DisplayMode == System.Web.UI.WebControls.WebParts.WebPartManager.EditDisplayMode
					|| this.WebPartManager.DisplayMode == System.Web.UI.WebControls.WebParts.WebPartManager.DesignDisplayMode)
			{
				/*edit panel*/
				editPh.Controls.Add(new LiteralControl(DrawEditPanel()));
				editPh.Visible = true;

				this.Controls.Add(editPh);
			} /*end edit panel*/
			else //display panel
			{
				displayPh.Controls.Add(new LiteralControl(DrawDisplayPanel()));
				displayPh.Visible = true;
				this.Controls.Add(displayPh);
			} //end display panel


			//messages
			if (errorArray.Count > 0)
			{
				this.Controls.Add(new LiteralControl("<div class=\"yammerEmbedWpMessages\">"));
				BulletedList errors = new BulletedList();
				errors.ID = "ErrorList";
				errors.DataSource = errorArray;
				errors.DataBind();
				this.Controls.Add(errors);
				this.Controls.Add(new LiteralControl("</div>"));
			}
			//end messages

			this.Controls.Add(new LiteralControl("</div>"));
		}

		#region helper functions
		private String DrawEditPanel()
		{
			StringBuilder sbEdit = new StringBuilder();
			sbEdit.Append("<div class=\"yammerEmbedWpEditPanel\">\r\n");
			sbEdit.Append("<span class=\"yammerEmbedWpWarning\">Edit this web part's properties to modify settings.</span>\r\n");
			sbEdit.Append("<hr />\r\n");
			//NetworkPermalink
			if (!String.IsNullOrEmpty(NetworkPermalink))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Network Permalink:</span> <span class\"yammerEmbedWpValue\">" + NetworkPermalink + "</span></div>\r\n");
			else
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Network Permalink:</span> <span class\"yammerEmbedWpValue yammerEmbedWpWarning\">Client default network</span></div>\r\n");

			sbEdit.Append("<hr />\r\n");

			//WrapperId
			if (!String.IsNullOrEmpty(WrapperId))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Custom wrapper element selector:</span> <span class\"yammerEmbedWpValue\">" + WrapperId + "</span></div>\r\n");

			//WrapperWidth
			if (WrapperWidth > 0)
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Custom width:</span> <span class\"yammerEmbedWpValue\">" + WrapperWidth + "</span></div>\r\n");

			//WrapperHeight
			if (WrapperHeight > 0)
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Custom height:</span> <span class\"yammerEmbedWpValue\">" + WrapperHeight + "</span></div>\r\n");

			sbEdit.Append("<hr />\r\n");
			//ShowHeader
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Show Header?</span> <span class\"yammerEmbedWpValue\">" + ShowHeader.ToString() + "</span></div>\r\n");

			//HideNetworkName
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Hide Network Name?</span> <span class\"yammerEmbedWpValue\">" + HideNetworkName.ToString() + "</span></div>\r\n");

			//PromptText
			if (!String.IsNullOrEmpty(PromptText))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Custom add comment prompt text:</span> <span class\"yammerEmbedWpValue\">" + PromptText + "</span></div>\r\n");

			//ShowFooter
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Show Footer?</span> <span class\"yammerEmbedWpValue\">" + ShowFooter.ToString() + "</span></div>\r\n");

			//UseSSO
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Use Sign Sign On (SSO)?</span> <span class\"yammerEmbedWpValue\">" + UseSSO.ToString() + "</span></div>\r\n");

			sbEdit.Append("<hr />\r\n");

			//if FeedType == OpenGraph
			if (!String.IsNullOrEmpty(OpenGraphUrl))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Url:</span> <span class\"yammerEmbedWpValue\">" + OpenGraphUrl + "</span></div>\r\n");
			else
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Url:</span> <span class\"yammerEmbedWpValue yammerEmbedWpWarning\">Required</span></div>\r\n");
			

			//OpenGraphType
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Type:</span> <span class\"yammerEmbedWpValue\">" + OpenGraphType + "</span></div>\r\n");

			//ShowOpenGraphPreview
			sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Show Open Graph Preview if available?</span> <span class\"yammerEmbedWpValue\">" + ShowOpenGraphPreview.ToString() + "</span></div>\r\n");
			if (ShowOpenGraphPreview)
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpWarning\"><b>Show Open Graph Preview enabled. May cause JavaScript error from Yammer.</b></span></div>\r\n");

			if (!String.IsNullOrEmpty(OpenGraphTitle))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Title:</span> <span class\"yammerEmbedWpValue\">" + OpenGraphTitle + "</span></div>\r\n");
			else
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Title:</span> <span class\"yammerEmbedWpValue yammerEmbedWpWarning\">(default)</span></div>\r\n");


			if (!String.IsNullOrEmpty(OpenGraphImageUrl))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Image Url:</span> <span class\"yammerEmbedWpValue\">" + OpenGraphImageUrl + "</span></div>\r\n");
			else
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Open Graph Image Url:</span> <span class\"yammerEmbedWpValue yammerEmbedWpWarning\">(default)</span></div>\r\n");
			

			sbEdit.Append("</div>\r\n");

			return sbEdit.ToString();
		}

		private String DrawDisplayPanel()
		{
			String sWrapperClass = GetWrapperClass();
			String sWrapperStyles = GetWrapperStyles();

			StringBuilder sbDisplay = new StringBuilder();
			sbDisplay.Append("<div id=\"" + GetWrapperId() + "\" class=\"yammerEmbedWpPanel " + sWrapperClass + "\" " + sWrapperStyles + "></div>");

			sbDisplay.Append(GenerateYammerEmbedJavaScript());

			return sbDisplay.ToString();
		}

		/*
		 * Function: Generate the Yammer Embed JavaScript based on settings.
		*/
		private String GenerateYammerEmbedJavaScript()
		{
			StringBuilder sbReturn = new StringBuilder();


			sbReturn.Append("<script>\r\n");


			sbReturn.Append("(function() {\r\n");

			//make sure that Yammer Embed JS has been loaded
			sbReturn.Append("if (typeof yam === 'undefined' || !yam || !yam.connect || typeof yam.connect.embedFeed !== 'function') {\r\n");
			sbReturn.Append("   var script = document.createElement('script');\r\n");
			sbReturn.Append("   script.type = 'text/javascript';\r\n");
			sbReturn.Append("   script.src = 'https://c64.assets-yammer.com/assets/platform_embed.js';\r\n");
			sbReturn.Append("   script.async = false;\r\n");
			sbReturn.Append("  	script.onload = loadYammerEmbed;\r\n"); //once script has loaded, we can then try to load yammer embed
			sbReturn.Append("   document.getElementsByTagName('head')[0].appendChild(script);\r\n");
			sbReturn.Append("}\r\n");
			sbReturn.Append("else loadYammerEmbed();\r\n"); //if yam already found, then go ahead and load

			
			sbReturn.Append("function loadYammerEmbed() {\r\n");

			sbReturn.Append("yam.connect.embedFeed(\r\n");
			sbReturn.Append("{\r\n");
			sbReturn.Append("   container: '" + GetContainerSelector() + "',\r\n");

			//NetworkPermalink
			if (!String.IsNullOrEmpty(NetworkPermalink))
				sbReturn.Append("   network: '" + NetworkPermalink + "',\r\n");

			//FeedType
			sbReturn.Append("   feedType: 'open-graph',\r\n");


			//configuration settings
			sbReturn.Append("   config: {\r\n");

			//ShowHeader
			sbReturn.Append("      header: " + ShowHeader.ToString().ToLower() + ",\r\n");

			//HideNetworkName
			sbReturn.Append("      hideNetworkName: " + HideNetworkName.ToString().ToLower() + ",\r\n");

			//PromptText
			if (!String.IsNullOrWhiteSpace(PromptText))
				sbReturn.Append("      promptText: '" + PromptText.ToString().Replace("'", "") + "',\r\n");

			//ShowFooter
			sbReturn.Append("      footer: " + ShowFooter.ToString().ToLower() + ",\r\n");

			//ShowOpenGraphPreview
			sbReturn.Append("      showOpenGraphPreview: " + ShowOpenGraphPreview.ToString().ToLower() + ",\r\n");

			//UseSSO
			sbReturn.Append("      use_sso: " + UseSSO.ToString().ToLower() + "\r\n");

			sbReturn.Append("   },\r\n");

			//objectProperties settings
			sbReturn.Append("   objectProperties: {\r\n");
			
			//url
			if (!String.IsNullOrWhiteSpace(OpenGraphUrl))
				sbReturn.Append("      url: '" + OpenGraphUrl.ToString().Replace("'", "") + "',\r\n");

			//type
			sbReturn.Append("      type: '" + OpenGraphType + "',\r\n");

			//url
			if (!String.IsNullOrWhiteSpace(OpenGraphTitle))
				sbReturn.Append("      title: '" + OpenGraphTitle.ToString().Replace("'", "") + "',\r\n");

			//url
			if (!String.IsNullOrWhiteSpace(OpenGraphImageUrl))
				sbReturn.Append("      image: '" + OpenGraphImageUrl.ToString().Replace("'", "") + "',\r\n");

			sbReturn.Append("   },\r\n");

			sbReturn.Append("});\r\n"); // end yam.connect.embedFeed

			sbReturn.Append("};\r\n"); //end loadYammerEmbed

			sbReturn.Append("})();\r\n");
			sbReturn.Append("</script>\r\n");

			return sbReturn.ToString();
		}

		/*
		 * Function: Get the wrapper Id. Either use default ID based on client id, or if WrapperId provided as an ID, ie starting with #, then use wrapperId
		*/
		private String GetWrapperId()
		{
			String sReturn = "";

			String sID = this.ClientID;

			//if no wrapperId provided with admin, then use client Id
			if (String.IsNullOrEmpty(WrapperId))
				sReturn = sID;
			else
			{
				WrapperId = WrapperId.Trim();

				//if we have a valid Id provided in WrapperId, then we can use
				if (WrapperId.StartsWith("#") && WrapperId.Length > 1)
					sReturn = WrapperId.Substring(1, WrapperId.Length - 1);
				else //otherwise revert to client Id
					sReturn = sID;
			}

			return sReturn.Replace("_", "");
		}

		/*
		 * Function: Get a wrapper class if one provided in WrapperId.
		*/
		private String GetWrapperClass()
		{
			String sReturn = "";

			//if no wrapperId provided with admin, then use client Id
			if (!String.IsNullOrEmpty(WrapperId))
			{
				WrapperId = WrapperId.Trim();

				//if we have a valid Id provided in WrapperId, then we can use
				if (WrapperId.StartsWith(".") && WrapperId.Length > 1)
					sReturn = WrapperId.Substring(1, WrapperId.Length - 1);
			}

			return sReturn;
		}

		/*
		 * Function: Get a wrapper style set including style=""
		*/
		private String GetWrapperStyles()
		{
			StringBuilder sbReturn = new StringBuilder();

			sbReturn.Append("style=\"");

			//set height, with a default height of 400px
			sbReturn.Append("height: " + (WrapperHeight > 0 ? WrapperHeight : 400) + "px; ");

			//set width with a default width of 100%
			if (WrapperWidth > 0)
				sbReturn.Append("width: " + WrapperWidth + "px; ");
			else
				sbReturn.Append("width: 100%; "); //fluid

			sbReturn.Append("\"");

			return sbReturn.ToString();
		}

		/*
		 * Function: Get the container selector we will use. If a class was provided in settings, we will use that. Otherwise if an ID was provide we will use provided ID.
		 * If nothing was provided, we will fall back to clientId.
		 * Function returns leading . or #.
		*/
		private String GetContainerSelector()
		{
			String sReturn = "";

			String sWraperID = GetWrapperId();
			String sWrapperClass = GetWrapperClass();

			if (String.IsNullOrEmpty(sWrapperClass))
				sReturn = "#" + sWraperID;
			else
				sReturn = "." + sWrapperClass;

			return sReturn;
		}
		#endregion
	}
}
