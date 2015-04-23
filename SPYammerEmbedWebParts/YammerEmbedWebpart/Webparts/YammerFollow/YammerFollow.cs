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

namespace YammerEmbedWebpart.Webparts.YammerFollow
{
	/*Yammer Action button documentation: https://developer.yammer.com/v1.0/docs/open-graph-buttons */
	/*
	 * When an open graph object is created for a page/file that is "followed", Yammer used embed.ly to pull open graph properties from Title and Meta data
	 * Meta data fields referenced at: https://developer.yammer.com/v1.0/docs/embedly
	 * 
	 */
	[ToolboxItemAttribute(false)]
	public class YammerFollow : WebPart
	{
		protected List<string> errorArray = new List<string>();

		#region Custom Webpart properties

		#region container settings
		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Unique wrapper container selector"),
		WebDescription("An unique selector for this instance, i.e. #embedded-follow. Leave blank for default.")]
		[Category("Yammer Embed Wrapper")]
		public String WrapperId { get; set; }
		#endregion

		#region Embed JavaScript Settings
		[Personalizable(PersonalizationScope.Shared),
		WebBrowsable,
		WebDisplayName("Network Permalink (network)"),
		WebDescription("Your Network, i.e. microsoft.com")]
		[Category("Yammer Embed Settings")]
		public String NetworkPermalink { get; set; }
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

			this.Controls.Add(new LiteralControl("<div class=\"yammerActionWpWrapper\">"));

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
			sbEdit.Append("<div class=\"yammerActionWpEditPanel\">\r\n");
			sbEdit.Append("<span class=\"yammerActionWpWarning\">Edit this web part's properties to modify settings.</span>\r\n");
			//NetworkPermalink
			if (!String.IsNullOrEmpty(NetworkPermalink))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Network Permalink:</span> <span class\"yammerEmbedWpValue\">" + NetworkPermalink + "</span></div>\r\n");
			else
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Network Permalink:</span> <span class\"yammerEmbedWpValue yammerEmbedWpWarning\">Client default network</span></div>\r\n");

			sbEdit.Append("<hr />\r\n");

			//WrapperId
			if (!String.IsNullOrEmpty(WrapperId))
				sbEdit.Append("<div class=\"yammerEmbedWpRow\"><span class=\"yammerEmbedWpLabel\">Custom wrapper element selector:</span> <span class\"yammerEmbedWpValue\">" + WrapperId + "</span></div>\r\n");

			sbEdit.Append("</div>\r\n");

			return sbEdit.ToString();
		}

		private String DrawDisplayPanel()
		{
			String sWrapperClass = GetWrapperClass();

			StringBuilder sbDisplay = new StringBuilder();
			sbDisplay.Append("<div id=\"" + GetWrapperId() + "\" class=\"yammerActionWpPanel " + sWrapperClass + "\"></div>");

			sbDisplay.Append(GenerateYammerActionJavaScript());

			return sbDisplay.ToString();
		}

		/*
		 * Function: Generate the Yammer Embed JavaScript based on settings.
		*/
		private String GenerateYammerActionJavaScript()
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
			sbReturn.Append("  	script.onload = loadYammerActionFollow;\r\n"); //once script has loaded, we can then try to load yammer embed
			sbReturn.Append("   document.getElementsByTagName('head')[0].appendChild(script);\r\n");
			sbReturn.Append("}\r\n");
			sbReturn.Append("else loadYammerActionLike();\r\n\r\n"); //if yam already found, then go ahead and load


			sbReturn.Append("function loadYammerActionFollow() {\r\n");

			sbReturn.Append("yam.connect.actionButton(\r\n");
			sbReturn.Append("{\r\n");
			sbReturn.Append("   container: '" + GetContainerSelector() + "',\r\n");

			//NetworkPermalink
			if (!String.IsNullOrEmpty(NetworkPermalink))
				sbReturn.Append("   network: '" + NetworkPermalink + "',\r\n");

			sbReturn.Append("   action: 'follow'\r\n");
			sbReturn.Append("});\r\n");
			sbReturn.Append("};\r\n"); //end loadYammerActionLike

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
