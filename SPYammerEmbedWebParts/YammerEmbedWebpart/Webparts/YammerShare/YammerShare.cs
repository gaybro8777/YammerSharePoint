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

namespace YammerEmbedWebpart.Webparts.YammerShare
{
	/*Yammer Social button documentation: https://developer.yammer.com/v1.0/docs/share-button */
	/*
	 * Yammer share button has no configuration options.
	 * This webpart makes it easy for an end user to add a share button without having to add JavaSrcipt or other script tags.
	 * 
	 */
	[ToolboxItemAttribute(false)]
	public class YammerShare : WebPart
	{
		protected List<string> errorArray = new List<string>();

		#region Custom Webpart properties
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
			sbEdit.Append("<span class=\"yammerActionWpWarning\">Webpart is pre-configured</span>\r\n");

			sbEdit.Append("</div>\r\n");

			return sbEdit.ToString();
		}

		private String DrawDisplayPanel()
		{
			StringBuilder sbDisplay = new StringBuilder();
			sbDisplay.Append("<div id=\"yj-share-button\"></div>");

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
			sbReturn.Append("if (typeof yam === 'undefined' || !yam || !yam.platform || typeof yam.platform.yammerShare !== 'function') {\r\n");
			sbReturn.Append("   var script = document.createElement('script');\r\n");
			sbReturn.Append("   script.type = 'text/javascript';\r\n");
			sbReturn.Append("   script.src = 'https://c64.assets-yammer.com/assets/platform_social_buttons.min.js';\r\n");
			sbReturn.Append("   script.async = false;\r\n");
			sbReturn.Append("  	script.onload = loadYammerSocialShare;\r\n"); //once script has loaded, we can then try to load yammer embed
			sbReturn.Append("   document.getElementsByTagName('head')[0].appendChild(script);\r\n");
			sbReturn.Append("}\r\n");
			sbReturn.Append("else loadYammerSocialShare();\r\n\r\n"); //if yam already found, then go ahead and load


			sbReturn.Append("function loadYammerSocialShare() {\r\n");

			sbReturn.Append("yam.platform.yammerShare();\r\n");

			sbReturn.Append("};\r\n"); // end loadYammerSocialShare

			sbReturn.Append("})();\r\n");
			sbReturn.Append("</script>\r\n");

			return sbReturn.ToString();
		}
		#endregion
	}
}
