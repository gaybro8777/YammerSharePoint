'use strict';

var Pxlml = Pxlml || {};

Pxlml.queryString = (function () {

	/**
	 * Given an INPUT "key", get the key's value from the QueryString
	 * 
	 * @param {string} key
	 * @returns {string}
	 */
	function getKeyValue(key) {
		if (!key || key.length < 1)
			return '';

		//get QueryString
		var params = document.URL.split("?")[1].split("&");

		//loop through querystring key/value pairs until we find our key
		for (var i = 0; i < params.length; i++) {
			var param = params[i].split('=');
			if (param[0] === key)
				return decodeURIComponent(param[1]);
		}

		//always return something
		return '';
	};

	return {
		getKeyValue: getKeyValue
	}
})();

Pxlml.appPart = (function () {

	/**
	 * send a message to parent to resize app App based on global vars
	 * 
	 * @param none
	 * @returns null
	 */
	function resize(id) {
		if (window.parent == null)
			return;

		//use postmessage to resize the app part. 
		//if gWrapperWidth is 0 then assume 100% for fluid container. We must provide a valid height though.
		var message = "<Message senderId=" + gSenderId + " >" + "resize(" + ((gWrapperWidth > 0) ? gWrapperWidth + 'px' : '100%') + "," + (gWrapperHeight + "px") + ")</Message>";
		window.parent.postMessage(message, gHostUrl);
	};

	return {
		resize: resize
	}
})();

Pxlml.yammerEmbed = (function () {

	/**
	 * draw the general yammer embed js 
	 * 
	 * @returns {nothing}
	 */
	function draw() {

		if (!gNetworkPermalink || gNetworkPermalink.length < 1)
		{
			document.write('Unable to generate Yammer Feed, Network required.');
			return false;
		}

		//verify variables
		gWrapperId = (!gWrapperId || gWrapperId.length < 1) ? '#embedded-feed' : gWrapperId;
		gWrapperWidth = (!gWrapperWidth || gWrapperWidth < 0) ? 0 : gWrapperWidth; //default width to 0, 0 will be converted to 100% later
		gWrapperHeight = (!gWrapperHeight || gWrapperHeight < 0) ? 0 : gWrapperHeight; //default height to 0, 0 will be converted to 100% later. Will require additional support on client side to set height of containing iframe

		gFeedType = (!gFeedType || gFeedType === 'myfeed') ? '' : gFeedType; //default to '', myfeed is also ''

		//get the actual wrapper id we will use for the container. want to remove # to normalize
		var wrapperId = GetWrapperId();

		//create the embed container first
		var container = '<div id="' + wrapperId + '" style="height: ' + ((gWrapperHeight > 0) ? gWrapperHeight + 'px' : '100%') + '; width: ' + ((gWrapperWidth > 0) ? gWrapperWidth + 'px' : '100%') + ';"></div>';

		//print the container to the page
		document.write(container);

		//set up yammer embed options
		//set up yammer embed commenting options, found that setting feedType when creating options is required
		var options = {
			feedType: gFeedType
		};
		options.container = '#' + wrapperId;
		options.network = gNetworkPermalink;

		if (gFeedId && gFeedId.length > 0)
			options.feedId = gFeedId;
		
		//yammer embed config settings
		options.config = {};

		options.config.use_sso = gUseSSO;							//use Single sign on
		options.config.header = gShowHeader;						//show the Yammer / network header
		options.config.footer = gShowFooter;						//show the yammer footer
		options.config.defaultToCanonical = gDefaultToCanonical;	//should posted message be sent to default network or the network provided. Suggested true
		options.config.hideNetworkName = gHideNetworkName;			//hide the network name in header?
		options.config.defaultGroupId = gDefaultGroupId;			//default group to post a new message to

		if (gPromptText && gPromptText.length > 0)
			options.config.promptText = gPromptText;				//the default prompt text in the new message box

		//yammer embed open graph options
		if (gFeedType === 'open-graph')
		{
			options.config.showOpenGraphPreview = gShowOpenGraphPreview;	//for OG objects, show a preview

			options.objectProperties = {};

			if (gOpenGraphUrl && gOpenGraphUrl.length > 0)
				options.objectProperties.url = gOpenGraphUrl;				//the specific OG Object URL
			if (gOpenGraphType && gOpenGraphType.length > 0)
				options.objectProperties.type = gOpenGraphType;				//the type of OG Object such as page, file, image, etc
			if (gOpenGraphTitle && gOpenGraphTitle.length > 0)
				options.objectProperties.title = gOpenGraphTitle;			//the title of te OG Object
			if (gOpenGraphImageUrl && gOpenGraphImageUrl.length > 0)
				options.objectProperties.image = gOpenGraphImageUrl;		//the preview image url of the OG Object
		}
		
		//go and load yammer request
		yam.connect.embedFeed(options);

		// bind to loading complete Yammer Embed function, will cause App Resize function to fire after load complete
		yam.on('/embed/feed/loadingCompleted', function (context) {
			Pxlml.appPart.resize(wrapperId);
		});
	};



	/**
	 * drawLike the final yammer embed like js 
	 *
	 * Requires both the network and the url to like because the default url will be the app domain url
	 * 
	 * @returns {nothing}
	 */
	function drawLike() {

		if (!gNetworkPermalink || gNetworkPermalink.length < 1) {
			document.write('Unable to generate Yammer Like, Network required.');
			return false;
		}

		//the url to "like" is requried as unable to get the iframe parent url. Could replace with postMessage with some effort to request url from parent
		if (!gOpenGraphUrl || gOpenGraphUrl.length < 1) {
			document.write('Unable to generate Yammer Like, Url of OG Object to like required.');
			return false;
		}

		//verify variables
		gWrapperId = (!gWrapperId || gWrapperId.length < 1) ? '#embedded-like' : gWrapperId;
		
		//get the actual wrapper id we will use for the container. want to remove # to normalize
		var wrapperId = GetWrapperId();

		//create the embed container first
		var container = '<div id="' + wrapperId + '"></div>';

		//print the container to the page
		document.write(container);

		//set up yammer embed like options. Needs the container, network and action.
		var options = {};
		options.container = '#' + wrapperId;
		options.network = gNetworkPermalink;
		options.action = 'like';

		//open graph options can be provided to tell what url is actually being liked. Can also assist in creating OG Object in Yammer if it does not yet exist.
		options.objectProperties = {};

		//OG Type
		if (gOpenGraphType && gOpenGraphType.length > 0)
			options.objectProperties.type = gOpenGraphType;			//you can like any type of OG Object, if one is provided in settings, then use
		else
			options.objectProperties.type = 'page';					//otherwise assume a page

		//OG Url
		options.objectProperties.url = gOpenGraphUrl;				//the specific OG Object URL to "like". Need to provide as otherwise will "like" the app page and not the parent page
		
		//OG Title
		if (gOpenGraphTitle && gOpenGraphTitle.length > 0)
			options.objectProperties.title = gOpenGraphTitle;		//the title of te OG Object

		//go and load yammer request
		yam.connect.actionButton(options);
	};


	/**
	* drawFollow the final yammer embed follow js 
	*
	* Requires both the network and the url to like because the default url will be the app domain url
	* 
	* @returns {nothing}
	*/
	function drawFollow() {

		if (!gNetworkPermalink || gNetworkPermalink.length < 1) {
			document.write('Unable to generate Yammer Follow, Network required.');
			return false;
		}

		//the url to "follow" is requried as unable to get the iframe parent url. Could replace with postMessage with some effort to request url from parent
		if (!gOpenGraphUrl || gOpenGraphUrl.length < 1) {
			document.write('Unable to generate Yammer Follow, Url of OG Object to like required.');
			return false;
		}

		//verify variables
		gWrapperId = (!gWrapperId || gWrapperId.length < 1) ? '#embedded-follow' : gWrapperId;

		//get the actual wrapper id we will use for the container. want to remove # to normalize
		var wrapperId = GetWrapperId();

		//create the embed container first
		var container = '<div id="' + wrapperId + '"></div>';

		//print the container to the page
		document.write(container);

		//set up yammer embed follow options. Needs the container, network and action.
		var options = {};
		options.container = '#' + wrapperId;
		options.network = gNetworkPermalink;
		options.action = 'follow';

		//open graph options can be provided to tell what url is actually being followed. Can also assist in creating OG Object in Yammer if it does not yet exist.
		options.objectProperties = {};

		//OG Type
		if (gOpenGraphType && gOpenGraphType.length > 0)
			options.objectProperties.type = gOpenGraphType;			//you can like any type of OG Object, if one is provided in settings, then use
		else
			options.objectProperties.type = 'page';					//otherwise assume a page

		//OG Url
		options.objectProperties.url = gOpenGraphUrl;				//the specific OG Object URL to "follow". Need to provide as otherwise will "follow" the app page and not the parent page

		//OG Title
		if (gOpenGraphTitle && gOpenGraphTitle.length > 0)
			options.objectProperties.title = gOpenGraphTitle;		//the title of te OG Object

		//go and load yammer request
		yam.connect.actionButton(options);
	};



	/**
	 * drawCommenting the final yammer embed commenting js 
	 * 
	 * @returns {nothing}
	 */
	function drawComment() {

		if (!gNetworkPermalink || gNetworkPermalink.length < 1) {
			document.write('Unable to generate Yammer Comment Feed, Network required.');
			return false;
		}

		//verify variables
		gWrapperId = (!gWrapperId || gWrapperId.length < 1) ? '#embedded-comment' : gWrapperId;
		gWrapperWidth = (!gWrapperWidth || gWrapperWidth < 0) ? 0 : gWrapperWidth; //default width to 0, 0 will be converted to 100% later
		gWrapperHeight = (!gWrapperHeight || gWrapperHeight < 0) ? 0 : gWrapperHeight; //default height to 0, 0 will be converted to 100% later. Will require additional support on client side to set height of containing iframe

		//get the actual wrapper id we will use for the container. want to remove # to normalize
		var wrapperId = GetWrapperId();

		//create the embed container first
		var container = '<div id="' + wrapperId + '" style="height: ' + ((gWrapperHeight > 0) ? gWrapperHeight + 'px' : '100%') + '; width: ' + ((gWrapperWidth > 0) ? gWrapperWidth + 'px' : '100%') + ';"></div>';

		//print the container to the page
		document.write(container);

		//set up yammer embed commenting options, found that setting feedType when creating options is required
		var options = {
			feedType: 'open-graph'
		};
		
		options.container = '#' + wrapperId;
		options.network = gNetworkPermalink;

		//yammer embed config settings
		options.config = {};

		options.config.use_sso = gUseSSO;							//use Single sign on
		options.config.header = gShowHeader;						//show the Yammer / network header
		options.config.footer = gShowFooter;						//show the yammer footer
		options.config.hideNetworkName = gHideNetworkName;			//hide the network name in header?

		if (gPromptText && gPromptText.length > 0)
			options.config.promptText = gPromptText;				//the default prompt text in the new message box

		//yammer embed open graph options
		options.config.showOpenGraphPreview = gShowOpenGraphPreview;	//for OG objects, show a preview?

		options.objectProperties = {};

		if (gOpenGraphUrl && gOpenGraphUrl.length > 0)
			options.objectProperties.url = gOpenGraphUrl;				//the specific OG Object URL
		if (gOpenGraphType && gOpenGraphType.length > 0)
			options.objectProperties.type = gOpenGraphType;				//the type of OG Object such as page, file, image, etc
		if (gOpenGraphTitle && gOpenGraphTitle.length > 0)
			options.objectProperties.title = gOpenGraphTitle;			//the title of te OG Object
		if (gOpenGraphImageUrl && gOpenGraphImageUrl.length > 0)
			options.objectProperties.image = gOpenGraphImageUrl;		//the preview image url of the OG Object

		//go and load yammer request
		yam.connect.embedFeed(options);

		// bind to loading complete Yammer Embed function, will cause App Resize function to fire after load complete
		yam.on('/embed/feed/loadingCompleted', function (context) {
			Pxlml.appPart.resize(wrapperId);
		});
	};



	/**
	 * Normalize the wrapperid by removeing a preceding # if one is found
	 * 
	 * @returns {string} - the container id without a leading #
	 */
	function GetWrapperId() {
		if (!gWrapperId || gWrapperId.length < 2)
			return '';

		if (gWrapperId[0] === '#')
			return gWrapperId.substring(1);
		else
			return gWrapperId;
	};


	return {
		draw: draw,
		drawLike: drawLike,
		drawFollow: drawFollow,
		drawComment: drawComment
	}
})();





Pxlml.yammerSocial = (function () {

	/**
	* drawShare the yammer share js 
	*
	* @returns {nothing}
	*/

	/*
		NOTE: Yammer Share as an App does not work well at this time.
		the Yammer function yam.platform.yammerShare() does not allow for the input of options such as a url
		yammerShare will use the url of the app web and not the true window url.
		Provided here as an example, but in production will likely not work as intended as the "Shared" Url will be incorrect.
	*/
	function drawShare() {
		//create the embed container first
		var container = '<div id="yj-share-button"></div>';

		//print the container to the page
		document.write(container);

		//go and load yammer request
		yam.platform.yammerShare();
	};


	return {
		drawShare: drawShare,
	}
})();


//Global Variables used in different functions. 
//var context = SP.ClientContext.get_current();
//var user = context.get_web().get_currentUser();

// Extracts the host url and sender Id  values from the query string. 
var gHostUrl = decodeURIComponent(Pxlml.queryString.getKeyValue("SPHostUrl"));
var gSenderId = decodeURIComponent(Pxlml.queryString.getKeyValue("SenderId"));

//Yammer Embed Properties
var gWrapperId = decodeURIComponent(Pxlml.queryString.getKeyValue("WrapperId"));
var gWrapperWidth = decodeURIComponent(Pxlml.queryString.getKeyValue("WrapperWidth"));
var gWrapperHeight = decodeURIComponent(Pxlml.queryString.getKeyValue("WrapperHeight"));
var gNetworkPermalink = decodeURIComponent(Pxlml.queryString.getKeyValue("NetworkPermalink"));
var gFeedType = decodeURIComponent(Pxlml.queryString.getKeyValue("FeedType"));
var gFeedId = decodeURIComponent(Pxlml.queryString.getKeyValue("FeedID"));
var gShowHeader = decodeURIComponent(Pxlml.queryString.getKeyValue("ShowHeader"));
var gHideNetworkName = decodeURIComponent(Pxlml.queryString.getKeyValue("HideNetworkName"));
var gPromptText = decodeURIComponent(Pxlml.queryString.getKeyValue("PromptText"));
var gDefaultGroupId = decodeURIComponent(Pxlml.queryString.getKeyValue("DefaultGroupId"));
var gDefaultToCanonical = decodeURIComponent(Pxlml.queryString.getKeyValue("DefaultToCanonical"));
var gShowFooter = decodeURIComponent(Pxlml.queryString.getKeyValue("ShowFooter"));
var gUseSSO = decodeURIComponent(Pxlml.queryString.getKeyValue("UseSSO"));
var gOpenGraphType = decodeURIComponent(Pxlml.queryString.getKeyValue("OpenGraphType"));
var gOpenGraphUrl = decodeURIComponent(Pxlml.queryString.getKeyValue("OpenGraphUrl"));
var gShowOpenGraphPreview = decodeURIComponent(Pxlml.queryString.getKeyValue("ShowOpenGraphPreview"));
var gOpenGraphTitle = decodeURIComponent(Pxlml.queryString.getKeyValue("OpenGraphTitle"));
var gOpenGraphImageUrl = decodeURIComponent(Pxlml.queryString.getKeyValue("OpenGraphImageUrl"));