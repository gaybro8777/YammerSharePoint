'use strict';

/*
var context = SP.ClientContext.get_current();
var user = context.get_web().get_currentUser();

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    getUserName();
});

// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getUserName() {
    context.load(user);
    context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
}

// This function is executed if the above call is successful
// It replaces the contents of the 'message' element with the user name
function onGetUserNameSuccess() {
    $('#message').text('Hello ' + user.get_title());
}

// This function is executed if the above call fails
function onGetUserNameFail(sender, args) {
    alert('Failed to get user name. Error:' + args.get_message());
}
*/

function getQueryStringValue(key) {
	if (!key || key.length < 1)
		return '';

	//get QueryString
	var params = document.URL.split("?")[1].split("&");

	//loop through querystring key/value pairs until we find our key
    for (var i=0; i < params.length; i++)
    {
    	var param = params[i].split('=');
    	if (param[0] === key)
    		return decodeURIComponent(param[1]);
    }

	//always return something
    return '';
}

//Global Variables used in different functions. 

// Extracts the host url and sender Id  values from the query string. 
var gHostUrl = decodeURIComponent(getQueryStringValue("SPHostUrl"));
var gSenderId = decodeURIComponent(getQueryStringValue("SenderId"));

//Yammer Embed Properties
var gNetwork = decodeURIComponent(getQueryStringValue("Network"));
var gFeedType = decodeURIComponent(getQueryStringValue("FeedType"));
var gFeedId = decodeURIComponent(getQueryStringValue("FeedID"));


//Main function to change the size dynamically. 
function ResizeAppPart() {
    if (window.parent == null)
        return;

    var yammerEmbedContainer = document.getElementById('embedded-feed');

    if (yammerEmbedContainer)
    {
    	//use postmessage to resize the app part. 
    	var message = "<Message senderId=" + gSenderId + " >" + "resize(" + yammerEmbedContainer.clientWidth + "," + yammerEmbedContainer.clientHeight + ")</Message>";
    	window.parent.postMessage(message, gHostUrl);
    }

    
}