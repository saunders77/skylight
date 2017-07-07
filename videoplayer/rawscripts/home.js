﻿function testAPIs() {
     write("hello2");
     Office.context.document.settings.set('mySetting', 2);
     Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
    });
}

function write(myText){
    document.getElementById("debug").innerHTML = document.getElementById("debug").innerHTML + "\n" + myText;
}

// from stackoverflow
function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

// user license info
var orgId;
var liveId;
var userId; // can become either the orgId or the liveId
var acqusitionDate;
var isPro = false;

var pingingForPayment = false;

function loadLicenseInfo(){
	write("loading license info");
	try{
		var tokenXml = atob(decodeURIComponent(getParameterByName("et")));
		var cleanTokenXml = "";
		var nul = String.fromCharCode(00);
		for(var i = 0;i < tokenXml.length;i++){
		  if(tokenXml[i] != nul){
			cleanTokenXml += tokenXml[i];
		  }
		}
		var $t = $($.parseXML(cleanTokenXml)).find("t");
		orgId = $t.attr("oid");
		liveId = $t.attr("cid");
		userId = orgId;
		
		if(liveId){
			userId = liveId;
		}
		acquisitionDate = $t.attr("ad");
		
		ga('set', 'userId', userId);
	} catch(err){
		write("Error licensing: " + err.message);
		ga("send","event","videoplayer","licensing",err.message);
	}
	
	// these literal assignments are for testing
	//liveId = "gorgey-our";
	//userId = liveId;

	write("user ID is " + userId);

}

function errorMessage(myText){
     document.getElementById("innerErrorDiv").innerHTML = myText;
     var theErrorDiv = document.getElementById("errorDiv");
     //theErrorDiv.style.visibility = "visible";
     $('#errorDiv').fadeIn();
}

function createVideo(){
	var urlString;
	urlString = Office.context.document.settings.get("vid");
	var myAutoplay = 0;
	var myStartTime = 0;
	if(urlString){
		// then we're loading from cache
		myAutoplay = Office.context.document.settings.get("autoplay");
		myStartTime = Office.context.document.settings.get("starttime");
	}
	else{
		// construct the video parameters
		urlString = document.getElementById("videoID").value;

		var timeArr = document.getElementById("timeinput").value.split(':');
		if(document.getElementById("autoplay").checked){
			myAutoplay = 1;
		}
		if(document.getElementById("customstarttime").checked){
			myStartTime = (+timeArr[timeArr.length - 1]) + ((+timeArr[timeArr.length - 2]) * 60);
			if(timeArr.length > 2){
				myStartTime += ((+timeArr[timeArr.length - 3]) * 60 * 60);
			}
		}
		write("setting video params");
	}
	
    if(urlString.indexOf("youtube.com") != -1 || urlString.indexOf("youtu.be") != -1)
    {
          ga('send','event','videoplayer','setvideo','youtube');
		  write("there's a video to create");
		  Office.context.document.settings.set("vid",urlString);
		  Office.context.document.settings.set("autoplay", myAutoplay);
		  Office.context.document.settings.set("starttime", myStartTime);
          Office.context.document.settings.saveAsync(function (asyncResult) {
            write('Settings saved with status: ' + asyncResult.status);
            if(true){//window.top==window){
                 // not in iFrame
                 document.getElementById('player').style.visibility = 'visible';
                 var script = document.createElement( 'script' );
                 script.type = 'text/javascript';
                 script.src = "../scripts/youtube.js";
                 $('body').append( script );
                 
            }
            else{
                // window.reload("www.youtube.com/watch?v=TJLY4Cgk18U");
                 $('#iframed').fadeIn();
                 $('#iframed').addClass("inWAC");
                 document.getElementById("iframed").style.visibility = 'visible';
                 //document.getElementById("iframed").href = Office.context.document.settings.get("vid");
                 write("iframedo");
                 
                 $('#iframed').click(function(){
                      window.open(Office.context.document.settings.get("vid"));
                 });
                 //document.getElementById("iframed").innerHTML = "<p>YouTube does not allow video-embedding within Office Online. Please use this document with the desktop version of Office to view the video.</p>";           
            
            }
          });
    }
    else if(urlString.indexOf("vimeo.com") != -1){
          ga('send','event','videoplayer','setvideo','vimeo');
		  Office.context.document.settings.set("vid",urlString);
		  Office.context.document.settings.set("autoplay", myAutoplay);
		  Office.context.document.settings.set("starttime", myStartTime);
          Office.context.document.settings.saveAsync(function (asyncResult) {
               if(true){//window.top==window){
               // not in iFrame
               write("creating vimeox");
               var vindex = urlString.indexOf("meo.com/");
               var vid = urlString.substring(vindex+8);
               var ifrm = document.getElementById('ifrm');
               ifrm.style.height = "100%";
               write("heighta: " + ifrm.height);
			   var queryString = "?";
			   if(myStartTime){
				   queryString += "&#t=" + myStartTime + "s";
			   }
			   if(myAutoplay){
				   queryString += "&autoplay=1";
			   }
			   queryString += "&title=0&amp;byline=0&amp;portrait=0";
			   
               ifrm.setAttribute("src","//player.vimeo.com/video/" + vid + queryString);
               write("heighto: " + ifrm.height);
               write(ifrm.style.width);
               write("zindex: " + ifrm.style.zIndex);
            }
            else{
                var vindex = urlString.indexOf("meo.com/");
               var vid = urlString.substring(vindex+8);
                window.location.href = "//player.vimeo.com/video/" + vid + "?title=0&amp;byline=0&amp;portrait=0";
                /*
                // window.reload("www.youtube.com/watch?v=TJLY4Cgk18U");
                 $('#iframed').fadeIn();
                 $('#iframed').addClass("inWAC");
                 document.getElementById("iframed").style.visibility = 'visible';
                 document.getElementById("iframed").href = Office.context.document.settings.get("vid");
                 write("iframedo");
                 $('#iframed').click(function(){
                      window.open(Office.context.document.settings.get("vid"));
                 });
                 //document.getElementById("iframed").innerHTML = "<p>YouTube does not allow video-embedding within Office Online. Please use this document with the desktop version of Office to view the video.</p>";           
            */
            }
          });
          
    }
    else if(urlString.indexOf("liveleakjhkjkljkl;jkl.com") != -1){
          // 2c07fcfdbe0742afaf70103d70e72a3b
          var embedLink = document.createElement("a");
          embedLink.setAttribute()
          
          Office.context.document.settings.set("vid",urlString);
          Office.context.document.settings.saveAsync(function (asyncResult) {
          
          });
    }
    else{
          errorMessage("Choose a valid URL for your video.");
    }
    
}

function saveVid(){
     
}

Office.initialize = function (reason) {
    $(document).ready(function(){
		
		if(true){//window.top==window){
            //not in iframe
            /*
			if(document.getElementById("links").innerHTML.indexOf("Rate") == -1){
               document.getElementById("links").innerHTML += "<a href='https://store.office.com/writereview.aspx?assetid=WA104221182'>Rate</a>";
            }
            */
            $('#iframed').fadeOut()
            
        }
        else if(Office.context.document.settings.get("vid"))
        {
            
            document.getElementById("iframed").style.visibility = 'visible';
               document.getElementById("iframed").href = Office.context.document.settings.get("vid");
            $('#iframed').addClass("inWAC");
            $('#iframed').click(function(){
               window.open(Office.context.document.settings.get("vid"));
            });
            //document.getElementById("iframed").innerHTML = "<p>YouTube does not allow video-embedding within Office Online. Please use this document with the desktop version of Office to view the video.</p>";

        }
        $('#setVid').click(function(){
            write("creating video");
            createVideo();
        });
        $('#errorDiv').click(function(){
        	$(this).fadeOut();
        });
		$('.payButton').click(function(){
			window.open("../pages/purchasewindow.html?custom=" + encodeURIComponent(userId));
			pingingForPayment = true;
			$('#waitingPay').show();
			setTimeout(pingForPro,10000);
			ga("send","event","videoplayer","checkping");
		});
		$('#cancelPay').click(function(){
        	$('#waitingPay').hide();
			pingingForPayment = false;
        });
        
		function pingForPro(){
			checkServerDatabase(function(myStatus){
				if(myStatus == 200){
					write("result succeeded");
					ga("send","event","videoplayer","purchasesucceeded");
					pingingForPayment = false;
					$('#proPrompt').fadeOut();
					turnOnPro();
				}
				else{
					if(pingingForPayment){
						setTimeout(pingForPro,2000);
						ga("send","event","videoplayer","checkping");					
					}
				}
			});
		}

		function turnOnPro(){
			$('.startsDisabled').prop("disabled", false);
			$('#timeinput').css("user-select", "text")
		}

		function showAd(){
			$('#proPrompt').fadeIn();
		}

        if(Office.context.document.settings.get("vid")){
			//document.getElementById("videoID").value = Office.context.document.settings.get("vid");
            ga('send','event','videoplayer','loadplayer','existingvideo');
			createVideo();
        }
        else{			
			
			
			function dialogCallback(asyncResult) {
				if (asyncResult.status == "failed") {

					// In addition to general system errors, there are 3 specific errors for 
					// displayDialogAsync that you can handle individually.
					switch (asyncResult.error.code) {
						case 12004:
							write("Domain is not trusted");
							break;
						case 12005:
							write("HTTPS is required");
							break;
						case 12007:
							write("A dialog is already opened.");
							break;
						default:
							write(asyncResult.error.message);
							break;
					}
				}
				else {
					dialog = asyncResult.value;
					/*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
					dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

					/*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
					dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
				}
			}

			
			ga('send','event','videoplayer','loadplayer','novideo');
			
			loadLicenseInfo();
			
			function checkServerDatabase(callback){
				var xhttp = new XMLHttpRequest();
				xhttp.onreadystatechange = function() {
					if (this.readyState == 4) {
						callback(this.status);
					}
				};
				xhttp.open("POST", "https://michael-saunders.com/server/checkdatabase.php", true);
				xhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
				xhttp.send("custom=" + userId);
			}
			
			checkServerDatabase(function(myStatus){
				if(myStatus == 200){
					write("result succeeded");
					turnOnPro();
				}
				else{
					if(userId && Office.context.commerceAllowed){
						showAd();
					}
					
					write("result status: " + myStatus)
				}
			});

            document.getElementById("cloak").style.visibility = 'hidden';
            
			$("#videoID").focus();
			
			
        }
    });
    
};