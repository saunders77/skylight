function testAPIs() {
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
	liveId = "888888-88888-88888";
	userId = liveId;

}

function errorMessage(myText){
     document.getElementById("innerErrorDiv").innerHTML = myText;
     var theErrorDiv = document.getElementById("errorDiv");
     //theErrorDiv.style.visibility = "visible";
     $('#errorDiv').fadeIn();
}

function createVideo(){
    var urlString = document.getElementById("videoID").value;
	var myAutoplay = 0
	var myStartTime = 0;
	var timeArr = document.getElementById("timeinput").split(':');
	if(document.getElementById("autoplay").checked){
		myAutoplay = 1;
	}
	if(document.getElementById("customstarttime").checked){
		myStartTime = (+timeArr[timeArr.length - 1]) + ((+timeArr[timeArr.length - 2]) * 60);
		if(timeArr.length > 2){
			myStartTime += ((+timeArr[timeArr.length - 3]) * 60 * 60);
		}
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
                 $('#cloak').fadeOut();
                 $('#cloak').remove();
                 document.getElementById("cloak").style.visibility = 'hidden';
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
          Office.context.document.settings.saveAsync(function (asyncResult) {
               if(true){//window.top==window){
               // not in iFrame
               write("creating vimeox");
               var vindex = urlString.indexOf("meo.com/");
               var vid = urlString.substring(vindex+8);
               var ifrm = document.getElementById('ifrm');
               ifrm.style.height = "100%";
               write("heighta: " + ifrm.height);
               ifrm.setAttribute("src","//player.vimeo.com/video/" + vid + "?title=0&amp;byline=0&amp;portrait=0");
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
		$('#buyNow').click(function(){
			window.open("../pages/purchasewindow.html?custom=" + encodeURIComponent(userId));
		});
        
		function turnOnPro(){
			$('.startsDisabled').prop("disabled", false);
			$('.startsDisabled').css("color", black);
			$('#premiumFeatures').css("color", black);
			$('#premiumFeatures').css("border-color", black);
		}

		function showAd(){
			$('#proPrompt').fadeIn();
		}

        if(Office.context.document.settings.get("vid")){
			document.getElementById("videoID").value = Office.context.document.settings.get("vid");
            ga('send','event','videoplayer','loadplayer','existingvideo');
			createVideo();
        }
        else{
            
			// checks for the userId of paid users
			function checkFirebase(customerId,callback){
				write("checking firebase");
				$.getScript("https://www.gstatic.com/firebasejs/3.2.1/firebase.js", function(response, status){
					// code directly from Firebase
					var config = {
						apiKey: "AIzaSyAmbxHuUjquac2ltM5hFoHqSIFe9bLN9u0",
						authDomain: "web-video-firebase.firebaseapp.com",
						databaseURL: "https://web-video-firebase.firebaseio.com",
						storageBucket: "web-video-firebase.appspot.com",
						// not sure why the following line is needed here but not for Stock Connector
						messagingSenderId: "915381707394"
					};
					firebase.initializeApp(config);
					
					// my code
					var database = firebase.database();
					var userReference = firebase.database().ref("customers/" + customerId);
					userReference.once('value').then(function(dataSnapshot) {
						// handle read data.
						if(dataSnapshot.val()){
							callback(true);
							ga("send","event","videoplayer","confirmedFirebase",customerId);
						}
						else{
							callback(false);
						}
					});
				});
				
			}
			
			function checkForPro(callback){
				write("checking for pro");
				// don't need to check the document
				// uses the code names "weight" and "light"
				
				// check localStorage
				if (typeof(Storage) !== "undefined" && localStorage.getItem("weight") == "light") {
					callback(true);
				}
				else{
					// third, check firebase if the user is logged in
					if(userId){
						checkFirebase(userId,function(paid){
							if(paid){
								callback(true);
								localStorage.setItem("weight","light");
							}
							else{
								callback(false);
							}
						});
					}
					else{
						callback(false);
					}					
				}
				
			}
			
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

			//displayProAd
			write("should I display the ad? " + (userId && Office.context.commerceAllowed));
			
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
					showAd();
					write("result status: " + myStatus)
				}
			});

			write("sent request");

			if(userId && Office.context.commerceAllowed){
				$("#premiumFeatures").show();
			}
			
			//$('#cloak').fadeOut();
            //$('#cloak').remove();
            document.getElementById("cloak").style.visibility = 'hidden';
            
			$("#videoID").focus();
			
			
			$("#autoplay").change(function(){
				if(this.checked){
					$("#loadingGif").show();
					checkForPro(function(result){
						if(result){
							// the user is pro
							$("#loadingGif").hide();
						}
						else{
							write("I'm not pro");
							//Office.context.ui.displayDialogAsync("https://michael-saunders.com/videoplayer/xstaging/pages/dialog.html", { height: 50, width: 10 }, dialogCallback);
						}
					});
				}
				else{
					// the user just turned off autoplay
					Office.context.document.settings.set("weight","light");
					Office.context.document.settings.saveAsync();
				}
			});
			
			// show the ad
			
			//document.getElementById("goog").style.display = "inline";
			/*
			var script1 = document.createElement('script');
			script1.async = "async";
			script1.src = "//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js";
			var ins1 = document.createElement('ins');
			ins1.style = "display:block";
			ins1.data-ad-client="ca-pub-6300181586260439";
			ins1.data-ad-slot="4285461505";
			ins1.data-ad-format="auto";
			var script2 = document.createElement('script');
			script2.innerHTML = "(adsbygoogle = window.adsbygoogle || []).push({});"
			document.getElementById('goog').append(script1);
			document.getElementById('goog').append(in1);
			document.getElementById('goog').append(script2);
			
			*/
			//document.body.style.visibility = 'visible';
            //Office.context.document.settings.set("vid","2hCg3OptVCs");
           // document.getElementById("videoID").value = "2hCg3OptVCs";
            /*Office.context.document.settings.saveAsync(function (asyncResult) {
                write('Settings saved with status: ' + asyncResult.status);
                //just this time
               // createVideo(); 
            });      */     
			
			
        }
    });
    
};