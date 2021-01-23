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
/*
function regularInfo(){
  
  write("oHeight:" + outerHeight  + "; ");
  write("oWidth:" + outerWidth + "; ");
  write("iHeight:" + innerHeight + "; ");
  write("iWidth:" + innerWidth + "; ");
  write("pageX:" + pageXOffset + "; ");
  write("pageY:" + pageYOffset + "; ");
  write("sHeight:" + screen.height  + "; ");
  write("sWidth:" + screen.width + "; ");
  write("aHeight:" + screen.availHeight + "; ");
  write("aWidth:" + screen.availWidth + "; ");
  
  write("sLeft:" + screenLeft  + "; ");
  write("sTop:" + screenTop + "; ");
  write("sX:" + screenX + "; ");
  write("sY:" + screenY + "; ");

  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{}, function (asyncResult) {
				var error = asyncResult.error;
				if (asyncResult.status === Office.AsyncResultStatus.Failed) {
					write(error.name + ": " + error.message);
				} 
				else {
					// Get selected data.
					var dataValue = asyncResult.value; 
					write('Selected data is ' + JSON.stringify(dataValue));
				}            
			});
  write("<br>");
  
  
  setTimeout(regularInfo, 2000);
	
	}

window.onscroll = function(){write("scrolled<br>");};
window.onselect = function(){write("selected<br>");};
window.onresize = function(){write("resized<br>");};
window.onpointermove = function(){write("pointermoved<br>");};
*/
//setTimeout(regularInfo, 2000);

// user license info
var orgId;
var liveId;
var userId; // can become either the orgId or the liveId
var acqusitionDate;
var isPro = false;

var readyToRegister = false;
var signedIn = false;
var pass;

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
	//liveId = "1EF5C754CE1B2ADf";
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
	function rememberSuccess(){
		if (typeof(Storage) !== "undefined") {
			if(!localStorage.getItem("lastUseDay")){
				var d = new Date();
				localStorage.setItem("lastUseDay", d.toDateString());
				localStorage.setItem("usedDays", 1);
			}
		} else {
			write("Sorry, your browser does not support Web Storage...");
		}
	}	

	var urlString;
	urlString = Office.context.document.settings.get("vid");
	var myAutoplay = 0;
	var myStartTime = 0;
	var myEndTime = 0;
	if(urlString){
		// then we're loading from cache
		myAutoplay = Office.context.document.settings.get("autoplay");
		myStartTime = Office.context.document.settings.get("starttime");
		myEndTime = Office.context.document.settings.get("endtime");
	}
	else{
		// construct the video parameters
		urlString = document.getElementById("videoID").value;

		var timeArr = document.getElementById("timeinput").value.split(':');
		var endTimeArr = document.getElementById("endtimeinput").value.split(':');
		if(document.getElementById("autoplay").checked){
			myAutoplay = 1;
		}
		if(document.getElementById("customstarttime").checked){
			myStartTime = (+timeArr[timeArr.length - 1]) + ((+timeArr[timeArr.length - 2]) * 60);
			if(timeArr.length > 2){
				myStartTime += ((+timeArr[timeArr.length - 3]) * 60 * 60);
			}
		}
		if(document.getElementById("customendtime").checked){
			myEndTime = (+endTimeArr[endTimeArr.length - 1]) + ((+endTimeArr[endTimeArr.length - 2]) * 60);
			if(endTimeArr.length > 2){
				myEndTime += ((+endTimeArr[endTimeArr.length - 3]) * 60 * 60);
			}
		}
		write("setting video params");
	}

	// start checking to reload the window when the user leaves to another slide
	if(Office.context.document.settings.get("slideId")){
		var playingSlideId = Office.context.document.settings.get("slideId");
		function watchForNextSlide(){
			Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{}, function (asyncResult) {
				var error = asyncResult.error;
				if (asyncResult.status === Office.AsyncResultStatus.Failed) {
					write(error.name + ": " + error.message);
				} 
				else {
					if(asyncResult.value["slides"][0]["id"] == playingSlideId){
						// we're still playing
						setTimeout(watchForNextSlide, 600);
					}
					else{
						// the user has moved away
						window.location.reload();
					}
				}            
			});
		}

		watchForNextSlide();
	}
	// otherwise it's a legacy video, so just leave it to keep playing
	
    if(urlString.indexOf("youtube.com") != -1 || urlString.indexOf("youtu.be") != -1)
    {
          ga('send','event','videoplayer','setvideo','youtube');
		  rememberSuccess();
		  write("there's a video to create");
		  Office.context.document.settings.set("vid",urlString);
		  Office.context.document.settings.set("autoplay", myAutoplay);
		  Office.context.document.settings.set("starttime", myStartTime);
		  Office.context.document.settings.set("endtime", myEndTime);
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
		  rememberSuccess();
		  Office.context.document.settings.set("vid",urlString);
		  Office.context.document.settings.set("autoplay", myAutoplay);
		  Office.context.document.settings.set("starttime", myStartTime);
		  Office.context.document.settings.set("endtime", myEndTime);
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
			   // vimeo doesn't support end times
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
            if(document.getElementById("videoID").value == "debug"){
				errorMessage("Click <a href='mailto:webvideoplayer@outlook.com?subject=Support Request for " + userId + "&body=Please enable my account. Thank you!'>here</a> to send ID code " + userId + " for support.");
			}
			else{
				var waitingForSlideId = true;
				
				Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{}, function (asyncResult) {
					if(waitingForSlideId){
						var error = asyncResult.error;
						write("selecteddata result is " + asyncResult.status);
						if (asyncResult.status === Office.AsyncResultStatus.Failed) {
							write(error.name + ": " + error.message);
							createVideo();
						} 
						else {
							// remember which slide it's on
							var slideId = asyncResult.value["slides"][0]["id"]; 
							Office.context.document.settings.set("slideId", slideId);
							Office.context.document.settings.saveAsync(function (asyncResult) {
								createVideo();
							});
						}
					}
					// else do nothing because we created the video without a slide ID already   
				});

				// the above call never returns if permissions doesn't allow it. So let's give it 400ms:
				setTimeout(function(){
					waitingForSlideId = false;
					createVideo();
				}, 400);			
			}
			
        });
        $('#errorDiv').click(function(){
        	$(this).fadeOut();
        });
		$('#timeinput').change(function(){
        	$('#customstarttime').attr('checked', true);
        });
		// some weird bug preventing me from doing this with jquery
		document.getElementById("endtimeinput").onchange = function(){
			document.getElementById("customendtime").checked = true;
		};
		$('.payButton').click(function(){
			if(userId){
				// then we can just go to PayPal
				window.open("../pages/purchasewindow.html?custom=" + encodeURIComponent(userId));
				pingingForPayment = true;
				$('#waitingPay').show();
				setTimeout(pingForPro,10000);
				ga("send","event","videoplayer","checkping");
			}
			else{
				
			}
			
		});

		$('#email').focus(function(){
			if($('#email').val() == "email"){
				$('#email').attr({
					"value": "",
				});
			}
		});

		$('#password').focus(function(){
			if($('#password').val() == "password"){
				$('#password').attr({
					"value": "",
				});
				// this part is to replace the password field
				var oldInput = document.getElementById("password");
				var newInput= oldInput.cloneNode(false);
				newInput.type='password';
				oldInput.parentNode.replaceChild(newInput,oldInput);
				$('#password').focus();
				var waitToCheck; // timeout
				function checkPassReady(){
					var emailString = $('#email').val();
					if($('#password').val().length > 3 && emailString.length > 3 && emailString.indexOf('@') > -1 && emailString.indexOf('.') > -1){
						waitToCheck = setTimeout(function(){
							// check database
							pass = sha256("wvp" + $('#password').val());
							write("pass=" + pass);
							userId = $('#email').val();
							signedIn = true;
							// check whether the user is registered
							checkServerDatabase(function(myStatus){
								if(myStatus == 200){
									write("found the id");
									$('#proPrompt').fadeOut();
									turnOnPro();
									$("#signinlink").fadeOut();
									write("hiding sign in");											
									// cache locally
									localStorage.setItem("email", userId);
								}
								else{
									write("not found");
									$("#registerSignIn").css("background-color","#bbe2fa");
									readyToRegister = true;
								}
							},pass);
							
						},500);
					}
					else{
						$("#registerSignIn").css("background-color","#c8c8c8");
						readyToRegister = false;	
					}
				}
				$('.registerBox').keyup(function(){
					readyToRegister = false;
					$("#registerSignIn").css("background-color","#c8c8c8");
					if(waitToCheck){
						clearInterval(waitToCheck);
					}
					checkPassReady();
				});

				$('#registerSignIn').click(function(){
					if(readyToRegister){								
						// trigger paypal flow

						window.open("../pages/purchasewindowpw.html?myUserId=" + encodeURIComponent(userId) + "&myPass=" + encodeURIComponent(pass));
						pingingForPayment = true;
						$('#waitingPay').show();
						setTimeout(function(){
							pingForPro(pass);
						},10000);
						ga("send","event","videoplayer","checkpingpw");				
					}
				});

			}
		});

		$('#cancelPay').click(function(){
        	$('#waitingPay').hide();
			pingingForPayment = false;
        });

		//help buttons
		$('#contactlink').click(function(){
        	window.open("mailto:webvideoplayer@outlook.com");
        });
		$('#helplink').click(function(){
        	window.open("https://www.michael-saunders.com/videoplayer/pages/info.html#howto");
        });
		$('#privacylink').click(function(){
        	window.open("https://www.michael-saunders.com/videoplayer/pages/privacy.html");
        });
		$('#ratelink').click(function(){
        	window.open("https://store.office.com/writereview.aspx?assetid=WA104221182");
        });
		$('#signinlink').click(function(){
        	$.getScript("https://michael-saunders.com/videoplayer/scripts/sha256.js",function(response,status){});
			$('#proContent').fadeOut(200);
			$('#buyNow').fadeOut(200);
			$('.registering').fadeOut(200);
			setTimeout(function(){
				$('#signInText').html("Sign in to Web Video Player to access your premium features.");
				$('#registerSignInText').html("Sign in");
				$('#email').html("email");
				$('.registering').fadeIn(200);
			},250);
			
        });
        
		function pingForPro(password){
			checkServerDatabase(function(myStatus){
				if(myStatus == 200){
					if(signedIn){
						localStorage.setItem("email", userId);
					}
					write("result succeeded");
					ga("send","event","videoplayer","purchasesucceeded");
					pingingForPayment = false;
					$('#proPrompt').fadeOut();
					turnOnPro();
				}
				else{
					if(pingingForPayment){
						setTimeout(function(){
							pingForPro(password);
						},2000);
						ga("send","event","videoplayer","checkping");					
					}
				}
			},password);
		}

		function turnOnPro(){
			$('.startsDisabled').prop("disabled", false);
			$('#timeinput').css("user-select", "text");
			$('#endtimeinput').css("user-select", "text");
		}

		function showAd(){
			$('#proPrompt').fadeIn();
		}

		// add log to localStorage
		if (typeof(Storage) !== "undefined") {
			if(localStorage.getItem("lastUseDay")){
				var d = new Date();
				write("my used days=" + localStorage.getItem("usedDays"));
				write("datestring=" + d.toDateString());
				if(localStorage.getItem("lastUseDay") != d.toDateString()){
					// then we can log a new unique day
					localStorage.setItem("lastUseDay", d.toDateString());
					var myUsedDays = localStorage.getItem("usedDays");
					myUsedDays++;
					localStorage.setItem("usedDays", myUsedDays);
					if(myUsedDays > 1){
						write("show rating link");

						// disabling for now while ratings don't work
						//$("#ratelink").css("visibility", "visible");
					}
				}
			}
		} else {
			write("Sorry, your browser does not support Web Storage...");
		}

        if(Office.context.document.settings.get("vid")){
            ga('send','event','videoplayer','loadplayer','existingvideo');

			if(Office.context.document.settings.get("slideId")){
				// verify that the slide is active before creating the video

				var savedSlideId = Office.context.document.settings.get("slideId");

				function checkActiveSlide(){
					Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{}, function (asyncResult) {
						var error = asyncResult.error;
						if (asyncResult.status === Office.AsyncResultStatus.Failed) {
							write(error.name + ": " + error.message);
						} 
						else {
							if(asyncResult.value["slides"][0]["id"] == savedSlideId){
								// we're on the correct slide
								createVideo();
							}
							else{
								// we're on a different slide
								setTimeout(checkActiveSlide, 600);
							}
						}            
					});
				}

				checkActiveSlide();

			}
			else{
				createVideo();
			}			
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
			
			function checkServerDatabase(callback, password){
				write("checking for pw" + password);
				var xhttp = new XMLHttpRequest();
				xhttp.onreadystatechange = function() {
					if (this.readyState == 4) {
						callback(this.status);
					}
				};
				var phpUrl = "https://michael-saunders.com/server/";
				var paramsString = "custom=" + userId;
				if(password){
					// then it's a WVP-authenticated check
					phpUrl += "checkdatabasepw.php";
					paramsString += "&pass=" + password;
				}
				else{
					phpUrl += "checkdatabase.php";
					write("correct entrypoint");
				}
				xhttp.open("POST", phpUrl, true);
				xhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
				xhttp.send(paramsString);
			}
			
			if(userId){
				$("#signinlink").hide();
				checkServerDatabase(function(myStatus){
					if(myStatus == 200){
						write("result succeeded");
						turnOnPro();
					}
					else{
						// try again once after waiting
						setTimeout(function(){
							checkServerDatabase(function(myStatus2){
								if(myStatus2 == 200){
									write("result succeeded");
									turnOnPro();
								}
								else{
									// now we give up
									if(Office.context.commerceAllowed){
										showAd();
									}
									else{
										// hide additional features options for iPad
										if(!Office.context.commerceAllowed){
											$('#premiumFeatures').hide();
											$('#helpLink').attr("href", "../pages/helpnocommerce.html");
										}
										else{
											// there's no user ID
											// this should no longer happen
											document.getElementById('premiumFeatures').title += '. Sign in before purchase.';
											$('#premiumFeatures').hide();
										}
										
									}
									
									write("result status: " + myStatus);
								}								
							});
						},2000);
					}
				});
			}
			else if(localStorage.getItem("email")){
				turnOnPro();
				userId = email;
				signedIn = true;
				$("#signinlink").hide();
				ga("send","event","videoplayer","autosignin",localStorage.getItem("email"));
			}
			else if(Office.context.commerceAllowed){
				$("#price").html("$9.95");
				write("adjust price");
				showAd();
				// then they need to log in first
				$.getScript("https://michael-saunders.com/videoplayer/scripts/sha256.js",function(response,status){});

				$('#proContent').hide();
				$('#buyNow').hide();
				// set UI for registration				
				
				setTimeout(function(){
					$('#signInText').html('Register and contribute <b><span id="price">$9.95</span></b> to unlock the <a href="javascript:;" class="linkstyle" id="registerHelp" style="color:#bbe2fa">premium features</a> for all your videos.');
					$('#registerHelp').click(function(){
						window.open("https://www.michael-saunders.com/videoplayer/pages/help.html");
					});
					$('#registerSignInText').html("Register");

					$('.registering').fadeIn(200);
					
					

				},250);
				document.getElementById('premiumFeatures').title += '. Sign in to Office before purchase.';
			}
			else{
				$('#premiumFeatures').hide();
				$('#helpLink').attr("href", "../pages/helpnocommerce.html");
			}
			

            document.getElementById("cloak").style.visibility = 'hidden';
            
			$("#videoID").focus();
			
			
        }
    });
    
};