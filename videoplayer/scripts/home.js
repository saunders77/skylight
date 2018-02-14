function testAPIs(){write("hello2");Office.context.document.settings.set("mySetting",2);Office.context.document.settings.saveAsync(function(n){write("Settings saved with status: "+n.status)})}function write(n){document.getElementById("debug").innerHTML=document.getElementById("debug").innerHTML+"\n"+n}function getParameterByName(n,t){t||(t=window.location.href);n=n.replace(/[\[\]]/g,"\\$&");var r=new RegExp("[?&]"+n+"(=([^&#]*)|&|#|$)"),i=r.exec(t);return i?i[2]?decodeURIComponent(i[2].replace(/\+/g," ")):"":null}function loadLicenseInfo(){var n,t;write("loading license info");try{var i=atob(decodeURIComponent(getParameterByName("et"))),r="",f=String.fromCharCode(00);for(n=0;n<i.length;n++)i[n]!=f&&(r+=i[n]);t=$($.parseXML(r)).find("t");orgId=t.attr("oid");liveId=t.attr("cid");userId=orgId;liveId&&(userId=liveId);acquisitionDate=t.attr("ad");ga("set","userId",userId)}catch(u){write("Error licensing: "+u.message);ga("send","event","videoplayer","licensing",u.message)}write("user ID is "+userId)}function errorMessage(n){document.getElementById("innerErrorDiv").innerHTML=n;var t=document.getElementById("errorDiv");$("#errorDiv").fadeIn()}function createVideo(){function e(){if(typeof Storage!="undefined"){if(!localStorage.getItem("lastUseDay")){var n=new Date;localStorage.setItem("lastUseDay",n.toDateString());localStorage.setItem("usedDays",1)}}else write("Sorry, your browser does not support Web Storage...")}var n,t,i,o,h;n=Office.context.document.settings.get("vid");var u=0,r=0,f=0;if(n?(u=Office.context.document.settings.get("autoplay"),r=Office.context.document.settings.get("starttime"),f=Office.context.document.settings.get("endtime")):(n=document.getElementById("videoID").value,t=document.getElementById("timeinput").value.split(":"),i=document.getElementById("endtimeinput").value.split(":"),document.getElementById("autoplay").checked&&(u=1),document.getElementById("customstarttime").checked&&(r=+t[t.length-1]+ +t[t.length-2]*60,t.length>2&&(r+=+t[t.length-3]*3600)),document.getElementById("customendtime").checked&&(f=+i[i.length-1]+ +i[i.length-2]*60,i.length>2&&(f+=+i[i.length-3]*3600)),write("setting video params")),Office.context.document.settings.get("slideId")){o=Office.context.document.settings.get("slideId");function s(){Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{},function(n){var t=n.error;n.status===Office.AsyncResultStatus.Failed?write(t.name+": "+t.message):n.value.slides[0].id==o?setTimeout(s,600):window.location.reload()})}s()}n.indexOf("youtube.com")!=-1||n.indexOf("youtu.be")!=-1?(ga("send","event","videoplayer","setvideo","youtube"),e(),write("there's a video to create"),Office.context.document.settings.set("vid",n),Office.context.document.settings.set("autoplay",u),Office.context.document.settings.set("starttime",r),Office.context.document.settings.set("endtime",f),Office.context.document.settings.saveAsync(function(n){if(write("Settings saved with status: "+n.status),!0){document.getElementById("player").style.visibility="visible";var t=document.createElement("script");t.type="text/javascript";t.src="../scripts/youtube.js";$("body").append(t)}else $("#iframed").fadeIn(),$("#iframed").addClass("inWAC"),document.getElementById("iframed").style.visibility="visible",write("iframedo"),$("#iframed").click(function(){window.open(Office.context.document.settings.get("vid"))})})):n.indexOf("vimeo.com")!=-1?(ga("send","event","videoplayer","setvideo","vimeo"),e(),Office.context.document.settings.set("vid",n),Office.context.document.settings.set("autoplay",u),Office.context.document.settings.set("starttime",r),Office.context.document.settings.set("endtime",f),Office.context.document.settings.saveAsync(function(){var i,f,e;if(1){write("creating vimeox");var f=n.indexOf("meo.com/"),e=n.substring(f+8),t=document.getElementById("ifrm");t.style.height="100%";write("heighta: "+t.height);i="?";r&&(i+="&#t="+r+"s");u&&(i+="&autoplay=1");i+="&title=0&amp;byline=0&amp;portrait=0";t.setAttribute("src","//player.vimeo.com/video/"+e+i);write("heighto: "+t.height);write(t.style.width);write("zindex: "+t.style.zIndex)}else f=n.indexOf("meo.com/"),e=n.substring(f+8),window.location.href="//player.vimeo.com/video/"+e+"?title=0&amp;byline=0&amp;portrait=0"})):n.indexOf("liveleakjhkjkljkl;jkl.com")!=-1?(h=document.createElement("a"),h.setAttribute(),Office.context.document.settings.set("vid",n),Office.context.document.settings.saveAsync(function(){})):errorMessage("Choose a valid URL for your video.")}function saveVid(){}var orgId,liveId,userId,acqusitionDate,isPro=!1,pingingForPayment=!1;Office.initialize=function(){$(document).ready(function(){function i(){checkServerDatabase(function(n){n==200?(write("result succeeded"),ga("send","event","videoplayer","purchasesucceeded"),pingingForPayment=!1,$("#proPrompt").fadeOut(),r()):pingingForPayment&&(setTimeout(i,2e3),ga("send","event","videoplayer","checkping"))})}function r(){$(".startsDisabled").prop("disabled",!1);$("#timeinput").css("user-select","text");$("#endtimeinput").css("user-select","text")}function e(){$("#proPrompt").fadeIn()}var n,t,u;if($("#iframed").fadeOut(),$("#setVid").click(function(){if(document.getElementById("videoID").value=="debug")errorMessage("Click <a href='mailto:webvideoplayer@outlook.com?subject=Support Request for "+userId+"&body=Please enable my account. Thank you!'>here<\/a> to send ID code "+userId+" for support.");else{var n=!0;Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{},function(t){var i,r;n&&(i=t.error,write("selecteddata result is "+t.status),t.status===Office.AsyncResultStatus.Failed?(write(i.name+": "+i.message),createVideo()):(r=t.value.slides[0].id,Office.context.document.settings.set("slideId",r),Office.context.document.settings.saveAsync(function(){createVideo()})))});setTimeout(function(){n=!1;createVideo()},400)}}),$("#errorDiv").click(function(){$(this).fadeOut()}),$("#timeinput").change(function(){$("#customstarttime").attr("checked",!0)}),document.getElementById("endtimeinput").onchange=function(){document.getElementById("customendtime").checked=!0},$(".payButton").click(function(){window.open("../pages/purchasewindow.html?custom="+encodeURIComponent(userId));pingingForPayment=!0;$("#waitingPay").show();setTimeout(i,1e4);ga("send","event","videoplayer","checkping")}),$("#cancelPay").click(function(){$("#waitingPay").hide();pingingForPayment=!1}),$("#contactlink").click(function(){window.open("mailto:webvideoplayer@outlook.com")}),$("#helplink").click(function(){window.open("https://www.michael-saunders.com/videoplayer/pages/info.html#howto")}),$("#privacylink").click(function(){window.open("https://www.michael-saunders.com/videoplayer/pages/privacy.html")}),$("#ratelink").click(function(){window.open("https://store.office.com/writereview.aspx?assetid=WA104221182")}),typeof Storage!="undefined"?localStorage.getItem("lastUseDay")&&(n=new Date,write("my used days="+localStorage.getItem("usedDays")),write("datestring="+n.toDateString()),localStorage.getItem("lastUseDay")!=n.toDateString()&&(localStorage.setItem("lastUseDay",n.toDateString()),t=localStorage.getItem("usedDays"),t++,localStorage.setItem("usedDays",t),t>1&&write("show rating link"))):write("Sorry, your browser does not support Web Storage..."),Office.context.document.settings.get("vid"))if(ga("send","event","videoplayer","loadplayer","existingvideo"),Office.context.document.settings.get("slideId")){u=Office.context.document.settings.get("slideId");function f(){Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,{},function(n){var t=n.error;n.status===Office.AsyncResultStatus.Failed?write(t.name+": "+t.message):n.value.slides[0].id==u?createVideo():setTimeout(f,600)})}f()}else createVideo();else{function o(n){if(n.status=="failed")switch(n.error.code){case 12004:write("Domain is not trusted");break;case 12005:write("HTTPS is required");break;case 12007:write("A dialog is already opened.");break;default:write(n.error.message)}else dialog=n.value,dialog.addEventHandler(Office.EventType.DialogMessageReceived,messageHandler),dialog.addEventHandler(Office.EventType.DialogEventReceived,eventHandler)}ga("send","event","videoplayer","loadplayer","novideo");loadLicenseInfo();function checkServerDatabase(n){var t=new XMLHttpRequest;t.onreadystatechange=function(){this.readyState==4&&n(this.status)};t.open("POST","https://michael-saunders.com/server/checkdatabase.php",!0);t.setRequestHeader("Content-type","application/x-www-form-urlencoded");t.send("custom="+userId)}checkServerDatabase(function(n){n==200?(write("result succeeded"),r()):(userId&&Office.context.commerceAllowed?e():Office.context.commerceAllowed?(document.getElementById("premiumFeatures").title+=". Sign in to Office before purchase.",$("#premiumFeatures").hide()):($("#premiumFeatures").hide(),$("#helpLink").attr("href","../pages/helpnocommerce.html")),write("result status: "+n))});document.getElementById("cloak").style.visibility="hidden";$("#videoID").focus()}})}