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

function errorMessage(myText){
     document.getElementById("innerErrorDiv").innerHTML = myText;
     var theErrorDiv = document.getElementById("errorDiv");
     //theErrorDiv.style.visibility = "visible";
     $('#errorDiv').fadeIn();
}

function createVideo(){
    var urlString = document.getElementById("videoID").value;
    if(urlString.indexOf("youtube.com") != -1 || urlString.indexOf("youtu.be") != -1)
    {
          write("there's a video to create");
          Office.context.document.settings.set("vid",urlString);
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
            if(document.getElementById("links").innerHTML.indexOf("Rate") == -1){
               document.getElementById("links").innerHTML += "<a href='https://store.office.com/writereview.aspx?assetid=WA104221182'>Rate</a>";
            }
            
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
        
        if(Office.context.document.settings.get("vid")){
            //first hide the ad
			document.getElementById("goog").style.display = "none";
			
			document.getElementById("videoID").value = Office.context.document.settings.get("vid");
            createVideo();
			
        }
        else{
            
			$('#cloak').fadeOut();
            $('#cloak').remove();
            document.getElementById("cloak").style.visibility = 'hidden';
            
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