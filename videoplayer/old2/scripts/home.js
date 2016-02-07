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

function createVideo(){
    if(document.getElementById("videoID").value != "")
    {
        write("there's a video to create");
        Office.context.document.settings.set("vid",document.getElementById("videoID").value);
        Office.context.document.settings.saveAsync(function (asyncResult) {
          write('Settings saved with status: ' + asyncResult.status);
          if(window.top==window){
               // not in iFrame
               document.getElementById('player').style.visibility = 'visible';
               var script = document.createElement( 'script' );
               script.type = 'text/javascript';
               script.src = "../scripts/youtube.js";
               $('body').append( script );
               $('#cloak').fadeOut();
               $('#cloak').remove();
               document.getElementById("clock").style.visibility = 'hidden';
          }
          else{
              
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
          }
        }); 
    }
    
}

function saveVid(){
     
}

Office.initialize = function (reason) {
    $(document).ready(function(){
        if(window.top==window){
            //not in iframe
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
        
        if(Office.context.document.settings.get("vid")){
            document.getElementById("videoID").value = Office.context.document.settings.get("vid");
            createVideo();
        }
        else{
            $('#cloak').fadeOut();
            $('#cloak').remove();
            document.getElementById("clock").style.visibility = 'hidden';
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