

// 2. This code loads the IFrame Player API code asynchronously.
var tag = document.createElement('script');

tag.src = "https://www.youtube.com/iframe_api";
var firstScriptTag = document.getElementsByTagName('script')[0];
firstScriptTag.parentNode.insertBefore(tag, firstScriptTag);

// 3. This function creates an <iframe> (and YouTube player)
//    after the API code downloads.
var player;
var vurl = Office.context.document.settings.get("vid");
var vautoplay = Office.context.document.settings.get("autoplay");
var vstarttime = Office.context.document.settings.get("starttime");
var vendtime = Office.context.document.settings.get("endtime");
var vindex;
var vid; // video ID code

function getParameterByName(name, url) {
    // from Stack Overflow 
    // if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

if(vurl.indexOf("watch?v=") != -1){
  // example: https://www.youtube.com/watch?v=dQw4w9WgXcQ
  vindex = vurl.indexOf("watch?v=");
}
else if(vurl.indexOf("m/embed/") != -1){
  // example: https://www.youtube.com/embed/dQw4w9WgXcQ
  vindex = vurl.indexOf("m/embed/");
}
else if(vurl.indexOf("e.com/v/") != -1){
  // example: https://www.youtube.com/v/dQw4w9WgXcQ
  vindex = vurl.indexOf("e.com/v/");
}
else{
  // example: https://youtu.be/dQw4w9WgXcQ
  vindex = vurl.indexOf("outu.be/"); 
}

vid = vurl.substring(vindex+8); // the stuff after the found substring

// now the proper way to get the vid ID
if(getParameterByName("v",vurl)){
  // example: https://www.youtube.com/watch?time_continue=5&v=dQw4w9WgXcQ or https://www.youtube.com/watch?v=dQw4w9WgXcQ&list=dQw4w9WgXcQ
  vid = getParameterByName("v",vurl);
}

function onYouTubeIframeAPIReady() {
  player = new YT.Player('player', {
    height: '342px',
    width: '608px',
    videoId: vid,
	  playerVars: {
      'autoplay': vautoplay,
      'start': vstarttime,
      'end': vendtime
	  },
    events: {
      'onReady': onPlayerReady,
      'onStateChange': onPlayerStateChange
    }
  });
  document.getElementById("cloak").style.visibility = 'hidden';
}
/*

var iframe = document.createElement('iframe');
iframe.type="text/html";
iframe.frameborder="0";
iframe.src = "https://www.youtube.com/embed/M7lc1UVf-VE&html5=1";
document.getElementById('player').appendChild(iframe);
*/

// 4. The API will call this function when the video player is ready.
function onPlayerReady(event) {

  
    
  
}

// 5. The API calls this function when the player's state changes.
//    The function indicates that when playing a video (state=1),
//    the player should play for six seconds and then stop.
var done = false;
function onPlayerStateChange(event) {
 /* if (event.data == YT.PlayerState.PLAYING && !done) {
    setTimeout(stopVideo, 6000);
    done = true;
  }*/
}
function stopVideo() {
  player.stopVideo();
}
function fillScreen(){
    ytFrame = document.getElementById("player");
    ytFrame.style.height = $(window).height();
    ytFrame.style.width = $(window).width();
    document.getElementById("cloak").style.visibility = 'hidden';
}
/*
  <!--  <img id="play" src="../content/next.png" style="position:absolute;right:10px;top:45%; cursor:pointer">
    <div id="bar"></div> 
    
<div id="debug" style="position:fixed;height:45%;width:45%;bottom:0px;right:0px;opacity:0.5;background-color:black;color:white">Start</div>
   <!-- <div id="borderdiv"></div>-->
   */