function getParameterByName(n,t){n=n.replace(/[\[\]]/g,"\\$&");var r=new RegExp("[?&]"+n+"(=([^&#]*)|&|#|$)"),i=r.exec(t);return i?i[2]?decodeURIComponent(i[2].replace(/\+/g," ")):"":null}function onYouTubeIframeAPIReady(){player=new YT.Player("player",{height:"342px",width:"608px",videoId:vid,playerVars:{autoplay:vautoplay,start:vstarttime,end:vendtime},events:{onReady:onPlayerReady,onStateChange:onPlayerStateChange}});document.getElementById("cloak").style.visibility="hidden"}function onPlayerReady(){}function onPlayerStateChange(){}function stopVideo(){player.stopVideo()}function fillScreen(){ytFrame=document.getElementById("player");ytFrame.style.height=$(window).height();ytFrame.style.width=$(window).width();document.getElementById("cloak").style.visibility="hidden"}var tag=document.createElement("script"),firstScriptTag,done;tag.src="https://www.youtube.com/iframe_api";firstScriptTag=document.getElementsByTagName("script")[0];firstScriptTag.parentNode.insertBefore(tag,firstScriptTag);var player,vurl=Office.context.document.settings.get("vid"),vautoplay=Office.context.document.settings.get("autoplay"),vstarttime=Office.context.document.settings.get("starttime"),vendtime=Office.context.document.settings.get("endtime"),vindex,vid;vindex=vurl.indexOf("watch?v=")!=-1?vurl.indexOf("watch?v="):vurl.indexOf("m/embed/")!=-1?vurl.indexOf("m/embed/"):vurl.indexOf("e.com/v/")!=-1?vurl.indexOf("e.com/v/"):vurl.indexOf("outu.be/");vid=vurl.substring(vindex+8);getParameterByName("v",vurl)&&(vid=getParameterByName("v",vurl));done=!1