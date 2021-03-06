<%Option Explicit%>
<%
OpenDBConnect()

function GetHTML( obRec, sStyle, nImageId )

   GetHTML = ""

   GetHTML = "<div class='gst05'" & sStyle & "><div class='gst051' id='fimg_" & obRec( "ImageId" ) & "'><img src='" & obRec( "RelativePath" ) & "/thmb/" & obRec( "ImageName" ) & "' id='img_" & obRec( "ImageId" ) & "' onclick='LoadImage(this)' onmouseover='OnImgOver(this,true)' onmouseout='OnImgOver(this,false)'"

   if ( sFirstFile = "" or nImageId = CInt( obRec( "ImageId" ) ) ) then
      Response.Write( "CurImgId='img_" & obRec( "ImageId" ) & "';" & vbCrLf )
      GetHTML = GetHTML & " OnLoad='HighliteImg(this," & nImageId 
      if ( sFirstFile = "" ) then GetHTML = GetHTML & ",true"
      GetHTML = GetHTML & ")'"
      sFirstFile = "src='" & obRec( "RelativePath" ) & "/" & obRec( "ImageName" ) & "' width='" & obRec( "Width" ) & "' height='" & obRec( "Height" ) & "'"
   end if

   GetHTML = GetHTML & "></a></div></div>" & vbCrLf

end function
%>
<script language='JavaScript'>

var sPath;
var bNewImageLoaded = false;
var obCurThumb = null;
var CurImgId = 0;
var aImgs = new Array();
var bSS = false;
var aTransitions = new Array();
var imgProgress = null;
var mozOpacity;
var mozImg;
var mozNextImg;
var bScroll = true;

function EmptyTransition()
{
   this.imgSrc = null;
   this.imgDst = null;

   this.Transit = function()
      {
      this.imgSrc.src = this.imgDst.src;
      this.imgSrc.width = this.imgDst.width;
      this.imgSrc.height = this.imgDst.height;
      }
}

function MozFadeout()
{
   if ( mozOpacity > 0 )
      {
      mozImg.style.opacity = mozOpacity;
      mozOpacity -= 0.1;
      setTimeout( "MozFadeout()", 20 );
      }
   else
      {
      mozOpacity = 0.04;
      MozFadein();
      }
}

function MozFadein()
{
   if ( mozOpacity == 0.04 )
      {
      mozImg.src = mozNextImg.src;
      mozImg.width = mozNextImg.width;
      mozImg.height = mozNextImg.height;
      }
   if ( mozOpacity < 0.99 )
      {
      mozImg.style.opacity = mozOpacity;
      mozOpacity += 0.1;
      setTimeout( "MozFadein()", 20 );
      }
}

function MozTransition()
{
   this.imgSrc = null;
   this.imgDst = null;

   this.Transit = function()
      {
      mozImg = this.imgSrc;
      mozNextImg = this.imgDst;
      mozOpacity = 0.99;
      MozFadeout();
      }
}

function FadeTransition()
{
   this.centers = new Array( "true", "false" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.Fade(center=" + RandRange(0,1) + ")";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.filters[0].Play(duration=1);
     }

}

aTransitions[aTransitions.length] = new EmptyTransition();
aTransitions[aTransitions.length] = new MozTransition();
aTransitions[aTransitions.length] = new FadeTransition();

var curTrans;
var loadAttempts;

function StartTransition()
{
   if ( bNewImageLoaded )
      {
      imgProgress.style.display = "none";
      curTrans.Transit();
      if ( bSS ) window.setTimeout( "NextImage()", 5000 );
      }
   else
      {
      loadAttempts++;
      if ( loadAttempts > 50 ) imgProgress.style.display = "";
      setTimeout( "StartTransition()", 20 );
      }
}

function TransitImages( img )
{
   var newImg = new Image();
   var curImg = document.getElementById( "mainimg1" );
   var galImg = FindImage( img.id );
   var nTransId = 0;

   if ( imgProgress == null ) imgProgress = document.getElementById( "imgWait" );

   newImg.width = galImg.Width;
   newImg.height = galImg.Height;
   newImg.onload = function() { bNewImageLoaded = true; };
   bNewImageLoaded = false;
   loadAttempts = 0;
   newImg.src = StrRep( img.src, "/thmb", "" );
   

   if ( IsIE() ) nTransId = 2;
   else if ( IsNN() || IsOpera() ) nTransId = 1;

   curTrans = aTransitions[nTransId];
   curTrans.imgSrc = curImg;
   curTrans.imgDst = newImg;
   StartTransition();
}


function GalleryImage( nId, img, nWidth, nHeight )
{
   this.ImageId = "img_" + nId;
   this.obImage = img;
   this.Width = nWidth;
   this.Height = nHeight;
}

function HighliteBorder( img, bHighlite )
{
   document.getElementById( "f" + img.id ).style.borderColor = bHighlite ? "C7B299" : "322D2A";
}

function HighliteImg( img, id, bFirstImage )
{
   if ( id != null )
      {
      if ( !bFirstImage ) img.scrollIntoView( true );
      HighliteBorder( img, true );
      }
   obCurThumb = img;
   if ( IsIE() ) img.filters(0).enabled = false;
}

function LoadImage( img, bAuto )
{
   if ( obCurThumb )
      {
      HighliteBorder( obCurThumb, false );
      if ( IsIE() ) obCurThumb.filters(0).enabled = true;
      obCurThumb = img;
      }
   if ( !bAuto ) { bSS = false; SwitchSSButtons(); }
   CurImgId = img.id;
   HighliteBorder( img, true );
   TransitImages( img );
}

function OnImgOver( img, bOver )
{
   if ( !IsIE() || ( obCurThumb && obCurThumb == img ) ) return;
   img.filters(0).enabled = !bOver;
}

function FindImage( sId )
{
   for ( var i=0; i<aImgs.length; i++ )
      {
      if ( aImgs[i].ImageId == sId ) break;
      }

   if ( i < aImgs.length ) return( aImgs[i] );
   else return( null );
}

function NextImage()
{
   var i;

   if ( !bSS ) return;

   for ( i=0; i<aImgs.length; i++ )
      {
      if ( aImgs[i].ImageId == CurImgId ) break;
      }
   if ( i >= aImgs.length ) return;
   if ( i == aImgs.length - 1 ) i = -1;

   var img = document.getElementById( aImgs[i+1].ImageId );
   OnImgOver( img, true );
   if ( bScroll ) img.scrollIntoView( true );
   LoadImage( img, true );
}

function SwitchSSButtons()
{
   var bp = document.getElementById( "btnPlay" );
   var bs = document.getElementById( "btnStop" );

   bp.src = "/Images/btn/play" + bSS.toString() + ".gif";
   bp.style.cursor = bSS ? "default" : "pointer";

   bs.src = "/Images/btn/stop" + (!bSS).toString() + ".gif";
   bs.style.cursor = bSS ? "pointer" : "default";
}

function PlaySS()
{
   if ( bSS || aImgs.length < 2 ) return;

   bSS = true;
   bScroll = true;

   SwitchSSButtons();
   window.setTimeout( "NextImage()", 1000 );
}

function StopSS()
{
   bSS = false;
   bScroll = false;
   SwitchSSButtons();
}

function OnStripScroll()
{
   bScroll = false;
}

function OnLoad()
{
window.scrollTo( 0, window.document.documentElement.scrollHeight );
}

<%
dim nId, nImageId, Gallery, obRec, i, sStyle, sHTML, sFirstFile, nHeight, obBr

nId = CInt( RequestString( "id" ) )
gsSubmenu = ""
gbGallery = true
nHeight = 0

set obBr = new Browser

nImageId = 0
if ( RequestString( "iid" ) <> "" ) then nImageId = CInt( RequestString( "iid" ) )

set Gallery = new CGallery
Gallery.Init( nId )

set obRec = ExecuteBlob( DebugSQL( "exec GetGalleries" ) )
ClearDebug()

while ( not obRec.EOF )
   if ( CInt( obRec( "GalleryId" ) ) <> nId ) then gsSubmenu = gsSubmenu & "<a href='/Gallery.asp?id=" & obRec( "GalleryId" ) & "'>&raquo; " else gsSubmenu = gsSubmenu & "<span>&nbsp; " end if
   gsSubmenu = gsSubmenu & obRec( "GalleryName" )
   if ( CInt( obRec( "GalleryId" ) ) <> nId ) then gsSubmenu = gsSubmenu & "</a>" else gsSubmenu = gsSubmenu & "</span>" end if
   obRec.MoveNext()
wend

obRec.Close() : set obRec = nothing

set obRec = ExecuteBlob( DebugSQL( "exec GetImages @GalleryId=" & nId ) )
ClearDebug()

i = 1 : sStyle = "" : sFirstFile = ""

if ( not obRec.EOF ) then Response.Write( "sPath='" & obRec( "RelativePath" ) & "';" & vbCrLf )
while ( not obRec.EOF )

   if ( i >= CInt( obRec( "ImageCount" ) ) ) then sStyle =  " style='margin-bottom:0'"
   sHTML = sHTML & GetHTML( obRec, sStyle, nImageId )
   nHeight = nHeight + CInt( obRec( "ThumbHeight" ) ) + 31 
   i = i+1

   Response.Write( "aImgs[aImgs.length]=new GalleryImage(" & obRec( "ImageId" ) & ",'" & obRec( "ImageName" ) & "'," & obRec( "Width" ) & "," & obRec( "Height" ) & ");" & vbCrLf )

   obRec.MoveNext()

wend

obRec.Close() : set obRec = nothing
%>
</script>

<!--#include virtual="/Framework/Open.asp"-->
<img src='/Images/progress.gif' id='imgWait' style='position:absolute;top:170px;display:none;'>
<table border='0' cellspacing='0' cellpadding='0' class='tblslideshow'>
   <tr>
      <td><h1><%=Gallery.Name%> <strong>gallery</strong></h1></td>
      <% if ( i > 1 ) then %>
         <td class='sst01'>
            <img src='/Images/btn/stoptrue.gif' id='btnStop' border='0' align='absmiddle' style='margin-left:20px' onclick='StopSS()' alt='Stop slide show'>
            <img src='/Images/btn/playfalse.gif' id='btnPlay' border='0' align='absmiddle' onclick='PlaySS()' alt='Play slide show' style='cursor:pointer'>
         </td>
      <% end if %>
   </tr>
</table>

<% if ( i > 1 ) then %>
   <table border='0' class='gtblcnt' cellspacing='0' cellpadding='0'>
      <tr>
         <td class='gst01'>
            <div class='gst03'><img id='mainimg1' <%=sFirstFile%> align='center'></div>
         </td>
         <td class='gst02'><div id='Thumbnails' class='gst04'<% if ( nHeight > 730 ) then %> style='overflow<% if ( obBr.bOpera ) then %>:auto<% else %>-y:scroll<% end if %>'<% end if %> onScroll='OnStripScroll()'><%=sHTML%></div></td>
      </tr>
   </table>
<% else %>
   <p style='text-align:left; margin-bottom:100px'>This gallery is under construction. It will be available very soon. Please come back later.</p>
<% end if %>


<% set Gallery = nothing : set obBr = nothing %>

<img src='/Images/btn/playtrue.gif' style='display:none'>
<img src='/Images/btn/stopfalse.gif' style='display:none'>

<!--#include virtual="/Framework/Close.asp"-->
