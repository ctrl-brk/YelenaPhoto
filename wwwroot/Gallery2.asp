<%Option Explicit%>
<%
OpenDBConnect()

function GetHTML( obRec, sStyle, nImageId )

   GetHTML = ""

   GetHTML = "<div class='gst05'" & sStyle & "><div class='gst051'><img src='" & obRec( "RelativePath" ) & "/thmb/" & obRec( "ImageName" ) & "' id='img_" & obRec( "ImageId" ) & "' onclick='LoadImage(this)' onmouseover='OnImgOver(this,true)' onmouseout='OnImgOver(this,false)'"

   if ( sFirstFile = "" or nImageId = CInt( obRec( "ImageId" ) ) ) then
      sFirstFile = "src='" & obRec( "RelativePath" ) & "/" & obRec( "ImageName" ) & "' width='" & obRec( "Width" ) & "' height='" & obRec( "Height" ) & "'"
      Response.Write( "CurImgId='img_" & obRec( "ImageId" ) & "';" & vbCrLf )
      GetHTML = GetHTML & " OnLoad='HighliteImg(this)'"
   end if

   GetHTML = GetHTML & "></a></div></div>" & vbCrLf

end function
%>
<script language='JavaScript'>

var sPath;
var bMainImgLoaded = false, bNewImageLoaded = false;
var obCurThumb = null;
var CurImgId = 0;
var aImgs = new Array();
var bSS = false;
var aTransitions = new Array();
var imgProgress = null;
var mozOpacity;
var mozImg;
var mozNextImg;

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

function BarnTransition()
{
   this.motions = new Array( "in", "out" );
   this.orientations = new Array( "vertical", "horizontal" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.Barn(motion='" + this.motions[RandRange( 0, 1 )] + "',orientation='" + this.orientations[RandRange(0,1)] + "')";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

function BlindsTransition()
{
   this.directions = new Array( "up", "down", "right", "left" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.Blinds(bands=14,direction='" + this.directions[RandRange( 0, 3 )] + "')";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

function CheckerBoardTransition()
{
   this.directions = new Array( "up", "down", "right", "left" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.CheckerBoard(squaresx=14,squaresy=14,direction='" + this.directions[RandRange( 0, 3 )] + "')";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
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
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

function GradientWipeTransition()
{
   this.motions = new Array( "forward", "reverse" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.gradientWipe(gradientsize=0.5,motion='" + this.motions[RandRange(0,1)] + "',WipeStyle=" + RandRange(0,1) + ")";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

function InsetTransition()
{
   this.motions = new Array( "forward", "reverse" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.Inset();";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

function IrisTransition()
{
   this.motions = new Array( "in", "out" );
   this.styles = new Array( "DIAMOND","CIRCLE","CROSS","PLUS","SQUARE","STAR" );
   this.imgSrc = null;
   this.imgDst = null;
   this.Transit = function()
     {
        this.Filter = "progid:DXImageTransform.Microsoft.Iris(motion='" + this.motions[RandRange(0,1)] + "',irisStyle='" + this.styles[RandRange(0,5)] + "')";

        this.imgSrc.style.filter = this.Filter;
        this.imgSrc.style.visibility = "hidden";
        this.imgSrc.filters[0].Apply();
        this.imgSrc.src = this.imgDst.src;
        this.imgSrc.width = this.imgDst.width;
        this.imgSrc.height = this.imgDst.height;
        this.imgSrc.style.visibility = "visible";
        this.imgSrc.filters[0].Play(duration=0.5);
     }

}

aTransitions[aTransitions.length] = new EmptyTransition();
aTransitions[aTransitions.length] = new MozTransition();
aTransitions[aTransitions.length] = new BarnTransition();
aTransitions[aTransitions.length] = new BlindsTransition();
aTransitions[aTransitions.length] = new CheckerBoardTransition();
aTransitions[aTransitions.length] = new FadeTransition();
aTransitions[aTransitions.length] = new GradientWipeTransition();
aTransitions[aTransitions.length] = new InsetTransition();
aTransitions[aTransitions.length] = new IrisTransition();

var curTrans;
var loadAttempts;

function StartTransition()
{
   if ( bNewImageLoaded )
      {
      imgProgress.style.display = "none";
      curTrans.Transit();
      }
   else
      {
      loadAttempts++;
      if ( loadAttempts > 9 ) imgProgress.style.display = "";
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
   

   if ( IsIE() ) nTransId = RandRange( 2, aTransitions.length - 1 );
   nTransId=1;
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

function OnMainImageLoad( img )
{
   bMainImgLoaded = true;
}

function HighliteImg( img )
{
   if ( IsIE() )
      {
      obCurThumb = img;
      img.filters(0).enabled = false;
      }
}

function LoadImage( img, bAuto )
{
   if ( IsIE() )
      {
      if ( obCurThumb ) obCurThumb.filters(0).enabled = true;
      obCurThumb = img;
      }
   if ( !bAuto ) { bSS = false; SwitchSSButtons(); }
   CurImgId = img.id;
   bMainImgLoaded = false;
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
   LoadImage( img, true );
   window.setTimeout( "NextImage()", 5000 );
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

   SwitchSSButtons();
   window.setTimeout( "NextImage()", 1000 );
}

function StopSS()
{
   bSS = false;
   SwitchSSButtons();
}

<%
dim nId, nImageId, Gallery, obRec, i, sStyle, sHTML, sFirstFile, nHeight

nId = CInt( RequestString( "id" ) )
gsSubmenu = ""
gbGallery = true
nHeight = 0

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
            <div class='gst03'><img id='mainimg1' <%=sFirstFile%> align='center' onload='OnMainImageLoad(this)'></div>
         </td>
         <td class='gst02'><div class='gst04'<% if ( nHeight > 730 ) then %> style='overflow-y:scroll'<% end if %>><%=sHTML%></div></td>
      </tr>
   </table>
<% else %>
   <p style='text-align:left; margin-bottom:100px'>This gallery is under construction. It will be available very soon. Please come back later.</p>
<% end if %>


<% set Gallery = nothing %>

<!--#include virtual="/Framework/Close.asp"-->
