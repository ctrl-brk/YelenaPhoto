<%Option Explicit%>
<% gbAdmin = true %>
<!--#include virtual="/Framework/Open.asp"-->
<% if ( not gcUser.bAdmin ) then LoginRedirect( hRef( "/Admin/Default.asp" ) ) %>

<script language='JavaScript'>

var gnGalleryId = 0, gbUpload = false, gobCurImg = null;

function LoadPanel( sId, sReq )
{
   var s = AjaxRequest( "ajax/" + sReq );
   document.getElementById( sId ).innerHTML = s;
}

function ClearPanel( sId )
{
   document.getElementById( sId ).innerHTML = "&nbsp;";
}

function ClearImagePanels()
{
   ClearPanel( "GalImgs" );
   ClearPanel( "ImgsAction" );
}

function UploadFile( sSaveBtn, nImgId )
{
   //IE->Tools->Security->(whatever zone)->Settings->
   //   Initialize and script ActiveX controls not marked as safe for scripting - must be "Enable" (not prompt)
   //   Access data across domains - enable too.


   var btn = GetControl( sSaveBtn );
   var f = GetControlValue( "ImageFile" );
   var btnClr = btn.style.color;
   var btnVal = btn.value;
   btn.value = "Please wait...";
   btn.style.color = "red";

   gbUpload = true;
   
   // create ADO-stream Object
   var Stream = new ActiveXObject( "ADODB.Stream" );

   // create XML document with default header and primary node
   var DOM = new ActiveXObject( "MSXML2.DOMDocument" );
   DOM.loadXML( '<?xml version="1.0" ?> <root/>' );
   // specify namespaces datatypes
   DOM.documentElement.setAttribute( "xmlns:dt", "urn:schemas-microsoft-com:datatypes" );

   // create a new node and set binary content
   if ( f != "" )
      {
      var ImgNode = DOM.createElement( "Image1" );
      ImgNode.dataType = "bin.base64";
      // open stream object and read source file
      Stream.Type = 1;  // 1=adTypeBinary 
      Stream.Open(); 
      Stream.LoadFromFile( f );
      // store file content into XML node
      ImgNode.nodeTypedValue = Stream.Read( -1 ); // -1=adReadAll
      Stream.Close();
      DOM.documentElement.appendChild( ImgNode );
      }

   // we can create more XML nodes for multiple file upload
   var ImgIdNode = DOM.createElement( "ImageId" );
   ImgIdNode.dataType = "string";
   if ( nImgId != null ) ImgIdNode.nodeTypedValue = nImgId;
   DOM.documentElement.appendChild( ImgIdNode );
   
   var ExistsNode = DOM.createElement( "ImageExists" );
   ExistsNode.dataType = "string";
   if ( f != "" ) ExistsNode.nodeTypedValue = "Y";
   DOM.documentElement.appendChild( ExistsNode );

   var FeaturedNode = DOM.createElement( "FeaturedFlag" );
   FeaturedNode.dataType = "string";
   FeaturedNode.nodeTypedValue = GetControlValue( "ImageFeatured" ) == "on" ? "Y" : "N";
   DOM.documentElement.appendChild( FeaturedNode );

   var DescNode = DOM.createElement( "Description" );
   DescNode.dataType = "string";
   DescNode.nodeTypedValue = GetControlValue( "ImageDesc" );
   DOM.documentElement.appendChild( DescNode );

   var SortNode = DOM.createElement( "SortValue" );
   SortNode.dataType = "string";
   SortNode.nodeTypedValue = GetControlValue( "ImageSortValue" );
   DOM.documentElement.appendChild( SortNode );

   var ActiveNode = DOM.createElement( "ActiveFlag" );
   ActiveNode.dataType = "string";
   ActiveNode.nodeTypedValue = GetControlValue( "ImageEnabled" ) == "on" ? "Y" : "N";
   DOM.documentElement.appendChild( ActiveNode );

   // send XML documento to Web server
   var s = AjaxRequest( "ajax/FileReceive.asp?gid=" + gnGalleryId, DOM );

   gbUpload = false;
   btn.value = btnVal;
   btn.style.color = btnClr;

   // show server message in message-area
   return( s );
}

function OnGalleryChange()
{
   var sel = GetControl( "Galleries" );
   LoadGalleryInfo( sel.options[sel.selectedIndex] ); 
   ClearPanel( "ImgsAction" );
   gobCurImg = null;
}

function LoadGalleryInfo( opt )
{
   gnGalleryId = opt.value;
   LoadPanel( "GalInfo", "GalleryInfo.asp?id=" + gnGalleryId );
   LoadPanel( "GalImgs", "GalleryImages.asp?id=" + gnGalleryId );
}

function OnAddImage()
{
   if ( gnGalleryId == 0 ) { alert( "Select gallery first" ); return; }

   LoadPanel( "ImgsAction", "AddImageForm.asp?id=" + gnGalleryId );
}

function OnAddImageSave()
{
   if ( gbUpload ) { alert( "File is uploading. Please wait" ); return; }

   var f = GetControl( "ImageFile" ).value;
   if ( f == "" ) { alert( "Select file first" ); return; }
   if ( f.toLowerCase().indexOf( ".jpg" ) < 1 ) { alert( "Only jpg files, please" ); return; }

   var s = UploadFile( "btnImgAddSave" );

   if ( s == null || s == "" ) { alert( "File upload error" ); return; }
   if ( !isValidDigit( s ) ) { alert( s ); return; }

   LoadPanel( "GalImgs", "GalleryImages.asp?id=" + gnGalleryId + "&iid=" + s );
}

function OnImageSelect( nId, nGalId )
{
   if ( !nGalId )
      {
      if ( gobCurImg != null ) document.getElementById( "td" + gobCurImg.id ).style.borderColor = "#352D2B";
      gobCurImg = document.getElementById( "img_" + nId );
      document.getElementById( "td" + gobCurImg.id ).style.borderColor = "#EBD3AA";
      }
   else
      {
      gnGalleryId = nGalId;
      LoadPanel( "GalList", "Galleries.asp?id=" + gnGalleryId );
      LoadPanel( "GalInfo", "GalleryInfo.asp?id=" + gnGalleryId );
      LoadPanel( "GalImgs", "GalleryImages.asp?id=" + gnGalleryId + "&iid=" + nId + "&sel=Y" );
      }

   LoadPanel( "ImgsAction", "EditImageForm.asp?id=" + nId );
   gobCurImg = document.getElementById( "img_" + nId );
}

function OnImageSave( nId )
{
   if ( gbUpload ) { alert( "File is uploading. Please wait." ); return; }

   var f = GetControlValue( "ImageFile" );
   if ( f != "" && f.toLowerCase().indexOf( ".jpg" ) < 1 ) { alert( "Only jpg files, please" ); return; }
   var s = UploadFile( "btnImgSave", nId );

   if ( s == null || s == "" ) { alert( "File upload error" ); return; }
   if ( !isValidDigit( s ) ) { alert( s ); return; }

   LoadPanel( "GalImgs", "GalleryImages.asp?id=" + gnGalleryId + "&iid=" + s );
   gobCurImg = null;
   OnImageSelect( s );
   alert( "Image updated" );
}

function OnFeaturedImages()
{
   ClearPanel( "ImgsAction" );
   LoadPanel( "GalImgs", "FeaturedImages.asp" );
}

function OnImageDelete( nId )
{
   if ( gbUpload ) { alert( "File is uploading. Please wait." ); return; }

   if ( !confirm( "Are you sure?" ) ) return;

   gobCurImg = null;
   LoadPanel( "ImgsAction", "DeleteImage.asp?id=" + nId );
   LoadPanel( "GalImgs", "GalleryImages.asp?id=" + gnGalleryId );
}

function OnAddGallery()
{
   ClearImagePanels();
   GetControl( "Galleries" ).selectedIndex = -1;
   LoadPanel( "GalInfo", "AddGalleryForm.asp" );   
   gnGalleryId = 0;
}

function OnAddSubGallery()
{
   var sel = GetControlValue( "Galleries" );

   if ( sel != null ) LoadPanel( "GalInfo", "AddGalleryForm.asp?id=" + sel );   
   else alert( "Please select parent gallery first" );
}

function OnGallerySave( nId, nParentId )
{
   var Stream = new ActiveXObject( "ADODB.Stream" );
   var DOM = new ActiveXObject( "MSXML2.DOMDocument" );
   var f = GetControlValue( "GalleryImage" );

   DOM.loadXML( '<?xml version="1.0" ?> <root/>' );
   DOM.documentElement.setAttribute( "xmlns:dt", "urn:schemas-microsoft-com:datatypes" );

   if ( f != "" )
      {
      var ImgNode = DOM.createElement( "Image" );
      ImgNode.dataType = "bin.base64";
      Stream.Type = 1;  // 1=adTypeBinary 
      Stream.Open(); 
      Stream.LoadFromFile( f );

      ImgNode.nodeTypedValue = Stream.Read( -1 ); // -1=adReadAll
      Stream.Close();
      DOM.documentElement.appendChild( ImgNode );

      var ResizeNode = DOM.createElement( "ResizeImage" );
      ResizeNode.dataType = "string";
      ResizeNode.nodeTypedValue = GetControlValue( "GalleryImageResize" ) == "on" ? "Y" : "N";
      DOM.documentElement.appendChild( ResizeNode );
      }

   var IdNode = DOM.createElement( "GalleryId" );
   IdNode.dataType = "string";
   if ( nId != null ) IdNode.nodeTypedValue = nId;
   DOM.documentElement.appendChild( IdNode );

   var IdNode = DOM.createElement( "ParentId" );
   IdNode.dataType = "string";
   if ( nParentId != null ) IdNode.nodeTypedValue = nParentId;
   DOM.documentElement.appendChild( IdNode );
   
   var ExistsNode = DOM.createElement( "ImageExists" );
   ExistsNode.dataType = "string";
   if ( f != "" ) ExistsNode.nodeTypedValue = "Y";
   DOM.documentElement.appendChild( ExistsNode );

   var NameNode = DOM.createElement( "Name" );
   NameNode.dataType = "string";
   NameNode.nodeTypedValue = GetControlValue( "GalleryName" );
   DOM.documentElement.appendChild( NameNode );

   var DescNode = DOM.createElement( "Description" );
   DescNode.dataType = "string";
   DescNode.nodeTypedValue = GetControlValue( "GalleryDesc" );
   DOM.documentElement.appendChild( DescNode );

   var PathNode = DOM.createElement( "Path" );
   PathNode.dataType = "string";
   PathNode.nodeTypedValue = GetControlValue( "GalleryPath" );
   DOM.documentElement.appendChild( PathNode );

   var SortNode = DOM.createElement( "SortValue" );
   SortNode.dataType = "string";
   SortNode.nodeTypedValue = GetControlValue( "GallerySortValue" );
   DOM.documentElement.appendChild( SortNode );

   var ActiveNode = DOM.createElement( "ActiveFlag" );
   ActiveNode.dataType = "string";
   ActiveNode.nodeTypedValue = GetControlValue( "GalleryEnabled" ) == "on" ? "Y" : "N";
   DOM.documentElement.appendChild( ActiveNode );

   if ( AjaxRequest( "ajax/SaveGallery.asp", DOM ) != null )
      {
      if ( nId != null )
         {
         var sel = GetControl( "Galleries" );
         sel.options[sel.selectedIndex].text = GetControlValue( "GalleryName" );
         alert( "Gallery updated" );
         }
      else
         {
         LoadPanel( "GalList", "Galleries.asp" );
         ClearPanel( "GalInfo" );
         }
      }
}

function OnGalleryDelete( nId )
{
   if ( !confirm( "Are you sure?" ) ) return;

   gobCurImg = null;
   LoadPanel( "GalInfo", "DeleteGallery.asp?id=" + nId );
   LoadPanel( "GalList", "Galleries.asp" );
   ClearPanel( "GalImgs" );
   ClearPanel( "ImgsAction" );
}

function OnLoad()
{
   LoadPanel( "GalList", "Galleries.asp" );
}
</script>

<h1>Site <strong>admin</strong></h1>
<table border='0' width='100%'>
   <tr>
      <td valign='top' width='280' style='border-right:1px solid #62584e;'>
         <b>Galleries</b>&nbsp;&nbsp;&nbsp;<input type='button' value='New' onclick='OnAddGallery()'>&nbsp;&nbsp;<input type='button' value='Sub' onclick='OnAddSubGallery()'><br>
         <div id='GalList'></div>
      </td>
      <td valign='top'><div id='GalInfo'></div></td>
   </tr>
</table>
<hr>
<table border='0' width='100%'>
   <tr>
      <td valign='top' width='280' style='border-right:1px solid #62584e;'><div id='ImgsAction'>&nbsp;</div></td>
      <td valign='top'>
         <b>Pictures</b>&nbsp;&nbsp;&nbsp;<input type='button' value='Add new picture' onclick='OnAddImage()'>&nbsp;&nbsp;&nbsp;<input type='button' value='Featured' onclick='OnFeaturedImages()'>
         <div id='GalImgs'></div>
      </td>
   </tr>
</table>
<hr>

<!--#include virtual="/Framework/Close.asp"-->
