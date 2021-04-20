<%Option Explicit%>
<!--#include virtual="/Include/Functions.asp"-->
<!--#include virtual="/Include/DebugFunctions.asp"-->
<!--#include virtual="/Include/OpenDBConnect.asp"-->
<!--#include virtual="/Include/BizObj.asp"-->
<%
Response.ContentType = "text/plain;charset=windows-1251"
ExpirePage()

' define variables and COM objects
dim Stream, DOM, File1, Description, SortValue, FeaturedFlag, ActiveFlag, Gallery, bRet, sFileName, ImageId, bImageExists
dim obRec, obSG, nImgX, nImgY, nThmbX, nThmbY

set Gallery = new CGallery
Gallery.Init( RequestString( "gid" ) )
if ( not Gallery.Exists ) then Response.End

' create Stream Object
set Stream = Server.CreateObject( "ADODB.Stream" )
' create XMLDOM object and load it from request ASP object
set DOM = Server.CreateObject( "MSXML2.DOMDocument" )
DOM.load( Request )

call Dom.setProperty( "SelectionLanguage", "XPath" )

dim ImageExists
set ImageExists = DOM.selectSingleNode( "//root/ImageExists" )

sFileName = ""
if ( ImageExists.nodeTypedValue = "Y" ) then bImageExists = true else bImageExists = false end if

if ( bImageExists ) then
   ' retrieve XML node with binary content
   set File1 = DOM.selectSingleNode( "//root/Image1" )

   ' open stream object and store XML node content into it   
   Stream.Type = 1  ' 1=adTypeBinary 
   Stream.open() 
   Stream.Write( File1.nodeTypedValue )
   ' save uploaded file
   sFileName = GUID( true ) & ".jpg"
   call Stream.SaveToFile( Gallery.Path & "/" & sFileName, 2 ) ' 2=adSaveCreateOverWrite 
   Stream.close()
end if

set ImageId = DOM.selectSingleNode( "//root/ImageId" )
set Description = DOM.selectSingleNode( "//root/Description" )
set SortValue = DOM.selectSingleNode( "//root/SortValue" )
set ActiveFlag = DOM.selectSingleNode( "//root/ActiveFlag" )
set FeaturedFlag = DOM.selectSingleNode( "//root/FeaturedFlag" )

bRet = true
if ( bImageExists ) then bRet = CreateThumbnail( Gallery.Path, sFileName )

if ( bRet ) then

   if ( bImageExists ) then 

      set obSG = CreateObject( "shotgraph.image" )
      call obSG.GetFileDimensions( Gallery.Path & "/" & sFileName, nImgX, nImgY )
      call obSG.GetFileDimensions( Gallery.Path & "/thmb/" & sFileName, nThmbX, nThmbY )
 
   end if

   dim sSQL
   if ( ImageId.nodeTypedValue <> "" ) then

      sSQL = "exec UpdateGalleryImage @ImageId=" & ImageId.nodeTypedValue &_
                                    ",@Description=" & ToSQLString( Description.nodeTypedValue ) &_
                                    ",@SortValue=" & ToSQLString( SortValue.nodeTypedValue ) &_
                                    ",@ActiveFlag=" & ToSQLString( ActiveFlag.nodeTypedValue ) &_
                                    ",@FeaturedFlag=" & ToSQLString( FeaturedFlag.nodeTypedValue )
      if ( sFileName <> "" ) then
         sSQL = sSQL &              ",@ImageName=" & ToSQLString( sFileName ) &_
                                    ",@Width=" & nImgX &_
                                    ",@Height=" & nImgY &_
                                    ",@ThumbWidth=" & nThmbX &_
                                    ",@ThumbHeight=" & nThmbY
      end if

   else

      sSQL = "exec AddGalleryImage @GalleryId=" & Gallery.Id &_
                                 ",@ImageName=" & ToSQLString( sFileName ) &_
                                 ",@Description=" & ToSQLString( Description.nodeTypedValue ) &_
                                 ",@SortValue=" & ToSQLString( SortValue.nodeTypedValue ) &_
                                 ",@ActiveFlag=" & ToSQLString( ActiveFlag.nodeTypedValue ) &_
                                 ",@FeaturedFlag=" & ToSQLString( FeaturedFlag.nodeTypedValue ) &_
                                 ",@Width=" & nImgX &_
                                 ",@Height=" & nImgY &_
                                 ",@ThumbWidth=" & nThmbX &_
                                 ",@ThumbHeight=" & nThmbY
   end if

   set obRec = obConn.Execute( sSQL )

   set obSG = nothing
   if ( IsNull( obRec( "Msg" ) ) ) then
      Response.Write( obRec( "ImageId" ) )
   else
      Response.Write( obRec( "Msg" ) )
   end if
   obRec.Close()

end if

set Stream = nothing  : set DOM = nothing : set Description = nothing : set SortValue = nothing : set ActiveFlag = nothing : set Gallery = nothing : set obRec = nothing

%>
<!--#include virtual="/Include/CloseDBConnect.asp"-->
