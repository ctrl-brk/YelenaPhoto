<script language="VBScript" runat="Server">

sub OpenFramework

   Server.Execute( "/Framework/Open.asp" )

end sub

'-----------------------------------------------------------------------------

sub CloseFramework

   Server.Execute( "/Framework/Close.asp" )

end sub

'-----------------------------------------------------------------------------

sub ExpirePage()

   Response.Expires = 0
   Response.Expiresabsolute = Now() - 1442
   call Response.AddHeader( "pragma","no-cache" )
   call Response.AddHeader( "cache-control","private" )
   Response.CacheControl = "no-cache"

end sub

'-----------------------------------------------------------------------------

sub Redirect( sURL )

   Response.Redirect( sURL )
   Response.End()

end sub

'-----------------------------------------------------------------------------

function RandomNumber()

   dim nRnd
   Randomize()
   nRnd=Int((100000 - 1 + 1) * Rnd() + 1)
   RandomNumber = nRnd

end function

'-----------------------------------------------------------------------------

function PadLeft( sStr, cPad, nLen )

   PadLeft = CStr( sStr )
   while ( Len( PadLeft ) < nLen ) : PadLeft = cPad & PadLeft : wend

end function

'-----------------------------------------------------------------------------

function PadRight( sStr, cPad, nLen )

   PadRight = CStr( sStr )
   while ( Len( PadRight ) < nLen ) : PadRight = PadRight & cPad : wend

end function

'-----------------------------------------------------------------------------

function NullToStr( ByVal Value )

  if ( IsNull( Value ) ) then Value = "" end if
  NullToStr = Value

end function

'-----------------------------------------------------------------------------

function StrToNum( ByVal val )

   StrToNum = 0
   if ( IsNumeric( val ) ) then StrToNum = CLng( val )

end function

'-----------------------------------------------------------------------------

function ZeroToStr( ByVal Value )

  if ( IsNull( Value ) ) then
      Value = "&nbsp;"
  elseif ( Value = "" ) then
     Value = "&nbsp;"
  elseif ( IsNumeric( Value ) ) then
     if ( CLng( Value ) = 0 ) then Value = "&nbsp;"
  end if
 
  ZeroToStr = Value

end function

'-----------------------------------------------------------------------------

function EncodeSQLString( sSource )

   dim sResult

   sResult = Replace( sSource, "'", "''" )
   EncodeSQLString = sResult

end function

'-----------------------------------------------------------------------------

function ToSQL( ByVal sStr )

  if ( IsNull( sStr ) or IsEmpty( sStr ) or sStr = "" ) then
     sStr = "null"
  else
     sStr = EncodeSQLString( sStr )
  end if
 
  ToSQL = sStr

end function

'-----------------------------------------------------------------------------

function ToSQLString( ByVal sStr )

  if ( IsNull( sStr ) or sStr = "" ) then
     sStr = "null"
  else
     sStr = "'" & EncodeSQLString( sStr ) & "'"
  end if
 
  ToSQLString = sStr

end function

'-----------------------------------------------------------------------------

function ToSQLNum( sStr )

  if ( IsNull( sStr ) or IsEmpty( sStr ) or sStr = "" ) then
     ToSQLNum = "null"
  else
     ToSQLNum = CLng( sStr )
  end if

end function

'-----------------------------------------------------------------------------

function ToSQLFlag( bFlag )

   if ( bFlag = true ) then ToSQLFlag = "'Y'" else ToSQLFlag = "'N'" end if

end function

'-----------------------------------------------------------------------------

function InputToValue( ByVal sSrc, ByVal sT, ByVal sF )

   if ( IsNull( sSrc ) ) then sSrc = ""
   sSrc = UCase( sSrc )

   if ( sSrc = "" or sSrc = "0" or sSrc = "FALSE" or sSrc = "OFF" or sSrc = "NO" or sSrc = "N") then
      InputToValue = sF
   else
      InputToValue = sT
   end if

end function

'-----------------------------------------------------------------------------

function EmptyInputToValue( ByVal sSrc, ByVal sT, ByVal sF, ByVal sE )

   if ( IsNull( sSrc ) or sSrc = "" ) then
      EmptyInputToValue = sE
   else
      EmptyInputToValue = InputToValue( sSrc, sT, sF )
   end if

end function

'-----------------------------------------------------------------------------

function CreateSQLObject( sSQL )

   dim obRec, obField, obDict

   set obRec = SQLExecute( sSQL )

   if ( not obRec.EOF ) then
      set obDict = Server.CreateObject( "Scripting.Dictionary" )

      for each obField in obRec.Fields
         obDict.Add obField.Name, CStr( NullToStr( obField.Value ) )
      next

      set CreateSQLObject = obDict
   else
      set CreateSQLObject = Nothing
   end if

   obRec.Close()
   set obRec  = Nothing
   set obDict = Nothing

end function

'-----------------------------------------------------------------------------

function DecodeSQLString( sSource )

   dim sResult

   if ( IsNull( sSource ) ) then
      DecodeSQLString = ""
   else
   
      sResult=Replace(sSource,"&","&amp;")
      sResult=Replace(sResult,Chr(34),"&#34;")
      sResult=Replace(sResult,Chr(39),"&#39;")
      sResult=Replace(sResult,"<","&lt;")
      sResult=Replace(sResult,">","&gt;")
      sResult=Replace(sResult,vbCrLf,"<br>")

      if ( InStr( sResult, "%u" ) > 0 ) then

         dim bFirst, a, s, sTmp
         bFirst = true : sTmp = ""
         a = Split( sResult, "%u" )
         for each s in a
            if ( s <> "" and not bFirst ) then 
               sTmp = sTmp & "&#x" & Left( s, 4 ) & ";"
               if ( Len( s ) > 4 ) then sTmp = sTmp & Right( s, Len(s) - 4 )
            else
               sTmp = sTmp & s
            end if
            bFirst = false
         next

         sResult = sTmp

      end if

      DecodeSQLString = sResult

   end if

end function

'-----------------------------------------------------------------------------

function HighliteString( sSource, sSearch, sOpen, sClose )

   dim sResult, nStart

   nStart = InStr( 1, sSource, sSearch, 1 )
   sResult = Mid( sSource, 1, nStart-1 )
   sResult = sResult & sOpen & Mid( sSource, nStart, Len( sSearch ) ) & sClose
   sResult = sResult & Mid( sSource, nStart + Len( sSearch ) )
   HighliteString = sResult

end function

'-----------------------------------------------------------------------------

function RequestStringNoTrim( sName )

   dim sItem, sRet

   sRet = Request.QueryString( CStr( sName ) )
   if ( sRet = "" ) then
      if ( Left( Request.ServerVariables( "CONTENT_TYPE" ), 9 ) = "multipart" ) then
         on error resume next
         if ( IsObject( obUpload ) ) then
            for each sItem in obUpload.Form
               if ( sItem.Name = sName ) then
                  if ( sRet <> "" ) then sRet = sRet & ", "
                  sRet = sRet & sItem.Value
               end if
            next
         end if
         on error goto 0
      else
         sRet = Request.Form( sName )
      end if
   end if
   
   RequestStringNoTrim = sRet

end function

'-----------------------------------------------------------------------------

function RequestString( sReq )

   RequestString = Trim( RequestStringNoTrim( sReq ) )

end function

'-----------------------------------------------------------------------------

function RestoreRequestString()

   dim obDic, Item, sRes

   sRes = ""

   set obDic = Server.CreateObject( "Scripting.Dictionary" )

   for each Item in Request.QueryString
      if ( not obDic.Exists( Item ) ) then call obDic.Add( Item, Request.QueryString( Item ) ) end if
   next

   if ( Left( Request.ServerVariables( "CONTENT_TYPE" ), 9 ) = "multipart" ) then
      on error resume next
      for each Item in Session( "obUpload" ).Form
         if ( not obDic.Exists( Item ) ) then call obDic.Add( Item.Name, Item.Value ) end if
      next
      on error goto 0
   else
      for each Item in Request.Form
         if ( not obDic.Exists( Item ) ) then call obDic.Add( Item, Request.Form( Item ) ) end if
      next
   end if

   if ( obDic.Exists( "goto"            )) then obDic.Remove( "goto"            ) end if
   if ( obDic.Exists( "LoginEmail"      )) then obDic.Remove( "LoginEmail"      ) end if
   if ( obDic.Exists( "LoginPassword"   )) then obDic.Remove( "LoginPassword"   ) end if
   if ( obDic.Exists( "Password"        )) then obDic.Remove( "Password"        ) end if
   if ( obDic.Exists( "qlPassword"      )) then obDic.Remove( "qlPassword"      ) end if
   if ( obDic.Exists( "ConfirmPassword" )) then obDic.Remove( "ConfirmPassword" ) end if
   if ( obDic.Exists( "Email"           )) then obDic.Remove( "Email"           ) end if
   if ( obDic.Exists( "qlEmail"         )) then obDic.Remove( "qlEmail"         ) end if
   if ( obDic.Exists( "Rnd"             )) then obDic.Remove( "Rnd"             ) end if
   if ( obDic.Exists( "OnLoad"          )) then obDic.Remove( "OnLoad"          ) end if
   if ( obDic.Exists( "cus"             )) then obDic.Remove( "cus"             ) end if
   if ( obDic.Exists( "cucs"            )) then obDic.Remove( "cucs"            ) end if

   call obDic.Add( "Rnd", CStr( RandomNumber() ) )

   for each Item in obDic
      if ( sRes <> "" ) then sRes = sRes & "&" end if
      sRes = sRes & Item & "=" & obDic( Item )
   next
      
   set obDic = Nothing

   RestoreRequestString = sRes

end function

'-----------------------------------------------------------------------------

function RestoreRequest()

   dim obDic, Item, sRes

   sRes = ""

   set obDic = Server.CreateObject( "Scripting.Dictionary" )

   for each Item in Request.QueryString
      if ( not obDic.Exists( Item ) ) then call obDic.Add( Item, Request.QueryString( Item ) ) end if
   next

   if ( Left( Request.ServerVariables( "CONTENT_TYPE" ), 9 ) = "multipart" ) then
      on error resume next
      for each Item in Session( "obUpload" ).Form
         if ( not obDic.Exists( Item ) ) then call obDic.Add( Item.Name, Item.Value ) end if
      next
      on error goto 0
   else
      for each Item in Request.Form
         if ( not obDic.Exists( Item ) ) then call obDic.Add( Item, Request.Form( Item ) ) end if
      next
   end if

   if ( obDic.Exists( "goto"      )) then obDic.Remove( "goto"      ) end if
   if ( obDic.Exists( "Rnd"       )) then obDic.Remove( "Rnd"       ) end if
   if ( obDic.Exists( "OnLoad"    )) then obDic.Remove( "OnLoad"    ) end if

   for each Item in obDic
      if ( obDic( Item ) <> "" ) then sRes = sRes & "<input type=""hidden"" name=""" & Item & """ value=""" & obDic( Item ) & """>" & vbCrLf
   next
      
   set obDic = Nothing

   RestoreRequest = sRes

end function

'-----------------------------------------------------------------------------

function GetRegValue( sKey )
   
   GetRegValue = Application( "Settings\" & sKey )

end function

'-----------------------------------------------------------------------------

function hRef( sUrl )

   hRef = sUrl

end function

'-----------------------------------------------------------------------------

function TextAreaToJS( sText )

   TextAreaToJS = Replace( Replace( sText, vbCrLf, "\n" ), "<br>", "\n" )

end function

'-----------------------------------------------------------------------------

function TextAreaToHTML( sText )

   TextAreaToJS = Replace( sText, vbCrLf, "<br>" )

end function

'-----------------------------------------------------------------------------

function GUID( bSolid )

   dim obGUID

   set obGUID = Server.CreateObject( "GuidMakr.GUID" )

   if ( bSolid ) then
      GUID = Replace( Replace( Replace( obGUID.GetGUID(), "-", "" ), "}", "" ), "{", "" )
   else
      GUID = obGUID.GetGUID()
   end if

   set obGUID = nothing

end function

'-----------------------------------------------------------------------------

function FormatRFC822DateTime( dDate )

   dim aDays(7), aMonthes(12), nYear, iOffset

   iOffset = -7 'MST = GMT-7

   if ( not IsDate( dDate ) ) then dDate = Now()

   dDate = DateAdd( "h", iOffset, dDate )

   nYear = Year( dDate ) - 1900
   if ( nYear > 100 ) then nYear = nYear - 100
   aDays(0) = "Sun" : aDays(1) = "Mon" : aDays(2) = "Tue" : aDays(3) = "Wed" : aDays(4) = "Thu" : aDays(5) = "Fri" : aDays(6) = "Sat"
   aMonthes(0) = "Jan" : aMonthes(1) = "Feb" : aMonthes(2) = "Mar"  : aMonthes(3) = "Apr" : aMonthes(4) = "May" : aMonthes(5) = "Jun" : aMonthes(6) = "Jul" : aMonthes(7) = "Aug" : aMonthes(8) = "Sep" : aMonthes(9) = "Oct" : aMonthes(10) = "Nov" : aMonthes(11) = "Dec"

   FormatRFC822DateTime = aDays( Weekday( dDate ) - 1 ) & ", " & Day( dDate ) & " " & aMonthes( Month( dDate ) - 1 ) & " " & PadLeft( nYear, "0", 2 ) & " " &_
                          PadLeft( Hour( dDate ), "0", 2 ) & ":" & PadLeft( Minute( dDate ), "0", 2 ) & ":" & PadLeft( Second( dDate ), "0", 2 ) & " GMT"

end function

'-----------------------------------------------------------------------------

function FormatW3CDateTime( dDate )

   dim iOffset

   iOffset = -7 'MST = GMT-7

   if ( not IsDate( dDate ) ) then dDate = Now()

   dDate = DateAdd( "h", iOffset, dDate )

   FormatW3CDateTime = Year( dDate ) & "-" & PadLeft( Month( dDate ), "0", 2 ) & "-" & PadLeft( Day( dDate ), "0", 2 ) & "T" &_
                       PadLeft( Hour( dDate ), "0", 2 ) & ":" & PadLeft( Minute( dDate ), "0", 2 ) & ":" & PadLeft( Second( dDate ), "0", 2 ) & "Z"

end function

'-----------------------------------------------------------------------------

function RFC822ToVBDateTime( sDate )

   dim arParts, iMonth, iYear, iDay, iHour, iMinute, iSecond, dtD, iOffset

   iOffset = -7 'MST = GMT-7

   arParts = Split( Replace( Replace( sDate, ":", " " ), ",", "" ), " " )

   'Thu, 16 Oct 2003 00:04:30 GMT
   '0    1  2   3    4  5  6  7

   if ( UBound( arParts ) < 7 ) then
      RFC822ToVBDateTime = ""
   else
      iDay = CInt( arParts(1) )

      select case UCase( arParts(2) )
         case "JAN" : iMonth = 1
         case "FEB" : iMonth = 2
         case "MAR" : iMonth = 3
         case "APR" : iMonth = 4
         case "MAY" : iMonth = 5
         case "JUN" : iMonth = 6
         case "JUL" : iMonth = 7
         case "AUG" : iMonth = 8
         case "SEP" : iMonth = 9
         case "OCT" : iMonth = 10
         case "NOV" : iMonth = 11
         case "DEC" : iMonth = 12
      end select

      iYear = CInt( arParts(3) )
      if ( iYear < 50 ) then iYear = iYear + 2000
      if ( iYear < 100 ) then iYear = iYear + 1900
      iHour = CInt( arParts(4) )
      iMinute = CInt( arParts(5) )
      iSecond = CInt( arParts(6) )

      dtD = CDate( DateValue( DateSerial( iYear, iMonth, iDay ) ) & " " & TimeValue( TimeSerial( iHour, iMinute, iSecond ) ) )

      if arParts( 7 ) = "GMT" then dtD = DateAdd( "h", iOffset, dtD )

      RFC822ToVBDateTime = dtD
   end if

end function

'-----------------------------------------------------------------------------

function IsRobot()

   if ( InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Yandex" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Aport" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Rambler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "ia_archiver" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Mediapartners-Google" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "libwww-perl" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "eStyleSearch" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "IRLbot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "NetStat.Ru" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "NG/" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "appie " ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "LWP::Simple" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Microsoft URL Control" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Speedy Spider" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "FAST" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "moen_" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "OmniExplorer_Bot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "msnbot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Accoona-AI-Agent" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Baiduspider" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Biz360 spider" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "ConveraCrawler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Exabot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Gigabot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "IRLbot" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Krugle" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "MultiCrawler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Nusearch Spider" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Sensis" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Speedy Spider" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "StackSearch Crawler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Vespa Crawler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "WebAlta Crawler" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Yahoo! Slurp" ) > 0 or _
        InStr( Request.ServerVariables( "HTTP_USER_AGENT" ), "Googlebot" ) > 0 ) then
      IsRobot = true
   else
      IsRobot = false
   end if

end function

'-----------------------------------------------------------------------------

function EmailLink( sType, sParams )

   dim sAddr, aName, aDomain
   
   aName = Split( GetRegValue( sType ), "@" )
   aDomain = Split( aName(1), "." )

   EmailLink = "<scr" & "ipt language='javascript'>PrintEmLink('" & aName(0) & aDomain(0) & aDomain(1) & "'," & Len( aName(0) ) & "," & Len( aDomain(0) ) & ",'" & sParams & "');</scr" & "ipt>"

end function

'-----------------------------------------------------------------------------

sub SendMail( bHtml, sSubj, sBody, sFrom, sTo, sCC, sBCC )

   dim obCDO, obCDOConf, ns
                       
   ns = "http://schemas.microsoft.com/cdo/configuration/"

   set obCDOConf = CreateObject("CDO.Configuration") 
   set obCDO = CreateObject("CDO.Message")

   obCDOConf.Fields.Item( ns & "sendusing" ) = 2 '1 = pickup folder; 2 = smtp
   obCDOConf.Fields.Item( ns & "smtpserver" ) = "10.0.0.1"
   obCDOConf.Fields.Update()

   set obCDO.Configuration = obCDOConf
   obCDO.From = sFrom
   obCDO.To = sTo
   obCDO.Subject = sSubj
   if ( bHtml = true ) then
      obCDO.HTMLBody = sBody
   else
      obCDO.TextBody = sBody
   end if
      
   obCDO.Send()
         
   set obCDO = Nothing : set obCDOConf = nothing

end sub

'-----------------------------------------------------------------------------

function GetEmailAddress( sType, sEmail )

   GetEmailAddress = GetRegValue( "Emails\" & sType & "\" & sEmail )

end function

'-----------------------------------------------------------------------------

function Send500Mail( bSend )

   dim obASPError, sBody, sDesc, strKey, Item

   set obASPError = Server.GetLastError()

   if ( obASPError.ASPDescription <> "" ) then sDesc = obASPError.ASPDescription else sDesc = obASPError.Description end if
   
   sBody = "<html><head>" &_
           "<style>" & vbCrLf &_
           " body {font: 8pt/11pt Verdana; background: white}" & vbCrLf &_
           " .Norm {font: 8pt/11pt Verdana;}" & vbCrLf &_
           " .Err {font: 8pt/11pt Verdana; font-weight: bold;}" & vbCrLf &_
           " .Src {font: 10px Courier}" & vbCrLf &_
           "</style>" & vbCrLf &_
           "</head><body>" & vbCrLf &_
           "<b>Script error 500 occured on " & CStr( Now() )  & "</b><br><br>" &_
           "<table border='0'>" &_
           "<tr><td class='Norm'>Server:</td><td class='Norm'>"   & Request.ServerVariables("SERVER_NAME"    ) & "</td></tr>" &_
           "<tr><td class='Norm'>Port:</td><td class='Norm'>"     & Request.ServerVariables("SERVER_PORT"    ) & "</td></tr>" &_
           "<tr><td class='Norm'>Method:</td><td class='Norm'>"   & Request.ServerVariables("REQUEST_METHOD" ) & "</td></tr>" &_
           "<tr><td class='Norm'>Client:</td><td class='Norm'>"   & Request.ServerVariables("REMOTE_ADDR"    ) & "</td></tr>" &_
           "<tr><td class='Norm' valign='top'>ReqStr:</td><td class='Norm'>" & Request.QueryString             & "</td></tr>"
   
   sBody = sBody & "<tr><td class='Norm' valign='top'>Form:</td><td class='Norm'>"
   if ( Left( Request.ServerVariables( "CONTENT_TYPE" ), 9 ) = "multipart" ) then
      on error resume next
      for each Item in Session( "obUpload" ).Form
         sBody = sBody & "&" & Item.Name & "=" & Item.Value
      next
      on error goto 0
   else
      sBody = sBody & Request.Form
   end if
   sBody = sBody & "</td></tr>"

   sBody = sBody & "<tr><td class='Norm' valign='top'>SQL:</td><td class='Norm'>" & Replace( Session( "sDbgSQL" ), " ", "&nbsp;" ) & "</td></tr>" &_
                   "<tr><td class='Norm' valign='top'>Dbg:</td><td class='Norm'>" & Replace( Session( "sDbgStr" ), " ", "&nbsp;" ) & "</td></tr>" &_
                   "</table>" &_
                   "<hr size='1'>" &_
                   "<table border='0'>" &_
                   "<tr><td class='Norm' nowrap>" & obASPError.Category

   if ( obASPError.ASPCode > "" ) then
      sBody = sBody & Server.HTMLEncode(", " & obASPError.ASPCode)
   end if

   sBody = sBody & Server.HTMLEncode(" (0x" & Hex( obASPError.Number ) & ")" ) & "</td></tr>"

   if ( sDesc <> "" ) then
      sBody = sBody & "<tr><td class='Norm'>" & sDesc & "</td></tr>"
   end if

   if ( obASPError.File <> "?" ) then
      sBody = sBody & "<tr><td class='Err'>" & obASPError.File
      if ( obASPError.Line   > 0 ) then sBody = sBody & ", line "   & obASPError.Line
      if ( obASPError.Column > 0 ) then sBody = sBody & ", column " & obASPError.Column
      sBody = sBody & "</td></tr><tr><td class='Src' nowrap>"
      sBody = sBody & Server.HTMLEncode( obASPError.Source() ) & "<br>"
      if ( obASPError.Column > 0 ) then sBody = sBody & String( obASPError.Column - 1, "-" ) & "<font color='red'>^</font>"
      sBody = sBody & "</td></tr>"
   end if

   sBody = sBody & "</table>" &_
                   "<hr size='1'>" &_
                   "<table border='1'>"

   on error resume next
   for each strKey in Session.Contents
      sBody = sBody & "<tr><td class='Norm'>" & strKey & "</td><td class='Norm'>" & Session(strKey) & "&nbsp;</td></tr>"
   next
   on error goto 0

   sBody = sBody & "</table>" &_
                   "<hr size='1'>" &_
                   "<table border='1'>"

   for each strKey in Request.ServerVariables
      sBody = sBody & "<tr><td class='Norm'>" & strKey & "</td><td class='Norm'>" & Request.ServerVariables(strKey) & "&nbsp;</td></tr>"
   next
   sBody = sBody & "</table>" &_
                   "</body></html>"

   if ( bSend ) then
    
      call SendMail( true, "YelenaPhoto.com 500 Error", sBody, GetEmailAddress( "From", "System" ), GetEmailAddress( "To", "Error" ), "", "" )

   end if

   set obASPError = nothing

   Send500Mail = sBody

end function

'-----------------------------------------------------------------------------

sub WriteLog( sData, bClear )

   dim obFS, obFile, bExists
   
   set obFS = Server.CreateObject( "Scripting.FileSystemObject" )
   bExists = obFS.FileExists( Server.MapPath( GetRegValue( "LogFile" ) ) )
   if ( bClear ) then
      set obFile = obFS.CreateTextFile( Server.MapPath( GetRegValue( "LogFile" ) ) )
   else
      set obFile = obFS.OpenTextFile( Server.MapPath( GetRegValue( "LogFile" ) ), 2, true )
   end if
   obFile.WriteLine( sData )
   obFile.Close()
   set obFile = nothing : set obFS = nothing

end sub


'-----------------------------------------------------------------------------

function ThumbnailImage( sRelPath, sImageName )

   ThumbnailImage = sRelPath & "/thmb/" & sImageName

end function

'-----------------------------------------------------------------------------

function ResizeImage( sFrom, sTo, nMaxX, nMaxY )

   dim Pal, nX, nY, nDiv, obSG, bRes

   bRes = true

   set obSG = CreateObject( "shotgraph.image" )

   call obSG.GetFileDimensions( sFrom, nX, nY )

   nDiv = nX/nMaxX
   if ( nY/nMaxY > nDiv ) then nDiv = nY/nMaxY

   if ( nDiv = 0 or nX = 0 or nY = 0 ) then
      bRes = false
   else
      if ( nX > nMaxX or nY > nMaxY ) then
         call obSG.CreateImage( nX/nDiv, nY/nDiv, 256 )
         call obSG.InitClipboard( nX, nY )
         call obSG.SelectClipboard( true )
      else
         call obSG.CreateImage( nX, nY, 256 )
      end if
   
      call obSG.ReadImage( sFrom, Pal, 0, 0 )
   
      if ( nX > nMaxX or nY > nMaxY ) then
         call obSG.Resize( 0, 0, nX/nDiv, nY/nDiv, 0, 0, nX, nY )
         call obSG.SelectClipboard( false )
         call obSG.Sharpen()
      end if
   
      call obSG.JpegImage( 90, 0, sTo )
   end if
   
   set obSG = nothing

   ResizeImage = bRes

end function

'-----------------------------------------------------------------------------

function CreateThumbnail( sGalleryPath, sFileName )

   dim obFS

   set obFS = CreateObject( "Scripting.FileSystemObject" )
   if ( not obFS.FolderExists( sGalleryPath & "/thmb" ) ) then
      obFS.CreateFolder( sGalleryPath & "/thmb" )
   end if
   set obFS = nothing

   CreateThumbnail = ResizeImage( sGalleryPath & "/" & sFileName, sGalleryPath & "/thmb/" & sFileName, 100, 100 )

end function

'-----------------------------------------------------------------------------

function CreateGalleryThumbnail( sSrcName, sDestName )

   CreateGalleryThumbnail = ResizeImage( sSrcName, sDestName, 100, 100 )

end function

'-----------------------------------------------------------------------------

function ExecuteBlob( sSQL )

   set ExecuteBlob = Server.CreateObject( "ADODB.Recordset" )
   ExecuteBlob.CursorLocation = 3
   call ExecuteBlob.Open( sSQL, obConn, 3 )

end function

</script>
