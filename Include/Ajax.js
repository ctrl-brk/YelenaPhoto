var _obAjaxReq, _fnAjaxCallBack, _fnAjaxBefore, _fnAjaxAfter;

function _AjaxCallBack()
{
   if ( _fnAjaxCallBack && _obAjaxReq && _obAjaxReq.readyState == 4 && _obAjaxReq.status == 200 )
      {
      if ( _fnAjaxAfter != null ) _fnAjaxAfter();
      _fnAjaxCallBack( _obAjaxReq.getResponseHeader( "content-type" ).indexOf( "xml" ) >= 0 ? _obAjaxReq.responseXML : _obAjaxReq.responseText, _obAjaxReq );
      }
}

function AjaxRequest( sURL, obData, fnCallBack, fnBefore, fnAfter )
{
   var sRet = null;

   _obAjaxReq = null;
   _fnAjaxCallBack = fnCallBack == "none" ? null : fnCallBack;
   _fnAjaxBefore = fnBefore; _fnAjaxAfter = fnAfter;
   if ( window.XMLHttpRequest ) _obAjaxReq = new XMLHttpRequest();
   else if ( window.ActiveXObject ) _obAjaxReq = new ActiveXObject( "Microsoft.XMLHTTP" );
   if ( _obAjaxReq )
      {
      if ( _fnAjaxBefore != null ) _fnAjaxBefore();
      if ( _fnAjaxCallBack != null ) _obAjaxReq.onreadystatechange = _AjaxCallBack;
      _obAjaxReq.open( obData == null ? "GET" : "POST", sURL, _fnAjaxCallBack != null );
      _obAjaxReq.send( obData );
      }
   
   if ( fnCallBack == null && _fnAjaxAfter != null ) _fnAjaxAfter();

   if ( fnCallBack == null && _obAjaxReq.status == 200 )
      {
      sRet = _obAjaxReq.getResponseHeader( "content-type" ).indexOf( "xml" ) >= 0 ? _obAjaxReq.responseXML : _obAjaxReq.responseText;
      _obAjaxReq = null;
      }
   else alert( _obAjaxReq.getResponseHeader( "content-type" ).indexOf( "xml" ) >= 0 ? _obAjaxReq.responseXML : _obAjaxReq.responseText );

   return( sRet );
}
