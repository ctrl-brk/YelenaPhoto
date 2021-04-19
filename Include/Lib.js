
function RandRange( nMin, nMax )
{
   return( Math.floor( Math.random() * ( nMax-nMin+1 ) ) + nMin );
}

//----------------------------------------------------------------------------

function StrRep( sSource, sFind, sReplace )
{
   if ( sFind != sReplace ) return( sSource.split( sFind ).join( sReplace ) );
   return( sSource );
}

//----------------------------------------------------------------------------

function LTrim( str )
{
   if ( !str ) return( str );

   var whitespace = " \t\n\r";

   if ( whitespace.indexOf( str.charAt(0) ) != -1)
      {
      var j=0, i = str.length;
      while ( j < i && whitespace.indexOf( str.charAt(j) ) != -1) j++;
      str = str.substring( j, i );
      }
   return( str );
}

//----------------------------------------------------------------------------

function RTrim( str )
{
   if ( !str ) return( str );

   var whitespace = " \t\n\r";

   if ( whitespace.indexOf( str.charAt( str.length-1 ) ) != -1 )
      {
      var i = str.length - 1;
      while ( i >= 0 && whitespace.indexOf( str.charAt(i) ) != -1 ) i--;
      str = str.substring( 0, i+1 );
      }
   return( str );
}

//----------------------------------------------------------------------------

function Trim(str)
{
   return ( RTrim( LTrim( str ) ) );
}

//----------------------------------------------------------------------------

function ToSQLString( sStr )
{
   if ( !sStr || Trim( sStr ) == "" ) return( "null" );
   return( "'" + StrRep( Trim( sStr ), "'", "''" ) + "'" );
}

//----------------------------------------------------------------------------

function IsValue( sVal )
{
   return( sVal && sVal.toString() != "" );
}

//----------------------------------------------------------------------------

function hRef( sURL )
{
   return( sURL );
}

//----------------------------------------------------------------------------

function Redirect( sURL )
{
   document.location = hRef( sURL );
}

//----------------------------------------------------------------------------

function IsIE()
{
   return( navigator.userAgent && navigator.userAgent.indexOf( "MSIE" ) >= 0 && !IsOpera() );
}

//----------------------------------------------------------------------------

function IsNN()
{
   return( navigator.appName == "Netscape" );
}

//----------------------------------------------------------------------------

function IsOpera()
{
   return( window.opera ? true : false );
}

//----------------------------------------------------------------------------

function PrintEmLink( s, l1, l2, pars )
{
   document.write( "<a href='ma" );
   document.write( "ilto" );
   document.write( ":" + s.substring( 0, l1 ) );
   document.write( "@" );
   document.write( s.substring( l1, l1+l2 ) );
   document.write( "." );
   document.write( s.substring( l1 + l2 ) );
   document.write( "'" );
   if ( pars != "" ) document.write( " " + pars );
   document.write( ">" );
   document.write( s.substring( 0, l1 ) );
   document.write( "@" );
   document.write( s.substring( l1, l1+l2 ) );
   document.write( "." );
   document.write( s.substring( l1 + l2 ) );
   document.write( "</a>" );

}

//----------------------------------------------------------------------------
