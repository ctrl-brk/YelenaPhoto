var gobValidation;
var gValMgr = null;

//----------------------------------------------------------------------------
//Old stuff
function ValidationData()
{
   this.bWarningGiven = false;
   this.bFocusSet     = false;
   this.sWarnings     = "";
   this.nWarnings     = 0;
   this.bOK           = true;
}

//----------------------------------------------------------------------------

function Validators()
{
   this.Value = function( obCtrl, bGetMsg )
                {
                   var sMsg = "Required field";
                   if ( bGetMsg ) return( sMsg );

                   obCtrl.Value = GetControlValueDirect( obCtrl.Control, true );
                   if ( !obCtrl.Value ) { obCtrl.SetValidatorText( sMsg ); return( false ); }
                   return( true );
                };

   this.Digit = function( obCtrl, bGetMsg )
                {
                   var sMsg = "Numeric only";
                   if ( bGetMsg ) return( sMsg );

                   obCtrl.Value = GetControlValueDirect( obCtrl.Control, true );
                   if ( !isValidDigit( obCtrl.Value ) ) { obCtrl.SetValidatorText( sMsg ); return( false ); }
                   return( true );
                };

   this.NoZero = function( obCtrl, bGetMsg )
                 {
                    var sMsg = "Must not be zero";
                    if ( bGetMsg ) return( sMsg );

                    obCtrl.Value = GetControlValueDirect( obCtrl.Control, false, true );
                    if ( !obCtrl.Value ) { obCtrl.SetValidatorText( sMsg ); return( false ); }
                    return( true );
                 };

   this.ZIP = function( obCtrl, bGetMsg )
                 {
                    if ( bGetMsg ) return( null );

                    obCtrl.Value = GetControlValueDirect( obCtrl.Control, false, true );
                    if ( !isValidZIP( obCtrl.Value ) )
                       {
                       obCtrl.SetValidatorText( "Invalid format" );
                       return( false );
                       }
                    return( true );
                 };

   this.Phone = function( obCtrl, bGetMsg )
                {
                   if ( bGetMsg ) return( null );

                   obCtrl.Value = GetControlValueDirect( obCtrl.Control, false );
                   if ( !isValidPhone( obCtrl.Value ) )
                      {
                      obCtrl.SetValidatorText( "Invalid format" );
                      return( false );
                      }
                   return( true );
                };

   this.Email = function( obCtrl, bGetMsg )
                {
                   if ( bGetMsg ) return( null );

                   obCtrl.Value = GetControlValueDirect( obCtrl.Control, false );
                   if ( !isValidEmail( obCtrl.Value ) )
                      {
                      obCtrl.SetValidatorText( gValMgr.Clipboard );
                      return( false );
                      }
                   return( true );
                };

   this.CCNum = function( obCtrl, bGetMsg )
                {
                   if ( bGetMsg ) return( null );

                   obCtrl.Value = GetControlValueDirect( obCtrl.Control, false );
                   if ( !isValidCCNum( obCtrl.Value ) )
                      {
                      obCtrl.SetValidatorText( "Invalid number" );
                      return( false );
                      }
                   return( true );
                };
}

//----------------------------------------------------------------------------

function HideValidationTooltip( sName )
{
   if ( gValMgr.bBlur ) gValMgr.ShowControlTooltip( sName, false );
}

//----------------------------------------------------------------------------

function Tooltip( obParent )
{
   this.Parent = obParent;
   this.x = 5;
   this.y = null;
   this.Container = null;
   this.bVisible = false;
   this.bMouseOn = false;

   this.Locate = function()
                 {
                    if ( this.y ) return;
                    var obCtrl = this.Parent.Control;
                    this.y = obCtrl.offsetHeight;
                    for( ; obCtrl; obCtrl = obCtrl.offsetParent ) { this.x += obCtrl.offsetLeft; this.y += obCtrl.offsetTop; }
                 };

   this.CreateContainer = function()
                          {
                             if ( this.Container != null ) return;
                             this.Container = document.createElement( "div" );
                             this.Container.id = "tt_" + this.Parent.Name;
                             this.Container.className = "tooltip";
                             this.Container.noWrap = true;
                             this.Container.style.left = this.x; this.Container.style.top = this.y;
                             if ( gValMgr.bIE )
                                {
                                this.Container.style.filter = "progid:DXImageTransform.Microsoft.DropShadow(OffX=2,OffY=2,Color='#2C2422',Positive='true')";
                                }
                             document.body.appendChild( this.Container );
                          };

   this.Show = function( bShow )
               {
                  this.Locate();
                  if ( bShow )
                     {
                     this.CreateContainer();
                     this.Container.innerHTML = this.Parent.GetErrorText();
                     this.Container.style.display = "";
                     this.Parent.Control.SetTitle( null );
                     if ( this.Parent.bAutoHide ) window.setTimeout( "HideValidationTooltip('" + this.Parent.Name + "');", 2000 );
                     }
                  else if ( this.Container )
                     {
                     if ( !this.bMouseOn )
                        {
                        this.Container.style.display = "none";
                        this.Parent.Control.SetTitle( this.Parent.GetErrorText() );
                        }
                     else { if ( this.Parent.bAutoHide ) window.setTimeout( "HideValidationTooltip('" + this.Parent.Name + "');", 300 );}
                     }
               };
}

//----------------------------------------------------------------------------

function ValidateBlur( sName )
{
   gValMgr.ValidateControl( sName );
}

//----------------------------------------------------------------------------

function ValidationControl( obCtrl, sName, sErrorText, sDesc, bTooltip, bHL, bAutoHide )
{
   this.Control = obCtrl;
   this.Name = sName;
   this.Description = sDesc;
   this.Value = null;
   this.ErrorText = sErrorText;
   this.ValidatorText = null;
   this.Validators = new Array;
   this.bValid = true;
   this.Tooltip = null;
   this.bTooltip = bTooltip;
   this.bHL = bHL;
   this.bAutoHide = bAutoHide;
   this.bBlur = false;

   if ( this.bTooltip && this.bAutoHide ) { this.bBlur = true; this.Control.onblur = new Function( "ValidateBlur('" + this.Name + "');" ); }

   obCtrl.SetTitle = function( sTitle ) { this.title = sTitle ? sTitle : ""; };

   this.SetDescription = function( sDesc )
                         {
                         if ( sDesc == "" ) this.Description = this.Name;
                         else this.Description = sDesc;
                         };

   this.SetValid = function( bValid )
                   {
                      this.bValid = bValid;
                      if ( bValid ) { this.ErrorText = null; this.ValidatorText = null; }
                   };

   this.SetValidatorText = function( sText ) { this.ValidatorText = sText; };

   this.GetErrorText = function() { return( this.ErrorText ? this.ErrorText : this.ValidatorText ); };

   this.AddError = function( sText, bInfo )
                   {
                      this.ErrorText = sText;
                      if ( !bInfo ) this.bValid = false;
                   };

   this.AddValidator = function( fnValidator, bSetTitle )
                       {
                          if ( bSetTitle && !this.Control.title ) this.Control.SetTitle( fnValidator( this.Control, true ) );
                          this.Validators[this.Validators.length] = fnValidator;
                       };

   this.Validate = function()
                   {
                      var bValid = true;
                      for ( var i=0; i<this.Validators.length && bValid; i++ )
                         {
                         bValid = this.Validators[i]( this );
                         }
                      this.SetValid( bValid );
                      this.Refresh();
                      if ( this.bBlur ) this.bAutoHide = true;
                      return( this.bValid );
                   };

   this.ShowTooltip = function( bShow )
                      {
                         if ( this.Tooltip ) this.Tooltip.Show( bShow );
                      };

   this.Refresh = function()
                  {
                     if ( this.bHL ) SetValidClassName( this.Control, this.bValid );
                     if ( this.bValid && this.Tooltip ) this.ShowTooltip( false );
                     if ( !this.bValid )
                        {
                        if ( this.bTooltip )
                           {
                           if ( !this.Tooltip ) this.Tooltip = new Tooltip( this );
                           this.ShowTooltip( true );
                           }
                        }
                  };

   this.SetDescription( sDesc );
}

//----------------------------------------------------------------------------

function ValidationManager()
{
   this.bIE = IsIE();
   this.bNN = IsNN();
   this.bOpera = IsOpera();

   this.Controls = new Array;
   this.Validators = new Validators();
   this.Clipboard = null;
   this.bBlur = true;

   this.GetControl = function( sName )
                     {
                        for ( var i=0; i<this.Controls.length; i++ )
                           {
                           if ( this.Controls[i].Name == sName ) return( this.Controls[i] );
                           }
                        return( null );
                     };

   this.AddControl = function( sName, sErrorText, sDesc, bTooltip, bHL, bAutoHide )
                     {
                        if ( !bAutoHide ) this.bBlur = false;

                        var obCtrl = this.GetControl( sName );

                        if ( !obCtrl )
                           {
                           obCtrl = GetControl( sName );
                           if ( !obCtrl ) return( null );

                           this.Controls[this.Controls.length] = new ValidationControl( obCtrl, sName, sErrorText, sDesc, bTooltip, bHL, bAutoHide );
                           obCtrl = this.Controls[this.Controls.length-1];
                           }
                        else
                           {
                           obCtrl.SetDescription( sDesc );
                           obCtrl.SetValid( true );
                           obCtrl.Control.SetTitle( null );
                           obCtrl.Validators.length = 0;
                           obCtrl.Refresh();
                           obCtrl.bTooltip = bTooltip;
                           obCtrl.bHL = bHL;
                           obCtrl.bAutoHide = bAutoHide;
                           }
                        return( obCtrl );
                     };

   this.ValidateControl = function( sName )
                          {
                             var obCtrl;

                             if ( typeof( sName ) == "object" ) obCtrl = sName;
                             else obCtrl = this.GetControl( sName );

                             if ( obCtrl ) return( obCtrl.Validate() );
                          };

   this.SetControlValid = function( sName, bValid, sErrorText, sDesc )
                          {
                             var obCtrl;

                             if ( typeof( sName ) == "object" ) obCtrl = sName;
                             else obCtrl = this.GetControl( sName );

                             if ( obCtrl )
                                {
                                obCtrl.SetValid( bValid );
                                if ( sDesc ) obCtrl.Description = sDesc;
                                if ( bValid ) obCtrl.ErrorText = null;
                                else if ( sErrorText ) obCtrl.ErrorText = sErrorText;
                                obCtrl.Refresh();
                                return( obCtrl.bValid );
                                }
                          };

   this.ShowControlTooltip = function( sName, bShow )
                             {
                                var obCtrl = this.GetControl( sName );
                                if ( obCtrl ) obCtrl.ShowTooltip( bShow );
                             };
      
   this.IsValid = function()
                  {
                     for ( var i=0; i<this.Controls.length; i++ )
                        {
                        if ( !this.Controls[i].bValid ) return( false );
                        }
                     return( true );
                  };

   this.ValidationMsg = function()
                        {
                           var sMsg = "";

                           for ( var i=0; i<this.Controls.length; i++ )
                              {
                              if ( !this.Controls[i].bValid && this.Controls[i].ErrorText )
                                 {
                                 if ( i > 0 ) sMsg += "\n";
                                 sMsg += ( this.Controls[i].Description ? "\"" + this.Controls[i].Description + "\" - " : "" ) + this.Controls[i].ErrorText;
                                 }
                              }
                           return( sMsg );
                        };

   this.ValidationAlert = function()
                          {
                             var sMsg = this.ValidationMsg();
                             if ( sMsg != "" ) alert( sMsg );
                          };
}

//----------------------------------------------------------------------------

function CreateValidationControl( sCtrl, sErrorText, sDesc, bTooltip, bHL, bAutoHide )
{
   if ( !gValMgr ) gValMgr = new ValidationManager();
   return( gValMgr.AddControl( sCtrl, sErrorText, sDesc, bTooltip, bHL, bAutoHide ) );
}

//----------------------------------------------------------------------------

function SetControlValidState( sCtrl, bValid, sErrorText, sDesc )
{
   return( gValMgr.SetControlValid( sCtrl, bValid, sErrorText, sDesc ) );
}

//----------------------------------------------------------------------------


//============================================================================

function BeginValidation()
{
   gobValidation = new ValidationData();
}

//----------------------------------------------------------------------------

function AddValidationWarning( sWarn )
{
   if ( gobValidation.nWarnings ) gobValidation.sWarnings += "\n";
   gobValidation.sWarnings += (++gobValidation.nWarnings) + ") " + sWarn;
   gobValidation.bOK = false;
}

//----------------------------------------------------------------------------

function ValidationWarning( sWarn, bAlways )
{
   if ( bAlways || !gobValidation.bWarningGiven ) alert( sWarn );
   gobValidation.bWarningGiven = true;
   gobValidation.bOK = false;
}

//----------------------------------------------------------------------------

function FinalValidationWarning( sTxt, bAlways )
{
   if ( gValMgr ) { gValMgr.ValidationAlert(); return; }

   var sWarn;

   if ( sTxt ) AddValidationWarning( sTxt );

   if ( gobValidation.nWarnings == 0 ) return;

   if ( bAlways || !gobValidation.bWarningGiven )
      {
      sWarn = gobValidation.sWarnings;
      if ( gobValidation.nWarnings == 1 ) sWarn = StrRep( sWarn, "1) ", " " );
      alert( sWarn );
      }
   gobValidation.bWarningGiven = true;
}

//----------------------------------------------------------------------------

function SetValidationFocus( sName )
{
   if ( !sName ) sName = GetControlByClass( "hl" );
   if ( !sName ) return;
   SetControlFocus( sName );
}

//----------------------------------------------------------------------------

function ValidationError()
{
   gobValidation.bOK = false;
   return( false );
}

//----------------------------------------------------------------------------

function IsValidationOK()
{
   return( gValMgr ? gValMgr.IsValid() : gobValidation.bOK );
}



//============================================================================

function _isValidEmail( emailStr )
{
   var DomsPat      = /^(com|net|org|edu|int|mil|gov|pro|biz|nom|web|arpa|aero|name|nato|coop|info|firm|store|museum)$/;
   var knownDomsPat = typeof( gp_Domains ) == "undefined" ? DomsPat : gp_Domains;

   var checkTLD     = 1;
   var emailPat     = /^(.+)@(.+)$/;
   var specialChars = "\\(\\)><@,;:\\\\\\\"\\.\\[\\]";
   var validChars   = "\[^\\s" + specialChars + "\]";
   var quotedUser   = "(\"[^\"]*\")";
   var ipDomainPat  = /^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;
   var atom         = validChars + '+';
   var word         = "(" + atom + "|" + quotedUser + ")";
   var userPat      = new RegExp( "^" + word + "(\\." + word + ")*$" );
   var domainPat    = new RegExp( "^" + atom + "(\\." + atom +")*$" );

   var matchArray = emailStr.toLowerCase().match( emailPat );

   if ( matchArray == null)
      {
      if ( gValMgr ) gValMgr.Clipboard = "Address seems incorrect (check @ and .'s)";
      return( false );
      }

   var user   = matchArray[1];
   var domain = matchArray[2];

   for ( var i=0; i<user.length; i++ )
      {
      if ( user.charCodeAt(i) > 127 )
         {
         if ( gValMgr ) gValMgr.Clipboard = "The username contains invalid characters.";
         return( false );
         }
      }

   for ( var i=0; i<domain.length; i++ )
      {
      if ( domain.charCodeAt(i) > 127 )
         {
         if ( gValMgr ) gValMgr.Clipboard = "The domain name contains invalid characters.";
         return( false );
         }
      }

   if ( user.match( userPat ) == null )
      {
      if ( gValMgr ) gValMgr.Clipboard = "The username doesn't seem to be valid.";
      return( false );
      }

   var IPArray = domain.match( ipDomainPat );
   if ( IPArray != null )
      {
      for ( var i=1; i<=4; i++ )
         {
         if ( IPArray[i] > 255 )
            {
            if ( gValMgr ) gValMgr.Clipboard = "Destination IP address is invalid!";
            return( false );
            }
         }
      return( true );
      }

   var atomPat = new RegExp( "^" + atom + "$" );
   var domArr  = domain.split( "." );
   var len     = domArr.length;

   for (var i=0; i<len; i++ )
      {
      if ( domArr[i].search( atomPat ) == -1 )
         {
         if ( gValMgr ) gValMgr.Clipboard = "The domain name does not seem to be valid.";
         return( false );
         }
      }

//   if ( checkTLD && domArr[domArr.length-1].length != 2 && domArr[domArr.length-1].search(knownDomsPat) == -1 )
   if ( checkTLD && domArr[domArr.length-1].search(knownDomsPat) == -1 )
      {
      if ( gValMgr ) gValMgr.Clipboard = "The address must end in a well-known domain.";
      return( false );
      }

   if ( len < 2 )
      {
      if ( gValMgr ) gValMgr.Clipboard = "Address is missing a hostname!";
      return( false );
      }

   return( true );
}

//----------------------------------------------------------------------------

function isValidEmails( sEmails, bReq )
{
   var nAddrs = 0, bValid = true;

   if ( IsValue( sEmails ) )
      {
      sEmails = StrRep( sEmails, " ", ";" );
      var aEmails = sEmails.split( ";" );
      for ( var i = 0; i<aEmails.length && bValid; i++ )
         {
         if ( aEmails[i].length > 0 )
            {
            if ( !_isValidEmail( aEmails[i] ) ) bValid = false; 
            nAddrs++;
            }
         }
      if ( bValid && nAddrs == 0 ) bValid = !bReq;
      return( bValid );
      }
   else return( !bReq );
}        

//----------------------------------------------------------------------------

function isValidEmail( sEmail, bReq, bMultiple )
{
   if ( IsValue( sEmail ) )
      {
      if ( !bMultiple ) return( _isValidEmail( sEmail ) );
      else return( isValidEmails( sEmail, bReq ) );
      }
   else return( !bReq );
}        

//----------------------------------------------------------------------------

function isValidAlpha( sStr, bReq )
{
   if ( IsValue( sStr ) )
      {
      var re = new RegExp( /[^A-Za-z0-9 ]/ );
      return( !re.test( sStr ) );
      }
   else return( !bReq );
}       

//----------------------------------------------------------------------------

function isValidDigit( sStr, bReq )
{
   if ( IsValue( sStr ) )
      {
      var re = new RegExp( /[\D]/ );
      return( !re.test( sStr ) );
      }
   else return( !bReq );
}       

//----------------------------------------------------------------------------

function isValidZIP( sZIP, bReq, sType )
{
  if ( !sZIP ) return( bReq ? false : true );

  sZIP = StripChars( sZIP, " -" );

  sZIP = sZIP ? sZIP : "";

  if ( sZIP.length != 0 && sZIP.length != 5 && sZIP.length != 6 && sZIP.length != 9 ) return( false );
  if ( sZIP.length != 0 && sZIP.length != 5 && sZIP.length != 9 && sType == "US" ) return( false );
  if ( sZIP.length != 0 && sZIP.length != 6 && sType == "CAN" ) return( false );

  if ( bReq && sZIP.length == 0 ) return( false );

  if ( sZIP.length == 5 || sZIP.length == 9 ) return ( isValidDigit( sZIP, bReq ));

  var sAlpha = sZIP.charAt(0) + sZIP.charAt(2) + sZIP.charAt(4);
  var sDigit = sZIP.charAt(1) + sZIP.charAt(3) + sZIP.charAt(5);

  return( isValidAlpha( sAlpha, false ) && isValidDigit( sDigit, false ) );
}

//----------------------------------------------------------------------------

//============================================================================

function ValidateRequiredControl( sName, sDesc, sMsg, bNoZero )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );

   obCtrl.AddValidator( gValMgr.Validators.Value, true );
   if ( bNoZero ) obCtrl.AddValidator( gValMgr.Validators.NoZero, true );
}       

//----------------------------------------------------------------------------

function ValidateDigitControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.Digit, true );
}       

//----------------------------------------------------------------------------

function isValidDigitCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.Digit );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidateEmailControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.Email, true );
}       

//----------------------------------------------------------------------------

function isValidEmailCtrl( sName, bReq, sDesc, sMsg, bMultiple )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.Email );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidatePhoneControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.Phone, true );
}       

//----------------------------------------------------------------------------

function isValidPhoneCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.Phone );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidateZIPControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.ZIP, true );
}       

//----------------------------------------------------------------------------

function isValidZIPCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.ZIP );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidateCCNumControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.CCNum, true );
}       

//----------------------------------------------------------------------------

function isValidCCNumCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.CCNum );

   return( gValMgr.ValidateControl( obCtrl ) );
}

//----------------------------------------------------------------------------

function DetermineControl( Ctrl, sForm )
{
   if ( Ctrl.toString().indexOf( "[object" ) != 0 ) Ctrl = GetControl( Ctrl, sForm );

   return( Ctrl );
}

//----------------------------------------------------------------------------

function GetControlByClass( sName, sForm )
{
   var obCtrl = null, obForm = sForm ? document.forms[ sForm ] : document.forms[0];

   for( var i=0; i < obForm.elements.length; i++ )
      {
      if ( obForm.elements[i] != null )
         {
         if ( obForm.elements[i].className == sName || obForm.elements[i].className.indexOf( " " + sName ) >=0 || obForm.elements[i].className.indexOf( sName + " " ) >=0 )
            {
            obCtrl = obForm.elements[i];
            break;              
            }
         }
      }
   return( obCtrl );
}

//----------------------------------------------------------------------------

function GetControlLoop( sName )
{
   var obCtrl = null;

   for( var i=0; i < document.forms[0].elements.length; i++ )
      {
      if ( document.forms[0].elements[i] != null )
         {
         if ( document.forms[0].elements[i].name == sName )
            {
            obCtrl = document.forms[0].elements[i];
            break;              
            }
         }
      }
   return( obCtrl );
}

//----------------------------------------------------------------------------

function GetControl( sName, sForm )
{
   var obCtrl = null, obForm = sForm ? document.forms[ sForm ] : document.forms[0];

   obCtrl = obForm.elements[ sName ];

   obCtrl = typeof( obCtrl ) == "undefined" ? null : obCtrl;
   
   if ( obCtrl && typeof( obCtrl.type ) == "undefined" ) obCtrl = GetControlLoop( sName );
   
   return( obCtrl  );
}

//----------------------------------------------------------------------------

function GetControlValueDirect( obCtrl, bReq, bNoZero, nMaxLen, sForm )
{
   var Ret = null;

   if ( !obCtrl )
      {
      alert( "GetControlValueDirect error: Control is null" );
      return( false );
      }

   if ( obCtrl.type == "radio" )
      {
      var obParent = document.forms[0];

      if ( sForm ) obParent = document.forms[sForm];
      for ( var i=0; i < obParent.elements.length; i++ )        
         {
         if (obParent.elements[i].type == "radio" && obParent.elements[i].name == obCtrl.name) 
            {
            if ( obParent.elements[i].checked )
               {                 
               Ret = obParent.elements[i].value;
               break;
               }
            }
         }
      }
   else if ( obCtrl.type == "select" || obCtrl.type == "select-one")
      {    
      if ( obCtrl.selectedIndex != -1 )
         Ret = obCtrl.options[obCtrl.selectedIndex].value;
      }
   else if ( obCtrl.type == "select-multiple" )
      {
      for ( var i=0; i < obCtrl.length; i++ )
         {
         if ( obCtrl.options[i].selected )
            {                
            if ( !Ret ) Ret = "";
            if ( Ret != "" ) Ret += ", ";
            Ret += obCtrl.options[i].value;
            }
         }
      }
   else if ( obCtrl.type == "text" || obCtrl.type == "textarea" || obCtrl.type == "hidden" || obCtrl.type == "password" || obCtrl.type == "file" )
      {
      Ret = "" + obCtrl.value;
      var len = parseInt( Ret.length, 10 );
      if ( len > 0 )
         {
         for ( var i=len-1; i >= 0; i-- )
            {
            if ( Ret.charAt(i) != " " ) break;
            }
         if ( i > -1 ) Ret = Ret.substring( 0, i+1 );
         }
      }
   else if ( obCtrl.type == "checkbox" )
      {
      if ( obCtrl.checked ) Ret = "on";
      }

   else alert( "GetControlValueDirect error: Unknown object type '" + obCtrl.type + "'" + obCtrl.name);

   if ( Ret )
      {
      Ret = Trim( Ret );
      if ( Ret.toString() == "" || ( bNoZero && parseFloat( Ret ) == 0 ) ) Ret = null;
      if ( nMaxLen && Ret.toString().length > nMaxLen ) Ret = null;
      }
   if ( bReq && !gValMgr ) SetValidClassName( obCtrl, Ret ? true : false, sForm );
   return( Ret );
}

//----------------------------------------------------------------------------

function _GetControlValue( sName, bReq, bNoZero, nMaxLen, sForm )
{
   var obCtrl = GetControl( sName, sForm );

   if ( !obCtrl )
      {
      alert( "GetControlValue error: Object '" + sName + "' not found");
      return( false );
      }

   return( GetControlValueDirect( obCtrl, bReq, bNoZero, nMaxLen, sForm ) );
}

//----------------------------------------------------------------------------

function GetControlValue( sName, bReq, sDesc, sMsg, bNoZero )
{
   var obCtrl;

   if ( !bReq && !bNoZero ) return( _GetControlValue( sName ) );

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   if ( bNoZero ) obCtrl.AddValidator( gValMgr.Validators.NoZero );

   gValMgr.ValidateControl( obCtrl );
   return( obCtrl.Value );
}       

//----------------------------------------------------------------------------

function SetControlValue( sName, sValue, bNoWarn, sForm )
{
   var obCtrl = GetControl( sName, sForm );

   if ( !obCtrl )
      {
      if( !bNoWarn ) alert( "SetControlValue error: Object '" + sName + "' not found");
      return( null );
      }

   if ( obCtrl.type == "file" ) return( obCtrl );

   if ( obCtrl.type == "text" || obCtrl.type == "textarea" || obCtrl.type == "hidden" || obCtrl.type == "password" || obCtrl.type == "file" ) obCtrl.value = sValue;
   else if ( obCtrl.type == "select-one" )
      {
      for( var i=0; i < obCtrl.options.length; i++ )
         {
         if ( obCtrl.options[i].value == sValue )
            {
            obCtrl.selectedIndex = i;
            break;
            }
         }
      }
   else if ( obCtrl.type == "select-multiple" )
      {
      var aOpt = sValue.split( ", " );
      
      for( var i=0; i < obCtrl.options.length; i++ ) obCtrl.options[i].selected = false;

      for( i=0; i < obCtrl.options.length; i++ )
         {
         for ( var j=0; j<aOpt.length; j++ )
            {
            if ( obCtrl.options[i].value == aOpt[j] ) obCtrl.options[i].selected = true;
            }
         }
      }
   else if ( obCtrl.type == "radio" )
      {
      var obParent = document.forms[0];
      if ( sForm ) obParent = document.forms[sForm];

      for ( var i=0; i < obParent.elements.length; i++ )
         {
         if ( obParent.elements[i].type == "radio" && obParent.elements[i].name == sName ) 
            {
            if ( obParent.elements[i].value == sValue || obParent.elements[i].value.toString() == ( sValue.toString() == "True" ? "1" : "0" ) )
               {
               obParent.elements[i].checked = true;
               break;
               }
            else obParent.elements[i].checked = false;           
            }
         }
      }
   else if ( obCtrl.type == "checkbox")
      {
      if( typeof( sValue ) == "boolean" ) obCtrl.checked = sValue;
      else if ( typeof( sValue ) == "string" )
         {
         sValue = sValue.toLowerCase();
         if ( sValue == "true" || sValue == "on" || sValue == "yes" || sValue == "y" ) obCtrl.checked = true;
         else if ( sValue == "" || sValue == "false" || sValue == "off" || sValue == "no" || sValue == "n" ) obCtrl.checked = false;
         }
      else
         {
         var nValue = parseInt( sValue, 10 );
         if ( isNaN( nValue )) nValue = 0;
         if ( nValue == 0 ) obCtrl.checked = false;
         else obCtrl.checked = true;
         }
      }
   else if ( obCtrl.type == "button")
      {
      if ( typeof( sValue ) == "string" ) obCtrl.value = sValue;
      }

   else alert("SetControlValue error: Unknown object type '" + obCtrl.type +"'" );

   return( obCtrl );
}

//----------------------------------------------------------------------------

function SetControlFocus( sName, sForm )
{
   var obCtrl = DetermineControl( sName, sForm );

   if ( !obCtrl )
      {
      alert( "SetControlFocus error: Object '" + sName.toString() + "' not found" );
      return( false );
      }

   obCtrl.focus();

   return( obCtrl );
}

//----------------------------------------------------------------------------

function SetValidClassName( obCtrl, bValid, sForm )
{
   var sClass = StrRep( StrRep( obCtrl.className, " hl", "" ), "hl", "" ); 

   if ( sClass != "" && !bValid ) sClass += " ";
   sClass += bValid ? "" : "hl";

   if ( obCtrl.type == "radio" )
      {
      var obParent = document.forms[0];
      if ( sForm ) obParent = document.forms[sForm];

      for ( var i=0; i < obParent.elements.length; i++ )        
         {
         if (obParent.elements[i].type == "radio" && obParent.elements[i].name == obCtrl.name ) 
            {
            obParent.elements[i].className = sClass;
            }
         }
      }
   else obCtrl.className = sClass;
}

//----------------------------------------------------------------------------

function SetValidCtrlClassName( sName, bValid, sForm )
{
   var obCtrl = GetControl( sName, sForm );

   if ( obCtrl ) SetValidClassName( obCtrl, bValid, sForm );
   else alert( "SetValidCtrlClassName error: Object '" + sName + "' not found" );
}

//----------------------------------------------------------------------------

function ValidateDigitControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.Digit, true );
}       

//----------------------------------------------------------------------------

function isValidPhoneCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.Phone );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidatePhoneControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.Phone, true );
}       

//----------------------------------------------------------------------------

function isValidZIPCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.ZIP );

   return( gValMgr.ValidateControl( obCtrl ) );
}       

//----------------------------------------------------------------------------

function ValidateZIPControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.ZIP, true );
}       

//----------------------------------------------------------------------------

function isValidCCNumCtrl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value );
   obCtrl.AddValidator( gValMgr.Validators.CCNum );

   return( gValMgr.ValidateControl( obCtrl ) );
}

//----------------------------------------------------------------------------

function ValidateCCNumControl( sName, bReq, sDesc, sMsg )
{
   var obCtrl;

   obCtrl = CreateValidationControl( sName, sMsg, sDesc, true, false, true );
   
   if ( bReq ) obCtrl.AddValidator( gValMgr.Validators.Value, true );
   obCtrl.AddValidator( gValMgr.Validators.CCNum, true );
}       

//----------------------------------------------------------------------------
