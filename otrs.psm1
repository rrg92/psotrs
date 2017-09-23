$ErrorActionPreference = "Stop";

## Global Var storing important values!
	if($Global:PSOtrs_Storage -eq $null){
		$Global:PSOtrs_Storage = @{
				SESSIONS = @()
				DEFAULT_SESSION = $null	
				WebService = $null
			}
	}


## Helpers
#Make calls to a zabbix server url api.
	Function Otrs_CallUrl([object]$data = $null,$url = $null,$method = "POST", $contentType = "application/json"){
		$ErrorActionPreference="Stop";

		try {
			if(!$data){
				$data = "";
			}
		
			if($data -is [hashtable]){
				write-verbose "Converting input object to json string..."
				$data = Otrs_ConvertToJson $data;
			}
			
			write-verbose "$($MyInvocation.InvocationName):  json that will be send is: $data"
			
			write-verbose "Usando URL: $URL"
		
			write-verbose "$($MyInvocation.InvocationName):  Creating WebRequest method... Url: $url. Method: $Method ContentType: $ContentType";
			$Web = [System.Net.WebRequest]::Create($url);
			$Web.Method = $method;
			$Web.ContentType = $contentType
			
			if($data){
				#Determina a quantidade de bytes...
				[Byte[]]$bytes = [byte[]][char[]]$data;
				
				#Escrevendo os dados
				$Web.ContentLength = $bytes.Length;
				write-verbose "$($MyInvocation.InvocationName):  Bytes lengths: $($Web.ContentLength)"
				
				
				write-verbose "$($MyInvocation.InvocationName):  Getting request stream...."
				$RequestStream = $Web.GetRequestStream();
				
				
				try {
					write-verbose "$($MyInvocation.InvocationName):  Writing bytes to the request stream...";
					$RequestStream.Write($bytes, 0, $bytes.length);
				} finally {
					write-verbose "$($MyInvocation.InvocationName):  Disposing the request stream!"
					$RequestStream.Dispose() #This must be called after writing!
				}
			}
			
			
			
			write-verbose "$($MyInvocation.InvocationName):  Making http request... Waiting for the response..."
			$HttpResp = $Web.GetResponse();
			
			
			
			$responseString  = $null;
			
			if($HttpResp){
				write-verbose "$($MyInvocation.InvocationName):  charset: $($HttpResp.CharacterSet) encoding: $($HttpResp.ContentEncoding). ContentType: $($HttpResp.ContentType)"
				write-verbose "$($MyInvocation.InvocationName):  Getting response stream..."
				$ResponseStream  = $HttpResp.GetResponseStream();
				
				write-verbose "$($MyInvocation.InvocationName):  Response stream size: $($ResponseStream.Length) bytes"
				
				$IO = New-Object System.IO.StreamReader($ResponseStream);
				
				write-verbose "$($MyInvocation.InvocationName):  Reading response stream...."
				$responseString = $IO.ReadToEnd();
				
				write-verbose "$($MyInvocation.InvocationName):  response json is: $responseString"
			}
			
			
			write-verbose "$($MyInvocation.InvocationName):  Response String size: $($responseString.length) characters! "
			return $responseString;
		} catch {
			throw "ERROR_INVOKING_URL: $_";
		} finally {
			if($IO){
				$IO.close()
			}
			
			if($ResponseStream){
				$ResponseStream.Close()
			}
			
			<#
			if($HttpResp){
				write-host "Finazling http request stream..."
				$HttpResp.finalize()
			}
			#>

		
			if($RequestStream){
				write-verbose "Finazling request stream..."
				$RequestStream.Close()
			}
		}
	}

	Function Otrs_TranslateResponseJson {
		param($Response)
		
		#Converts the response to a object.
		$ResponseO = Otrs_ConvertFromJson $Response;
		
		#Check outputs
		if($ResponseO.Error -ne $null){
			$ResponseError = $ResponseO.Error;
			$MessageException = "[$($ResponseError.ErrorCode)]: $($ResponseError.ErrorMessage)";
			$Exception = New-Object System.Exception($MessageException)
			$Exception.Source = "OtrsGenericInterface"
			throw $Exception;
			return;
		}
		
		
		#If not error, then return response result.
		return New-Object PsObject -Prop $ResponseO;
	}

	#Converts objets to JSON and vice versa,
	Function Otrs_ConvertToJson($o) {
		
		if(Get-Command ConvertTo-Json -EA "SilentlyContinue"){
			return Otrs_EscapeNonUnicodeJson(ConvertTo-Json $o);
		} else {
			Otrs_LoadJsonEngine
			$jo=new-object system.web.script.serialization.javascriptSerializer
			$jo.maxJsonLength=[int32]::maxvalue;
			return Otrs_EscapeNonUnicodeJson ($jo.Serialize($o))
		}
	}

	Function Otrs_ConvertFromJson([string]$json) {
	
		if(Get-Command ConvertFrom-Json  -EA "SilentlyContinue"){
			ConvertFrom-Json $json;
		} else {
			Otrs_LoadJsonEngine
			$jo=new-object system.web.script.serialization.javascriptSerializer
			$jo.maxJsonLength=[int32]::maxvalue;
			return $jo.DeserializeObject($json);
		}
		

	}
	
	Function Otrs_CheckAssembly {
		param($Name)
		
		if($Global:PsOtrs_Loaded){
			return $true;
		}
		
		if( [appdomain]::currentdomain.getassemblies() | ? {$_ -match $Name}){
			$Global:PsOtrs_Loaded = $true
			return $true;
		} else {
			return $false
		}
	}
	
	Function Otrs_LoadJsonEngine {

		$Engine = "System.Web.Extensions"

		if(!(Otrs_CheckAssembly $Engine)){
			try {
				Add-Type -Assembly  $Engine
				$Global:PsOtrs_Loaded = $true;
			} catch {
				throw "ERROR_LOADIING_WEB_EXTENSIONS: $_";
			}
		}

	}

	#Troca caracteres não-unicode por um \u + codigo!
	#Solucao adapatada da resposta do Douglas em: http://stackoverflow.com/a/25349901/4100116
	Function Otrs_EscapeNonUnicodeJson {
		param([string]$Json)
		
		$Replacer = {
			param($m)
			
			return [string]::format('\u{0:x4}', [int]$m.Value[0] )
		}
		
		$RegEx = [regex]'[^\x00-\x7F]';
		write-verbose "$($MyInvocation.InvocationName):  Original Json: $Json";
		$ReplacedJSon = $RegEx.replace( $Json, $Replacer)
		write-verbose "$($MyInvocation.InvocationName):  NonUnicode Json: $ReplacedJson";
		return $ReplacedJSon;
	}
	
## OTRS Implementations!
	## This implementations depends on configured WEB SERVICE!
	## Our standard is: No value is passed on URL (tthats is, we no use :AttName)
	## Route maps to same Connector name!
	## We use this documentation as source: http://doc.otrs.com/doc/manual/admin/5.0/en/html/genericinterface.html
	## We send all request as POST
	
	Function New-OtrsSession {
		[CmdLetBinding()]
		param($User,$Password, $Url, $WebService = $null)
		
		$MethodName = 'CreateSession'
		
		if($WebService){
			$Url += "/Webservice/$Webservice"
		}
		
		$Url2Call 	=  "$Url/$MethodName"
		
		$ResponseString = Otrs_CallUrl -data @{UserLogin=$User;Password=$Password} -url $Url2Call
		
		return (Otrs_TranslateResponseJson $ResponseString)
	}

	Function Get-OtrsTicket {
		[CmdLetBinding()]
		param(
			$Session = (Get-DefaultOtrsSession)
			
			,#Specify same set of aceptable parameters
				[hashtable]$Filters = @{}
		)
		
		$MethodName = 'GetTicket'
		$Url2Call 	=  "$($Session.RestUrl)/$MethodName"
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Filters;
		
		
		
		$ResponseString = Otrs_CallUrl -url $Url2Call -data $MethodParams
		$ResultTickets 	= (Otrs_TranslateResponseJson $ResponseString).Ticket;
		
		if($ResultTickets){
			return @($ResultTickets | %{
				New-Object PsObject -Prop $_
			})
		}
	}

	Function Search-OtrsTicket {
		[CmdLetBinding()]
		param(
			$Session = (Get-DefaultOtrsSession)
			
			,#Specify same set of aceptable parameters
				[hashtable]$Filters = @{}
				
			,#Finds the tickets using Get-OtrsTicket cmdlet!
				[switch]$GetTicket = $false
		)
		
		$MethodName = 'SearchTicket'
		$Url2Call 	=  "$($Session.RestUrl)/$MethodName"
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Filters;
		
		$ResponseString = Otrs_CallUrl -url $Url2Call -data $MethodParams
		$Tickets =  (Otrs_TranslateResponseJson $ResponseString);
		
		if($Tickets -and $GetTicket){
			write-verbose "$($MyInvocation.InvocationName): Getting full ticket! Total tickets: $(@($Tickets.TicketID).count)"
			return Get-OtrsTicket -Session $Session @{TicketID=$Tickets.TicketID}
		}
		
		return $Tickets;
		
	}
	
	Function Create-OtrsTicket {
		[CmdLetBinding()]
		param(
			$Session = (Get-DefaultOtrsSession)
			
			,#Specify same set of aceptable parameters
				[hashtable]$Attributes = @{}
				
			,#Finds the tickets using Get-OtrsTicket cmdlet!
				[switch]$GetTicket = $false
		)
		
		$MethodName = 'CreateTicket'
		$Url2Call 	=  "$($Session.RestUrl)/$MethodName"
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Attributes;
		
		$ResponseString = Otrs_CallUrl -url $Url2Call -data $MethodParams
		$Tickets =  (Otrs_TranslateResponseJson $ResponseString);

		
		return $Tickets;
		
	}

# Facilities!	

	#Authenticates in a otrs and stores in a session array!
	Function Auth-Otrs {
		[CmdLetBinding()]
		param($Url, $User, $Password, $WebService, [switch]$Force)
		
		#Gets a session from cache!
		$AllSessions 	= $Global:PSOtrs_Storage.SESSIONS
		
		if(!$User){
			$Creds 	= Get-Credential
			$User	= $Creds.GetNetworkCredential().UserName
			$Password	= $Creds.GetNetworkCredential().Password
		}
		
		#Find a session with same name and url!
		$Session = $AllSessions | ? {  $_.Url -eq $Url -and $_.User -eq $User };
		
		if(!$Force -and $Session){
			write-verbose "$($MyInvocation.InvocationName): Getting from cache!"
			return $Session;
		}
		
		if(!$Session){
			write-verbose "$($MyInvocation.InvocationName): Session object dont exist. Create new!"
			$Session = New-Object PSObject -Prop @{
					Url 		= $Url
					User 		= $User
					SessionID	= $null
					Webservice	= $WebService
					RestUrl		= "$Url/nph-genericinterface.pl/Webservice/$Webservice"
				}
			$IsNewSession = $true;
		}
	
		
		#Authenticates!
		$Session.SessionID = (New-OtrsSession -User $Session.User -Password $Password -Url $Session.RestUrl).SessionID
		
		if($IsNewSession){
			write-verbose "$($MyInvocation.InvocationName): Inserting on sessions cache"
			$Global:PSOtrs_Storage.SESSIONS += $Session;
		}
		
		
		return $Session;
	}
	
	Function Set-DefaultOtrsSession {
		[CmdLetBinding()]
		param(
			
			[Parameter(Mandatory=$True, ValueFromPipeline=$true)]
			$Session
		
		)
		
		begin {}
		process {}
		end {
			$Global:PSOtrs_Storage.DEFAULT_SESSION = $Session;
		}
		
	}
	
	Function Get-DefaultOtrsSession {
		
		if(@($Global:PSOtrs_Storage.SESSIONS).count -eq 1){
			return @($Global:PSOtrs_Storage.SESSIONS)[0];
		} else {
			return $Global:PSOtrs_Storage.DEFAULT_SESSION
		}
		
	}

	Function Get-DefaultOtrsSessionId {
		$d = Get-DefaultOtrsSession;
		
		if($d){return $d.SessionID};
	}

	
	
	Function Clean-PsOtrs {
		$Global:PSOtrs_Storage = @{};
	}