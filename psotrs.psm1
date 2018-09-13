#Requires -Version 3
param(
	$DebugMode = $false
	,$ResetStorage = $false
)

$ErrorActionPreference = "Stop";

if($ResetStorage){
	$Global:PSOtrs_Storage = $null;
}

## Global Var storing important values!
	if($Global:PSOtrs_Storage -eq $null){
		$Global:PSOtrs_Storage = @{
				SESSIONS = @()
				DEFAULT_SESSION = $null	
				WebService = $null
				
				#default endpoint mapping to call otrs.
				#	Suybkey is a identificaiton of method. Can be the function name...
				#		then, each subkey is:
				#			endpoint  	= The endpoint name (without starting /)
				#			method		= Acceptable methods. Defaults to DEFAULT_ENDPOINT_METHOD
				#
				WebServiceConfig = @{
					DefaultMethods = 'POST','GET'
					
					#Place where will store the configurations
					configs	= @{}
					
					#Default cofniguration
					default = @{
							'New-OtrsSession' = @{
								endpoint = 'CreateSession'
							}
							
							'Get-OtrsTicket' = @{
								endpoint = 'GetTicket'
							}
							
							'Search-OtrsTicket' = @{
								endpoint = 'SearchTicket'
							}
							
							'New-OtrsTicket' = @{
								endpoint = 'CreateTicket'
							}
							
							'Update-OtrsTicket' = @{
								endpoint = 'UpdateTicket'
							}
						}
				}
				
			}			
	}


## Helpers
#Make calls to a zabbix server url api.
	Function verbose {
		$ParentName = (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
		write-verbose ( $ParentName +':'+ ($Args -Join ' '))
	}
	
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
		write-verbose "$($MyInvocation.InvocationName): Converting from JSON!"
		$ResponseO = Otrs_ConvertFromJson $Response;
		
		write-verbose "$($MyInvocation.InvocationName): Checking properties of converted result!"
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
		if($ResponseO -is [hashtable]){
			return (New-Object PsObject -Prop $ResponseO);
		} else {
			return $ResponseO;
		}
	}

	#Converts objets to JSON and vice versa,
	Function Otrs_ConvertToJson($o) {
		
		if(Get-Command ConvertTo-Json -EA "SilentlyContinue"){
			write-verbose "$($MyInvocation.InvocationName): Using ConvertTo-Json"
			return Otrs_EscapeNonUnicodeJson(ConvertTo-Json $o -Depth 10);
		} else {
			write-verbose "$($MyInvocation.InvocationName): Using javascriptSerializer"
			Otrs_LoadJsonEngine
			$jo=new-object system.web.script.serialization.javascriptSerializer
			$jo.maxJsonLength=[int32]::maxvalue;
			return Otrs_EscapeNonUnicodeJson ($jo.Serialize($o))
		}
	}

	Function Otrs_ConvertFromJson([string]$json) {
	
		if(Get-Command ConvertFrom-Json  -EA "SilentlyContinue"){
			write-verbose "$($MyInvocation.InvocationName): Using ConvertFrom-Json"
			ConvertFrom-Json $json;
		} else {
			write-verbose "$($MyInvocation.InvocationName): Using javascriptSerializer"
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
				write-verbose "$($MyInvocation.InvocationName): Loading JSON engine!"
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
	
	#Thanks to CosmosKey answer in https://stackoverflow.com/questions/7468707/deep-copy-a-dictionary-hashtable-in-powershell
	function Otrs_CloneObject {
		param($DeepCopyObject)
		$memStream = new-object IO.MemoryStream
		$formatter = new-object Runtime.Serialization.Formatters.Binary.BinaryFormatter
		$formatter.Serialize($memStream,$DeepCopyObject)
		$memStream.Position=0
		$formatter.Deserialize($memStream)
	}
	
	#Gets a endpoint name...
	Function Otrs_BuildRequest {
		param(
			 $Url
			,$ConfigName	= $null
			,$EndpointName 	= $null
		)
		
		
		if(!$Url){
			throw "INVALID_BASE_URL"
		}
		
		verbose "Base url: $Url"
		
		if(!$EndpointName){
			#Gets the name from parent caller...
			$ParentInvocation = (Get-Variable MyInvocation -Scope 1).Value;
			$FuncName = $ParentInvocation.MyCommand.Name;
			$EndpointName = $FuncName;
		}
		
		verbose "EndpointName: $EndpointName"
		
		
		#Cannot determine...
		if(!$EndpointName){
			throw "CANNOT_DETERMINE_ENDPOINT";
		}
		
		if($ConfigName){
			$WsConfig = $Global:PSOtrs_Storage.WebServiceConfig.configs[$ConfigName];
		}
		
		if(!$WsConfig){
			verbose "Using the default WebService configuration"
			$WsConfig = $Global:PSOtrs_Storage.WebServiceConfig.default;
			$Config = 'DEFAULT';
		}
		
		$AvailableMethods = @($WsConfig.method);
		
		if(!$AvailableMethods){
			$AvailableMethods = @($Global:PSOtrs_Storage.WebServiceConfig.DefaultMethods)
		}
		
		
		$EndPointConfig  	= $WsConfig[$EndpointName]
		
		if(!$EndPointConfig){
			throw "ENDPOINTCONFIG_NOT_EXISTS: Config = $ConfigName EndPointName = $EndPointName"
		}
		
		$OtrsEndPoint		= $EndPointConfig.endpoint;
		
		
		if(!$OtrsEndPoint){
			throw "CANNOT_DETERMINE_OTRSENDPOINT: Config = $ConfigName EndPointName = $EndPointName"
		}
		
		verbose "EndpointName: $OtrsEndPoint HttpMethod = $AvailableMethods"
		
		return @{
			'url' 		= "$Url/$OtrsEndPoint"
			'method' 	= $AvailableMethods[0]
		}
		
	}
	
	
## OTRS Implementations!
	## This implementations depends on configured WEB SERVICE!
	## Our standard is: No value is passed on URL (tthats is, we no use :AttName)
	## Route maps to same Connector name!
	## We use this documentation as source: http://doc.otrs.com/doc/manual/admin/5.0/en/html/genericinterface.html
	## We send all request as POST
	
	<#
		.SYNOPSIS
			Opens new session with otrs
			
		.DESCRIPTION
				Creates a new session with otrs server and returns a object containing session information.
				The object returned is same as documented on otrs API.
				If errors occurs, exceptions is throws.
	#>
	Function New-OtrsSession {
		[CmdLetBinding()]
		param(
			$User
			,$Password
			,$Url
			,$WebService = $null
			,[switch]$NoNph = $false
			,$WebServiceConfig = $null
		)
		
		if(!$NoNph){
			$Url += '/nph-genericinterface.pl'
		}
		
		
		if($WebService){
			$Url += "/Webservice/$Webservice"
		}
		
		if(!$WebServiceConfig){
			$WebServiceConfig = $WebService;
		}
		
		$RequestData 	=  Otrs_BuildRequest -Url $Url -Config $WebServiceConfig
		
		$url	= $RequestData.url;
		$method	= $RequestData.method;
		
		$ResponseString = Otrs_CallUrl -data @{UserLogin=$User;Password=$Password} -url $url -Method $method;
		
		write-verbose "$($MyInvocation.InvocationName): Response received. Parsing result string!"
		$results =  (Otrs_TranslateResponseJson $ResponseString)
		$results | Add-Member -Type Noteproperty -Name _wsconfig -Value $WebServiceConfig;
		return $results;
	}

	Function Get-OtrsTicket {
		[CmdLetBinding()]
		param(
			#Specify same set of aceptable parameters
				[hashtable]$Filters = @{}
				
			,$Session = (Get-DefaultOtrsSession)
		)
		
		
		$RequestData 	=  Otrs_BuildRequest -Url $Session.RestUrl -Config $Session.WebServiceConfig
		$url	= $RequestData.url;
		$method	= $RequestData.method;
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Filters;
		
		
		
		$ResponseString = Otrs_CallUrl -url $url -data $MethodParams -Method $method
		$ResultTickets 	= (Otrs_TranslateResponseJson $ResponseString).Ticket;
		
		if($ResultTickets){
			return $ResultTickets
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
		
		
		$RequestData 	=  Otrs_BuildRequest -Url $Session.RestUrl -Config $Session.WebServiceConfig
		$url	= $RequestData.url;
		$method	= $RequestData.method;
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Filters;
		
		$ResponseString = Otrs_CallUrl -url $url -data $MethodParams -method $method
		$Tickets =  (Otrs_TranslateResponseJson $ResponseString);
		
		if($Tickets -and $GetTicket){
			verbose "Getting full ticket! Total tickets: $(@($Tickets.TicketID).count)"
			return Get-OtrsTicket -Session $Session @{TicketID=$Tickets.TicketID}
		}
		
		return $Tickets;
		
	}
	
	Function New-OtrsTicket {
		[CmdLetBinding()]
		param(
			$Session = (Get-DefaultOtrsSession)
			
			,#Specify same set of aceptable parameters
				[hashtable]$Attributes = @{}
		)
		
		$RequestData 	=  Otrs_BuildRequest -Url $Session.RestUrl -Config $Session.WebServiceConfig
		$url	= $RequestData.url;
		$method	= $RequestData.method;
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Attributes;
		
		$ResponseString = Otrs_CallUrl -url $url -data $MethodParams -method $method
		$Tickets =  (Otrs_TranslateResponseJson $ResponseString);

		
		return $Tickets;
		
	}

	Function Update-OtrsTicket {
		[CmdLetBinding()]
		param(
			$Session = (Get-DefaultOtrsSession)
			
			,#Specify same set of aceptable parameters
				[hashtable]$Attributes = @{}
		)
		
		$RequestData 	=  Otrs_BuildRequest -Url $Session.RestUrl -Config $Session.WebServiceConfig
		$url			= $RequestData.url;
		$method			= $RequestData.method;
		
		$MethodParams = @{
			SessionID=$Session.SessionId
		}
		
		$MethodParams += $Attributes;
		
		$ResponseString = Otrs_CallUrl -url $url -data $MethodParams -method $method
		$UpdateResult =  (Otrs_TranslateResponseJson $ResponseString);

		
		return $UpdateResult;
		
	}


	
# Facilities!	

<#
.SYNOPSIS
	Opens a new session with a otrs
	
.DESCRIPTION
	Connect-Otrs creates a new session with a otrs server.
	A object representing this session will be returned and you can pipe to other cmdlets (or pass in -Session parameter)
	If there are just one session, every cmdlet will use it as default.
	If there are more than one session, you must set a default using Set-DefaultOtrsSession, or pass the session to cmdlet.
	Calling this cmdlet with same -URL and -User, returns same session. If -Force is used, then a new connections i made.
#>
Function Connect-Otrs {
	[CmdLetBinding()]
	param(
		$Url
		,$User
		,$Password
		,$WebService
		,[switch]$Force
		,$WebServiceConfig
	)
	
	
	if(!$WebService){
		throw "PSOTRS_INVALID_WEBSERVICE";
	}
	
	if(!$WebServiceConfig){
		$WebServiceConfig = $WebService;
	}
	
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
		verbose "Getting from cache!"
		return $Session;
	}
	
	if(!$Session){
		verbose "Session object dont exist. Create new!"
		$Session = New-Object PSObject -Prop @{
				Url 		= $Url
				User 		= $User
				SessionID	= $null
				Webservice	= $WebService
				WebServiceConfig = $WebServiceConfig
				RestUrl		= "$Url/nph-genericinterface.pl/Webservice/$Webservice"
			}
		$IsNewSession = $true;
	}

	
	#Authenticates!
	$Session.SessionID = (New-OtrsSession -User $Session.User -Password $Password -Url $Session.RestUrl -WebServiceConfig $WebServiceConfig -NoNph).SessionID
	
	if($IsNewSession){
		verbose "Inserting on sessions cache"
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
	[CmdLetBinding()]
	param()
	
	if(@($Global:PSOtrs_Storage.SESSIONS).count -eq 1){
		$def =  @($Global:PSOtrs_Storage.SESSIONS)[0];
	} else {
		$def = $Global:PSOtrs_Storage.DEFAULT_SESSION
	}
	
	if(!$def){
		throw "NO_DEFAULT_OTRS_SESSION";
	}
	
	return $def;
	
}

Function Get-DefaultOtrsSessionId {
	[CmdLetBinding()]
	param()
	
	$d = Get-DefaultOtrsSession;
	
	if($d){return $d.SessionID};
}

Function Clear-PsOtrs {
	[CmdLetBinding()]
	param()
	
	$Global:PSOtrs_Storage = @{};
}


<#
.SYNOPSIS
	Get a template for mapping endpoints.
	
.DESCRIPTION
	Set the endpoints names to be used with a otrs session!
	All sessions uses the default.
	Use the "Get-EndPointMappingTemplate" to a list of all avilable endpoints to edit.
#>
Function Get-OtrsDefaultWSConfig {
	[CmdLetBinding()]
	param()
	
	return Otrs_CloneObject $Global:PSOtrs_Storage.WebServiceConfig.default;
}

<#
.SYNOPSIS
	Configure session web services and endpoints
	
.DESCRIPTION
	Set the endpoints names to be used with a otrs session!
	All sessions uses the default.
	Use the "Get-OtrsDefaultWSConfig" to a list of all avilable endpoints to edit.
#>
Function Set-OtrsWsConfig {
	[CmdletBinding()]
	Param(
		$ConfigName
		,$EndPointName
		,$OtrsEndPoint
		,$Method
		,[switch]$Remove
	)
	

	$WsConfigs = $Global:PSOtrs_Storage.WebServiceConfig.configs;
	
	
	#Gets the slot for current cofnig, if not exists, create a new one...
	if($WsConfigs[$ConfigName]){
		$Config = $WsConfigs[$ConfigName]
	} else {
		$Config = @{}
		$WsConfigs[$ConfigName] = $Config;
	}
	
	#Check the endpoint slot...
	$ValidEndpoints = (Get-OtrsDefaultWSConfig).keys
	if(-not($ValidEndpoints -Contains $EndPointName)){
		throw "INVALIDENDPOINT_NAME: Name = $EndPointName Valids = $ValidEndpoints"
	}
	
	$EndpointSlot = $Config[$EndPointName]
	if(!$EndpointSlot){
		$EndpointSlot = @{};
		$Config[$EndPointName] = $EndpointSlot;
	}
	
	if($Remove){
		$Config.remove($EndPointName);
	} else {
		$EndpointSlot.endpoint 	= $OtrsEndPoint
		$EndpointSlot.method 	= $Method	
	}


}
	
Function Get-WsConfig {
	[CmdletBinding()]
	param(
		$ConfigName
	)
	
	$WsConfigs = $Global:PSOtrs_Storage.WebServiceConfig.configs;
	return $WsConfigs[$ConfigName];
}
	

	
	
	
	
#Exporting all functions with Verb-Name
if(!$DebugMode){
	Export-ModuleMember -Function *-*;
}