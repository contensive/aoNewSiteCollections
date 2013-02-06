'
' New Site onInatall
'
function onInstall()
	'
	const landingPageGuid = "{925F4A57-32F7-44D9-9027-A91EF966FB0D}"
	const pageNotFoundGuid = "{2DB43A39-473D-42FF-ACE2-20094FBD3537}"
	'
	dim cs
	dim recordId
	dim http
	dim sourceUrl
	dim imageFilename
	dim copy
	dim agentCopy
	dim officeCopy
	dim link
	dim agentPhotoLink
	dim officeLogoLink
	dim agentFirstName
	dim agentLastName
	'
	set cs = cp.csNew()
	Call csv.SetSiteProperty("ALLOWLINKALIAS", "1", 0)
	'
	' set link aliases for all records in page content
	'   - this needs to be moved to the base content collection
	'
	onInstall = onInstall & "<br>set link aliases"
	call cs.open( "page content")
	Do While cs.ok()
		name = cs.getText("name")
		alias = "/" & Replace(name, " ", "-")
		Call verifyLinkAlias(name, alias)
		Call cs.goNext()
	Loop
	Call cs.close()
	'
	' set first landing page id, delete the rest
	'
	onInstall = onInstall & "<br>set landing page"
	if cs.open( "page content", "ccguid=" & cp.db.encodeSqlText( landingPageGuid )) then
		recordId = cs.getInteger( "id" )
		call cp.site.setProperty( "LANDINGPAGEID", recordId)
		'
		' delete any duplcate Guids
		'
		call cs.goNext()
		Do While cs.ok()
			call cs.Delete()
			Call cs.goNext()
		Loop
	end if
	Call cs.close()
	'
	' set page not found
	'
	onInstall = onInstall & "<br>set page not found"
	if cs.open( "page content", "ccguid=" & cp.db.encodeSqlText( pageNotFoundGuid )) then
		recordId = cs.getInteger( "id" )
		call cp.site.setProperty( "PAGENOTFOUNDPAGEID", recordId)
	end if
	Call cs.close()
	'
	' populate
	'
	siteKey = cp.site.getProperty( "siteKey" )
	link = "http://support.contensive.com/getSiteConfig?siteKey=" & siteKey
	onInstall = onInstall & "<br>request site app, link=" & link
	dim nvPairList
	set http = createObject( "kmaHttp3.httpClass" )
	nvPairList = http.getUrl( cStr( link ))
	set http = nothing
	if nvPairList <> "" then
		nvPairs = split( nvPairList, vbcrlf )
		cnt = ubound( nvPairs )+1
		onInstall = onInstall & "<br>setup default login."
		userEmail = getNameValue( nvPairs, cnt, "newServerSiteEmail" )
		set cs = cp.csNew()
		if cs.open( "people", "email=" & cp.db.encodeSqlText( userEmail ) ) Then
			onInstall = onInstall & "<br>user already created. no change."
		else
			call cs.close()
			call cs.insert( "members" )
			call cs.setfield( "name", getNameValue( nvPairs, cnt, "newServerSiteEmail" ))
			call cs.setfield( "email", getNameValue( nvPairs, cnt, "newServerSiteEmail" ))
			call cs.setfield( "username", getNameValue( nvPairs, cnt, "newServerSiteUsername" ))
			call cs.setfield( "password", getNameValue( nvPairs, cnt, "newServerSitePassword" ))
			call cs.setfield( "admin", "1" )
			m = "user added."
		end if
		call cs.close()
		'
		' populate site from application
		'
		onInstall = onInstall & "<br>setup agent."
		agentFirstName = getNameValue( nvPairs, cnt, "realEstateFirstName" )
		agentLastName = getNameValue( nvPairs, cnt, "realEstateLastName" )
		call cp.site.setproperty( "agentName", agentFirstName & " " & agentLastName )
		call cp.site.setproperty( "agentEmail", getNameValue( nvPairs, cnt, "realEstateEmail" ))
		call cp.site.setproperty( "agentNotificationEmail", getNameValue( nvPairs, cnt, "realEstateEmail" ))
		call cp.site.setproperty( "agentCodeList", getNameValue( nvPairs, cnt, "realEstateAgentId" ))
		call cp.site.setproperty( "officeCodeList", getNameValue( nvPairs, cnt, "realEstateOfficeId" ))
		call cp.site.setproperty( "agentWeb", getNameValue( nvPairs, cnt, "realEstateWWW" ))
		call cp.site.setproperty( "agentOfficePhone", getNameValue( nvPairs, cnt, "realEstateOfficePhone" ))
		call cp.site.setproperty( "agentofficename", getNameValue( nvPairs, cnt, "realEstateOfficeName" ))
		call cp.site.setproperty( "agentOfficeAddress1", getNameValue( nvPairs, cnt, "realEstateOfficeAddress1" ))
		call cp.site.setproperty( "agentOfficeAddress2", getNameValue( nvPairs, cnt, "realEstateOfficeAddress2" ))
		call cp.site.setproperty( "agentPhone", getNameValue( nvPairs, cnt, "realEstateCellPhone" ))
		'
		' copy agent photo from the signup site to this site
		'
		set http = createobject( "kmaHTTP.HTTPClass" )
		sourceUrl = "http://support.contensive.com/kmaintranet/files/upload/agentSample.jpg"
		imageFilename = "agentSample.jpg"
		agentPhotoLink = "goMethod/" & imageFilename
		http.Timeout = 60
		call http.GetURLToFile( cstr( sourceUrl ) , cp.site.physicalFilePath & "goMethod/" & imageFilename )
		call cs.insert( "library files" )
		call cs.setField( "name", "Agent Photo" )
		call cs.setField( "filename", agentPhotoLink )
		call cs.close()
		'
		' copy office photo from the signup site to this site
		'
		sourceUrl = "http://support.contensive.com/kmaintranet/files/upload/officeSample.jpg"
		imageFilename = "officeSample.jpg"
		officeLogoLink = "goMethod/" & imageFilename
		http.Timeout = 60
		call http.GetURLToFile( cstr( sourceUrl ) , cp.site.physicalFilePath & officeLogoLink )
		call cs.insert( "library files" )
		call cs.setField( "name", "Agent Office Logo" )
		call cs.setField( "filename", officeLogoLink )
		call cs.close()
		'
		' create business card copy
		'
		agentCopy = ""
		copy = cp.site.getProperty( "agentphone" )
		If copy <> "" Then
			agentCopy = agentCopy & "<br><nobr>(c) " & copy & "</nobr>"
		End If
		copy = cp.site.getProperty( "AgentOfficePhone" )
		If copy <> "" Then
			agentCopy = agentCopy & "<br><nobr>(p) " & copy & "</nobr>"
		End If
		copy = cp.site.getProperty( "AgentOfficeFax" )
		If copy <> "" Then
			agentCopy = agentCopy & "<br><nobr>(f) " & copy & "</nobr>"
		End If
		copy = cp.site.getProperty( "AgentWeb" )
		If copy = "" Then
			copy = cp.site.domainPrimary
		End If
		If InStr(1, copy, "://") = 0 Then
			link = "http://" & copy
		Else
			link = copy
		End If
		agentCopy = agentCopy & "<br><nobr><a href=""" & link & """>" & copy & "</a></nobr>"
		copy = cp.site.getProperty( "agentemail" )
		If copy <> "" Then
			agentCopy = agentCopy & "<br><nobr><a href=""mailto:" & copy & """>" & copy & "</a></nobr>"
		End If
		'
		officeCopy = ""
		copy = cp.site.getProperty( "AgentOfficePhone" )
		If copy <> "" Then
			officeCopy = officeCopy & "<br><nobr>(p) " & copy & "</nobr>"
		End If
		copy = cp.site.getProperty( "AgentOfficeFax" )
		If copy <> "" Then
			officeCopy = officeCopy & "<br><nobr>(f) " & copy & "</nobr>"
		End If
		'
		' create horizontal business card text box
		'
		onInstall = onInstall & "<br>create horizontal business card text box"
		layout = getLayout( "AgentHCard Default Layout" )
		layout = replace( layout, "$agentname$", cp.site.getproperty( "agentname" ) )
		layout = replace( layout, "$agentcopy$", agentcopy )
		layout = replace( layout, "$officename$", cp.site.getproperty( "agentofficename" ) )
		layout = replace( layout, "$officecopy$", officecopy )
		layout = replace( layout, "$agentphotolink$", cp.site.filePath & agentPhotoLink )
		layout = replace( layout, "$officelogolink$", cp.site.filePath & officeLogoLink )
		if not cs.open( "copy content", "name=" & cp.db.encodeSqlText( "AgentHCard" ), "id") then
			call cs.close()
			call cs.insert( "copy content" )
			call cs.setField( "name", "AgentHCard" )
		end if
		call cs.setField( "ccGuid", "{9D203002-747A-459A-BB44-1742203D0447}" )
		call cs.setField( "copy", layout )
		do
			call cs.goNext()
			if cs.ok() then cs.delete()
		loop while cs.ok()
		call cs.close()
		'
		' create vertical business card text box
		'
		onInstall = onInstall & "<br>create vertical business card text box"
		layout = getLayout( "AgentVCard Default Layout" )
		layout = replace( layout, "$agentname$", cp.site.getproperty( "agentname" ) )
		layout = replace( layout, "$agentcopy$", agentcopy )
		layout = replace( layout, "$officename$", cp.site.getproperty( "agentofficename" ) )
		layout = replace( layout, "$officecopy$", officecopy )
		layout = replace( layout, "$agentphotolink$", cp.site.filePath & agentPhotoLink )
		layout = replace( layout, "$officelogolink$", cp.site.filePath & officeLogoLink )
		if not cs.open( "copy content", "name=" & cp.db.encodeSqlText( "AgentVCard" ), "id") then
			call cs.close()
			call cs.insert( "copy content" )
			call cs.setField( "name", "AgentVCard" )
		end if
		call cs.setField( "ccGuid", "{AAE76682-4243-4116-A77F-CF007CC2FE55}" )
		call cs.setField( "copy", layout )
		do
			call cs.goNext()
			if cs.ok() then cs.delete()
		loop while cs.ok()
		call cs.close()
	end if
	'
	' clear cache
	'
	call clearCache()
	'
end function
'
'
'
Sub verifyLinkAlias(PageName, LinkAlias)
    Dim cs
    Dim IsFound
    Dim pageId
    '
	set cs = cp.csNew()
    IsFound = cs.open("Link Aliases", "name=" & cp.db.encodeSQLText(LinkAlias))
    call cs.close()
    '
    If Not IsFound Then
        if ( cs.insert("Link Aliases")) then
            pageId = cp.content.getRecordID("page content", PageName)
            call cs.setField( "name", LinkAlias)
            call cs.setField( "pageid", pageId)
        End If
        Call cs.close()
    End If
End Sub
'
'
'
function getLayout(LayoutName)
    Dim cs
    '
	getLayout = ""
	set cs = cp.csNew()
    if cs.open("layouts", "name=" & cp.db.encodeSQLText(LayoutName)) then
		getLayout = cs.getText( "layout" )
	end if
    call cs.close()
    '
end function
'
'
'
sub clearCache()
	dim cs
	dim fs
	'
	cp.cache.clear("")
	'
	set fs = createobject( "kmafilesystem3.filesystemclass" )
	'
	Call ccLib.ClearPageContentCache
	Call ccLib.ClearPageTemplateCache
	Call ccLib.ClearSiteSectionCache
	CS = ccLib.OpenCSContent("Content", , , , , , "name")
	Do While ccLib.IsCSOK(CS)
		Call ccLib.ClearBake(ccLib.GetCSText(CS, "name"))
		Call ccLib.NextCSRecord(CS)
	Loop
	Call ccLib.CloseCS(CS)
	Call ccLib.ExecuteSQL("default", "update ccpagecontent set childpagesfound=1")
	Call fs.DeleteFileFolder(ccLib.PhysicalFilePath & "AppCache")
	Call fs.CreateFileFolder(ccLib.PhysicalFilePath & "AppCache")
end sub
'
'
'
function getNameValue( nvPairs, cnt, Name )
	dim ptr
	dim nameLcase
	dim nvPair
	dim pairValue
	dim pairName
	dim wasFound
	'
	getNameValue = ""
	wasFound = false
	nameLcase = lcase( name )
	ptr = 0
	do
		pairName = nvPairs( ptr )
		pairValue = ""
		if instr( 1, pairName, "=" ) then
			nvPair = split( pairName, "=" )
			pairName = nvPair(0)
			pairValue= nvPair(1)
		end if
		if ( lcase( trim( pairName)) = nameLcase ) then
			getNameValue = pairValue
			wasFound = true
		end if
		ptr = ptr + 1
	loop while (not wasFound) and ( ptr < cnt )
end function
