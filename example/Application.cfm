<cfapplication name="iisvdmExamples" clientmanagement="no" sessionmanagement="no">

<cfif not structKeyExists(application, "site")>
	<cfset application.site = "iisvdm">
</cfif>