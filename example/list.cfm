<pre>&lt;cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") /&gt;
&lt;cfdump var="#variables.iisObj.list(sitePath:'/')#"&gt;</pre>
<cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") />
<cfdump var="#iisObj.list(sitePath:'/')#">