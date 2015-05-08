<pre>&lt;cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") /&gt;
&lt;cfdump var="#iisObj.create(sitePath:'/', vdirName:'test', vdirPath:'c:\')#"&gt;
&lt;cfinclude template="list.cfm"&gt;</pre>
<cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") />
<cfdump var="#iisObj.create(sitePath:'/', vdirName:'test', vdirPath:'c:\')#">
<cfinclude template="list.cfm">