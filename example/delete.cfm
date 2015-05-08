<pre>&lt;cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") /&gt;
&lt;cfdump var="#iisObj.delete(sitePath:'/', vdirName:'testRenamed')#"&gt;
&lt;cfinclude template="list.cfm"&gt;</pre>
<cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") />
<cfdump var="#iisObj.delete(sitePath:'/', vdirName:'testRenamed')#">
<cfinclude template="list.cfm">