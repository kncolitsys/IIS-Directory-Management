<pre>&lt;cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") /&gt;
&lt;cfdump var="#iisObj.edit(sitePath:'/', oldVdirName:'test', newVdirName:'testRenamed', newVdirPath:'c:\')#"&gt;
&lt;cfinclude template="list.cfm"&gt;</pre>
<cfset iisObj = createObject("component", "com.iisvdir").init(siteName:"#application.site#") />
<cfdump var="#iisObj.edit(sitePath:'/', oldVdirName:'test', newVdirName:'testRenamed', newVdirPath:'c:\')#">
<cfinclude template="list.cfm">