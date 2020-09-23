<div align="center">

## Find Server Uptime


</div>

### Description

This will find how long the computer has been running. I have tried it on three different computers, each one accurately displays the uptime. Please vote and leave comments. Thank you
 
### More Info
 
Uses FileSystemObject and reads the DateLastModified atribune from the pagefile.sys.

Returns Uptime


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rob R](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-r.md)
**Level**          |Intermediate
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rob-r-find-server-uptime__4-7902/archive/master.zip)





### Source Code

```
<%
 Dim PageFileModDate, UptimeInSeconds, TotalUptime
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set PageFile = fso.GetFile ("C:\pagefile.sys")
 'get the page
 PageFileModDate = PageFile.DateLastModified
 'Finds how many seconds between now and the pagefile mod date
 UptimeInSeconds = DateDiff("S", PageFileModDate, Now())
 'Calls a function to format the seconds
 FormatSeconds (UptimeInSeconds)
 'writes the Uptime to the browser
 Response.write "Total uptime:<br><br>"
 Response.write TotalUptime
 'format function
 Function FormatSeconds(TotalSeconds)
 Seconds = TotalSeconds Mod 60
 Minutes = TotalSeconds \ 60 Mod 60
 Hours = TotalSeconds \ 3600 Mod 24
 Days = TotalSeconds \ 3600 \ 24
 TotalUptime = Days & " days " & Hours & " hours " & Minutes & " minutes " & Seconds & " seconds"
 End Function
 'dispose of the objects
 Set PageFile = Nothing
 Set fso = Nothing
%>
```

