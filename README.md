<div align="center">

## Detect if Cookies are Enabled with 1 \(that's one\!\) line of code


</div>

### Description

In looking for an easy way to detect the client browser's cookie settings, everything I've found has been complex. Each example I've found requires a page refresh or redirect to itself. This severely hampers true application development as the Request.Form collection is then cleared. This code is something simple I stumbled on and it works (see explanations).
 
### More Info
 
I have NOT tested this on ALL platforms or with all clients, however, since the SITESERVER/ID cookie key links the client to a session established on the server, I'm fairly certain it will work for all cases. Please don't shoot me if it doesn't. :) I DO know that it works on Microsoft IIS and with IE v5 or better. I included in my example a redirect to an error page, however you could do absolutely anything you wanted if the test was true.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Allen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-allen.md)
**Level**          |Beginner
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-allen-detect-if-cookies-are-enabled-with-1-that-s-one-line-of-code__4-6407/archive/master.zip)





### Source Code

```

  <%@ Language=VBScript %>
  <%option explicit
  If Len(Request.Cookies("SITESERVER")("ID")) = 0 Then Response.Redirect "BrowserError.asp"
  %>
```

