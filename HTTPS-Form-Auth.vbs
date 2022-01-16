Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

' This CreateObject statement uses the new single-DLL ActiveX for v9.5.0
set http = CreateObject("Chilkat_9_5_0.Http")

'  Any string unlocks the component for the 1st 30-days.
success = http.UnlockComponent("Anything for 30-day trial")
If (success <> 1) Then
    outFile.WriteLine(http.LastErrorText)
    WScript.Quit
End If

'  Let's begin by building an HTTP request to mimic the form.
'  We must add the parameters, and set the path.
' This CreateObject statement uses the new single-DLL ActiveX for v9.5.0
set req = CreateObject("Chilkat_9_5_0.HttpRequest")
req.AddParam "username","mylogin"
req.AddParam "password","mypassword"
req.AddParam "redirectto","/web/demo.nsf/pgWelcome?Open"

'  The path part of the POST URL is obtained from the "action" attribute of the HTML form tag.
req.Path = "/auth.nsf?Login"

req.HttpVerb = "POST"
http.FollowRedirects = 1

'  Collect cookies in-memory and re-send in subsequent HTTP requests, including any redirects.
http.SendCookies = 1
http.SaveCookies = 1
http.CookieDir = "memory"

' resp is a Chilkat_9_5_0.HttpResponse
Set resp = http.SynchronousRequest("www.something123.com",443,1,req)
If (resp Is Nothing ) Then
    outFile.WriteLine(http.LastErrorText)
    WScript.Quit
End If

'  The HTTP response object can be examined.
'  To get the HTML of the response, examine the BodyStr property (assuming the POST returns HTML)
strHtml = resp.BodyStr


outFile.Close