Set objArgs = WScript.Arguments
Set objIE = CreateObject("InternetExplorer.Application")

objIE.Visible = True

If objArgs.Count > 0 Then
    objIE.Navigate(objArgs(0))
End If
