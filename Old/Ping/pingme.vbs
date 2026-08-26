Set Ping=CreateObject("DSPing.Ping")
Result=Ping.DoPing("cbdxaaa")
If Result = 0 then
    WScript.Echo "Ping successful."
    WScript.Echo "Approximate round trip times in milliseconds:"
    WScript.Echo "Minimum =" & Ping.Minimum & "ms"
    WScript.Echo "Maximum =" & Ping.Maximum & "ms"
    WScript.Echo "Average =" & Ping.Average & "ms"
Else
    cscript.echo "Ping failed."
End If