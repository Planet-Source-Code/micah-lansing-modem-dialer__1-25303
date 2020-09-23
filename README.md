<div align="center">

## Modem Dialer


</div>

### Description

This code shows how simple it is to access your modem and dial phone #s using tone or pulse dialing. This code requires the Microsoft Comm Control. Please vote and give me plenty of feedback.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Micah Lansing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/micah-lansing.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/micah-lansing-modem-dialer__1-25303/archive/master.zip)





### Source Code

```
Dim A As Integer
Dim Instring as String
Private Sub Dialcmd_Click()
On Error GoTo pe
If A = 0 Then
 MSComm1.CommPort = 3
 MSComm1.Settings = "9600,N,8,1" ' 9600 baud, no parity, 8 data, 1 stop bit
 MSComm1.InputLen = 0 'Sets to read all buffer when input is used
 MSComm1.PortOpen = True
 A = 1
 MSComm1.Output = "AT" + Chr$(13) ' Sends "attention" command to the modem
 Do
 DoEvents
 Loop Until MSComm1.InBufferCount >= 2 'Waits for "OK"
 Instring = MSComm1.Input 'The "OK": Instring should = "AT|||OK|"
 MSComm1.Output = ATDT & PhoneNumberHere & Chr(13) 'Dials phone #, ATDT(tone) or ATDP(pulse)
End If
GoTo 2
pe:
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
2
End Sub
Private Sub Hangup_Click()
If A = 1 Then MSComm1.PortOpen = False: A = 0 'closes port
End Sub
```

