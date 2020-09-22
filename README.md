<div align="center">

## Fax with Win2k & XP


</div>

### Description

With this snippet you can fax from any windows 2000 and windows XP box with Fax Services! The only other way to share a fax otherwise is Small Business server. All feedback is welcome!
 
### More Info
 
Make a reference to Fax COM Type lib 1.0


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Buckner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-buckner.md)
**Level**          |Intermediate
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-buckner-fax-with-win2k-xp__1-28215/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  On Error GoTo ErrHandler
  Dim strComputer As String
  strComputer = "yourComputerName"
  Dim oFaxServer As FAXCOMLib.FaxServer
  Set oFaxServer = New FAXCOMLib.FaxServer
  Dim oFaxDoc As FAXCOMLib.FaxDoc
  oFaxServer.Connect strComputer
  oFaxServer.ServerCoverpage = 0
  Set oFaxDoc = oFaxServer.CreateDocument(App.Path & "\" & "New Text Document.txt")
  With oFaxDoc
    .FaxNumber = "5551212"
    .DisplayName = "Fax Server"
    Dim lngSend As Long
    lngSend = .Send
  End With
  Set oFaxDoc = Nothing
  oFaxServer.Disconnect
  Set oFaxServer = Nothing
  Exit Sub
ErrHandler:
  MsgBox Err.Number & " " & Err.Description
End Sub
```

