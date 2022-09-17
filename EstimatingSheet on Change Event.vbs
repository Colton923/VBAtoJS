Private Sub Worksheet_Change(ByVal Target As Range)
Dim BayStart As Range
Dim BayEnd As Range
Dim Downspouts As Range
Dim PDoorStart As Range
Dim PDoorEnd As Range
Dim OHDoorStart As Range
Dim OHDoorEnd As Range
Dim WindowStart As Range
Dim WindowEnd As Range
Dim FOStart As Range
Dim FOEnd As Range
Dim OverhangTbl As Range
Dim ExtensionTbl As Range
Dim Overhangs As Range
Dim Extensions As Range
Dim cell As Range
Dim WallAvailability(1 To 4) As Boolean
Dim i As Integer

If Target.Count > 1 Then
    Call FullSheetCheck

























'''''''''''''''''''''''''''''''''''''''''''''' Sub for checking everything if a paste has happened
Private Sub FullSheetCheck()
Dim BayStart As Range
Dim BayEnd As Range
Dim Downspouts As Range
Dim PDoorStart As Range
Dim PDoorEnd As Range
Dim OHDoorStart As Range
Dim OHDoorEnd As Range
Dim WindowStart As Range
Dim WindowEnd As Range
Dim FOStart As Range
Dim FOEnd As Range
Dim OverhangTbl As Range
Dim ExtensionTbl As Range
Dim Overhangs As Range
Dim Extensions As Range
Dim cell As Range
Dim WallAvailability(1 To 4) As Boolean


'ranges
Set BayStart = EstSht.Range("Building_Height").offset(2, -1)
Set BayEnd = BayStart.offset(12, 0)
Set PDoorStart = EstSht.Range("pDoorCell1").offset(-1, 0)
Set PDoorEnd = EstSht.Range("pDoorCell12")
Set OHDoorStart = EstSht.Range("OHDoorCell1").offset(-1, 0)
Set OHDoorEnd = EstSht.Range("OHDoorCell12")
Set WindowStart = EstSht.Range("WindowCell1").offset(-1, 0)
Set WindowEnd = EstSht.Range("WindowCell12")
Set FOStart = EstSht.Range("MiscFOCell1").offset(-1, 0)
Set FOEnd = EstSht.Range("MiscFOCell12")
Set OverhangTbl = EstSht.Range("e1_GableOverhang").offset(-1, -1).Resize(5, 7)
Set ExtensionTbl = EstSht.Range("e1_GableExtension").offset(-1, -1).Resize(5, 7)
Set Overhangs = Range(EstSht.Range("e1_GableOverhang"), EstSht.Range("s4_EaveOverhang"))
Set Extensions = Range(EstSht.Range("e1_GableExtension"), EstSht.Range("s4_EaveExtension"))


''Wall Availability for Liners, Wainscot, FOs
'Assume all available, change if not
WallAvailability(1) = True
WallAvailability(2) = True
WallAvailability(3) = True
WallAvailability(4) = True
If Me.Range("e1_WallStatus").Value <> "Include" Then WallAvailability(1) = False
If Me.Range("s2_WallStatus").Value <> "Include" Then WallAvailability(2) = False
If Me.Range("e3_WallStatus").Value <> "Include" Then WallAvailability(3) = False
If Me.Range("s4_WallStatus").Value <> "Include" Then WallAvailability(4) = False

UpdatesEventsProtection (False)



















'''' bay number change
'unprotect
'hide row under bay number box when bay number is 0
If Me.Range("BayNum").Value = 0 Then
    If BayStart.offset(-1, 0).EntireRow.Hidden = False Then BayStart.offset(-1, 0).EntireRow.Hidden = True
Else
    If BayStart.offset(-1, 0).EntireRow.Hidden = True Then BayStart.offset(-1, 0).EntireRow.Hidden = False
End If
'change all bay lengths to 0
EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0
'check  value
Select Case Me.Range("BayNum").Value
Case ""
    Me.Range("BayNum").Value = "0"
Case "0"
   If EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = True
Case "1"
    If EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
    BayStart.Resize(2, 1).EntireRow.Hidden = False
Case "2"
    If EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
    BayStart.Resize(3, 1).EntireRow.Hidden = False
Case "3"
    If EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
    BayStart.Resize(4, 1).EntireRow.Hidden = False
Case "4"
    If EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
    BayStart.Resize(5, 1).EntireRow.Hidden = False
Case "5"
    If EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
    BayStart.Resize(6, 1).EntireRow.Hidden = False
Case "6"
    If EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
    BayStart.Resize(7, 1).EntireRow.Hidden = False
Case "7"
    If EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
    BayStart.Resize(8, 1).EntireRow.Hidden = False
Case "8"
    If EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
    BayStart.Resize(9, 1).EntireRow.Hidden = False
Case "9"
    If EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
    BayStart.Resize(10, 1).EntireRow.Hidden = False
Case "10"
    If EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
    BayStart.Resize(11, 1).EntireRow.Hidden = False
Case "11"
    If BayEnd.EntireRow.Hidden = False Then BayEnd.EntireRow.Hidden = True
    BayStart.Resize(12, 1).EntireRow.Hidden = False
Case "12"
   EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = False
End Select


 ''''' Alter Walls
'check if yes or no
Select Case EstSht.Range("AlterWalls").Value
Case ""
    EstSht.Range("AlterWalls").Value = "No"
Case "No"
    'If wainscot table isn't visible, hide column
    If EstSht.Range("Wainscot").Value <> "Yes" Then
        'hide column k, format J
        If EstSht.Columns("K:K").Hidden = False Then EstSht.Columns("K:K").Hidden = True
        EstSht.Columns("J:J").ColumnWidth = 5
    'remove row seperating alter walls and wainscot table
    ElseIf EstSht.Range("Wainscot").Value = "Yes" Then
        If EstSht.Range("LinerPanels").Value = "No" Then
            'resize section heading row
            EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
            EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
        End If
    End If
    'unhide row above wainscot table if needed
    If EstSht.Range("Wainscot").Value = "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
    EstSht.Range("Wainscot").offset(-3, 0).EntireRow.Hidden = False

    'Set Defaults
    Range(Me.Range("e1_WallStatus"), Me.Range("s4_WallStatus")).Value = "Include"
    Range(Me.Range("e1_WallStatus"), Me.Range("s4_WallStatus")).offset(0, 2).Value = 0
    Me.Range("e1_Expandable").Value = "No"
    Me.Range("e3_Expandable").Value = "No"
    WallAvailability(1) = True
    WallAvailability(2) = True
    WallAvailability(3) = True
    WallAvailability(4) = True

Case "Yes"
    '''''''''''''''''''''''''''''' Sheet Formatting
    'unhide column k, format J
    If EstSht.Columns("K:K").Hidden = True Then EstSht.Columns("K:K").Hidden = False
    EstSht.Columns("J:J").ColumnWidth = 30
    'unhide last table row
    If EstSht.Range("Wainscot").Value = "Yes" Then
        If EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = True Then _
        EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = False
    End If
    ''''''''''''''''''''''''''''''' Set Defaults
    With EstSht
        .Range("e1_WallStatus").Value = "Include"
        .Range("s2_WallStatus").Value = "Include"
        .Range("e3_WallStatus").Value = "Include"
        .Range("s4_WallStatus").Value = "Include"
        .Range("e1_Expandable").Value = "No"
        .Range("e3_Expandable").Value = "No"
    End With
End Select


''''' Wall Status Changes
With Me
    If .Range("e1_WallStatus").Value = "Partial" Then
        .Range("e1_WallStatus").offset(0, 2).Locked = False
        .Range("e1_WallStatus").offset(0, 2).Value = 0
    Else
        If .Range("e1_WallStatus").Value = "" Then
            .Range("e1_WallStatus").Value = "Include"
            WallAvailability(1) = True
        End If
        .Range("e1_WallStatus").offset(0, 2).Locked = True
        .Range("e1_WallStatus").offset(0, 2).Value = "N/A"
    End If
    If .Range("s2_WallStatus").Value = "Partial" Then
        .Range("s2_WallStatus").offset(0, 2).Locked = False
        .Range("s2_WallStatus").offset(0, 2).Value = 0
    Else
        If .Range("s2_WallStatus").Value = "" Then
            .Range("s2_WallStatus").Value = "Include"
            WallAvailability(2) = True
        End If
        .Range("s2_WallStatus").offset(0, 2).Locked = True
        .Range("s2_WallStatus").offset(0, 2).Value = "N/A"
    End If
    If .Range("e3_WallStatus").Value = "Partial" Then
        .Range("e3_WallStatus").offset(0, 2).Locked = False
        .Range("e3_WallStatus").offset(0, 2).Value = 0
    Else
        If .Range("e3_WallStatus").Value = "" Then
            .Range("e3_WallStatus").Value = "Include"
            WallAvailability(3) = True
        End If
        .Range("e3_WallStatus").offset(0, 2).Locked = True
        .Range("e3_WallStatus").offset(0, 2).Value = "N/A"
    End If
    If .Range("s4_WallStatus").Value = "Partial" Then
        .Range("s4_WallStatus").offset(0, 2).Locked = False
        .Range("s4_WallStatus").offset(0, 2).Value = 0
    Else
        If .Range("s4_WallStatus").Value = "" Then
            .Range("s4_WallStatus").Value = "Include"
            WallAvailability(4) = True
        End If
        .Range("s4_WallStatus").offset(0, 2).Locked = True
        .Range("s4_WallStatus").offset(0, 2).Value = "N/A"
    End If
End With

   ''''' Liner Panels Section
'check if yes or no
Select Case Me.Range("LinerPanels").Value
Case ""
    EstSht.Range("LinerPanels").Value = "No"
Case "No"
    'resize section heading row
    EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
    'hide liner panels section
    If Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False Then _
    Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True
    'unhide row above wainscot table if needed
    If EstSht.Range("Wainscot").Value = "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
    EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = False

    'clear table
    Range(Me.Range("e1_LinerPanels"), Me.Range("Roof_LinerPanels")).Value = "None"
    Range(Me.Range("e1_LinerPanels"), Me.Range("Roof_LinerPanels")).offset(0, 1).Resize(4, 4).Value = ""
Case "Yes"
    'unhide liner panels section
    If Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True Then _
    Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False
    If Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True Then _
    Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False
    'resize section heading row
    EstSht.Range("LinerPanels").offset(2, 0).EntireRow.AutoFit
End Select


    ''''' Liner Panels Options Change
For Each cell In Range(EstSht.Range("e1_LinerPanels"), EstSht.Range("Roof_LinerPanels"))
If cell.Value = "" Then cell.Value = "None"
If cell.Value = "None" Then
    cell.offset(0, 1).Value = ""
    cell.offset(0, 2).Value = ""
    cell.offset(0, 3).Value = ""
End If
Next cell


        ''''' Wainscot Section
'check if yes or no
Select Case EstSht.Range("Wainscot").Value
Case ""
    EstSht.Range("Wainscot").Value = "No"
Case "No"
    'If wainscot table isn't visible, hide column
    If EstSht.Range("AlterWalls").Value <> "Yes" Then
        'hide column k, format J
        If EstSht.Columns("K:K").Hidden = False Then EstSht.Columns("K:K").Hidden = True
        EstSht.Columns("J:J").ColumnWidth = 5
        If EstSht.Range("LinerPanels").Value = "No" Then
            'resize section heading row
            EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
            EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
        End If
    ElseIf EstSht.Range("AlterWalls").Value = "Yes" Then
        'hide row seperating walls and wainscot table
        If EstSht.Range("LinerPanels").Value = "No" Then
            'resize section heading row
            EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
            EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
        End If
    End If
    ''''''''''''''''''''''''''''''' reset rows
    With EstSht
        .Range("e1_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("s2_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("e3_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("s4_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        Range(.Range("e1_Wainscot"), .Range("s4_Wainscot")).Value = "None"
    End With

Case "Yes"
    'unhide column k, format J
    If EstSht.Columns("K:K").Hidden = True Then EstSht.Columns("K:K").Hidden = False
    EstSht.Columns("J:J").ColumnWidth = 30
'        'unhide row above table if needed
    If EstSht.Range("LinerPanels").Value <> "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
    EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = False
    ''''''''''''''''''''''''''''''' Set Defaults
    With EstSht
        .Range("e1_Wainscot").Value = "None"
        .Range("s2_Wainscot").Value = "None"
        .Range("e3_Wainscot").Value = "None"
        .Range("s4_Wainscot").Value = "None"
        .Range("e1_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("s2_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("e3_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
        .Range("s4_Wainscot").offset(0, 1).Resize(1, 2).Value = ""
    End With
End Select


        ''''' Wainscot Options Change
    For Each cell In Range(EstSht.Range("e1_Wainscot"), EstSht.Range("s4_Wainscot"))
    If cell.Value = "" Then cell.Value = "None"
    If cell.Value = "None" Then
        cell.offset(0, 1).Value = ""
        cell.offset(0, 2).Value = ""
    End If
    Next cell

    '''' Gutter & Downspouts
    'check if yes or no
    If EstSht.Range("GutterAndDownspouts").Value = "" Then EstSht.Range("GutterAndDownspouts").Value = "No"

    '''' personnel door number
     'hide row under quantity box when quantity is 0
    If EstSht.Range("PDoorNum").Value = 0 Or EstSht.Range("PDoorNum").Value = "" Then
        EstSht.Range("PDoorNum").Value = 0
        'reset table to blank
        Me.Range("pDoorCell1").offset(0, 1).Resize(12, 6).Value = ""
        'hide
        If PDoorStart.offset(-1, 0).EntireRow.Hidden = False Then PDoorStart.offset(-1, 0).EntireRow.Hidden = True
    Else
        If PDoorStart.offset(-1, 0).EntireRow.Hidden = True Then PDoorStart.offset(-1, 0).EntireRow.Hidden = False
    End If
    'check target value
    Select Case EstSht.Range("PDoorNum").Value
    Case "0"
       If EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden = True
    Case "1"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
        PDoorStart.Resize(2, 1).EntireRow.Hidden = False
    Case "2"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
        PDoorStart.Resize(3, 1).EntireRow.Hidden = False
    Case "3"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
        PDoorStart.Resize(4, 1).EntireRow.Hidden = False
    Case "4"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
        PDoorStart.Resize(5, 1).EntireRow.Hidden = False
    Case "5"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
        PDoorStart.Resize(6, 1).EntireRow.Hidden = False
    Case "6"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
        PDoorStart.Resize(7, 1).EntireRow.Hidden = False
    Case "7"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
        PDoorStart.Resize(8, 1).EntireRow.Hidden = False
    Case "8"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
        PDoorStart.Resize(9, 1).EntireRow.Hidden = False
    Case "9"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
        PDoorStart.Resize(10, 1).EntireRow.Hidden = False
    Case "10"
        If EstSht.Range(PDoorStart, PDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(PDoorStart, PDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
        PDoorStart.Resize(11, 1).EntireRow.Hidden = False
    Case "11"
        If PDoorEnd.EntireRow.Hidden = False Then PDoorEnd.EntireRow.Hidden = True
        PDoorStart.Resize(12, 1).EntireRow.Hidden = False
    Case "12"
       EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden = False
    End Select

    '''' OH door
    'hide row under quantity box when quantity is 0
    If EstSht.Range("OHDoorNum").Value = 0 Or EstSht.Range("OHDoorNum").Value = "" Then
        EstSht.Range("OHDoorNum").Value = 0
        'reset table to blank
        Me.Range("OHDoorCell1").offset(0, 1).Resize(12, 8).Value = ""
        'hide
        If OHDoorStart.offset(-1, 0).EntireRow.Hidden = False Then OHDoorStart.offset(-1, 0).EntireRow.Hidden = True
    Else
        If OHDoorStart.offset(-1, 0).EntireRow.Hidden = True Then OHDoorStart.offset(-1, 0).EntireRow.Hidden = False
    End If
   'check target value
    Select Case EstSht.Range("OHDoorNum").Value
    Case "0"
       If EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden = True
    Case "1"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(2, 1).EntireRow.Hidden = False
    Case "2"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(3, 1).EntireRow.Hidden = False
    Case "3"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(4, 1).EntireRow.Hidden = False
    Case "4"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(5, 1).EntireRow.Hidden = False
    Case "5"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(6, 1).EntireRow.Hidden = False
    Case "6"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(7, 1).EntireRow.Hidden = False
    Case "7"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(8, 1).EntireRow.Hidden = False
    Case "8"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(9, 1).EntireRow.Hidden = False
    Case "9"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(10, 1).EntireRow.Hidden = False
    Case "10"
        If EstSht.Range(OHDoorStart, OHDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(OHDoorStart, OHDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
        OHDoorStart.Resize(11, 1).EntireRow.Hidden = False
    Case "11"
        If OHDoorEnd.EntireRow.Hidden = False Then OHDoorEnd.EntireRow.Hidden = True
        OHDoorStart.Resize(12, 1).EntireRow.Hidden = False
    Case "12"
       EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden = False
    End Select


   '''' Windows
    'hide row under quantity box when quantity is 0
    If EstSht.Range("WindowNum").Value = 0 Or EstSht.Range("WindowNum").Value = "" Then
        EstSht.Range("WindowNum").Value = 0
        'reset table to blank
        Me.Range("WindowCell1").offset(0, 1).Resize(24, 3).Value = ""
        'hide
        If WindowStart.offset(-1, 0).EntireRow.Hidden = False Then WindowStart.offset(-1, 0).EntireRow.Hidden = True
    Else
        If WindowStart.offset(-1, 0).EntireRow.Hidden = True Then WindowStart.offset(-1, 0).EntireRow.Hidden = False
    End If
    'check target value
    Select Case EstSht.Range("WindowNum").Value
    Case "0"
       If EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden = True
    Case "1"
        If EstSht.Range(WindowStart, WindowEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
        WindowStart.Resize(2, 1).EntireRow.Hidden = False
    Case "2"
        If EstSht.Range(WindowStart, WindowEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
        WindowStart.Resize(3, 1).EntireRow.Hidden = False
    Case "3"
        If EstSht.Range(WindowStart, WindowEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
        WindowStart.Resize(4, 1).EntireRow.Hidden = False
    Case "4"
        If EstSht.Range(WindowStart, WindowEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
        WindowStart.Resize(5, 1).EntireRow.Hidden = False
    Case "5"
        If EstSht.Range(WindowStart, WindowEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
        WindowStart.Resize(6, 1).EntireRow.Hidden = False
    Case "6"
        If EstSht.Range(WindowStart, WindowEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
        WindowStart.Resize(7, 1).EntireRow.Hidden = False
    Case "7"
        If EstSht.Range(WindowStart, WindowEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
        WindowStart.Resize(8, 1).EntireRow.Hidden = False
    Case "8"
        If EstSht.Range(WindowStart, WindowEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
        WindowStart.Resize(9, 1).EntireRow.Hidden = False
    Case "9"
        If EstSht.Range(WindowStart, WindowEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
        WindowStart.Resize(10, 1).EntireRow.Hidden = False
    Case "10"
        If EstSht.Range(WindowStart, WindowEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(WindowStart, WindowEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
        WindowStart.Resize(11, 1).EntireRow.Hidden = False
    Case "11"
        If WindowEnd.EntireRow.Hidden = False Then WindowEnd.EntireRow.Hidden = True
        WindowStart.Resize(12, 1).EntireRow.Hidden = False
    Case "12"
       EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden = False
    End Select


   '''' Misc Framed Openings
    'hide row under quantity box when quantity is 0
    If EstSht.Range("MiscFONum").Value = 0 Or EstSht.Range("MiscFONum").Value = "" Then
        EstSht.Range("MiscFONum").Value = 0
        'reset table to blank
        Me.Range("MiscFOCell1").offset(0, 1).Resize(12, 5).Value = ""
        'hide
        If FOStart.offset(-1, 0).EntireRow.Hidden = False Then FOStart.offset(-1, 0).EntireRow.Hidden = True
    Else
        If FOStart.offset(-1, 0).EntireRow.Hidden = True Then FOStart.offset(-1, 0).EntireRow.Hidden = False
    End If
    'check target value
    Select Case EstSht.Range("MiscFONum").Value
    Case "0"
       If EstSht.Range(FOStart, FOEnd).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).EntireRow.Hidden = True
    Case "1"
        If EstSht.Range(FOStart, FOEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
        FOStart.Resize(2, 1).EntireRow.Hidden = False
    Case "2"
        If EstSht.Range(FOStart, FOEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
        FOStart.Resize(3, 1).EntireRow.Hidden = False
    Case "3"
        If EstSht.Range(FOStart, FOEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
        FOStart.Resize(4, 1).EntireRow.Hidden = False
    Case "4"
        If EstSht.Range(FOStart, FOEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
        FOStart.Resize(5, 1).EntireRow.Hidden = False
    Case "5"
        If EstSht.Range(FOStart, FOEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
        FOStart.Resize(6, 1).EntireRow.Hidden = False
    Case "6"
        If EstSht.Range(FOStart, FOEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
        FOStart.Resize(7, 1).EntireRow.Hidden = False
    Case "7"
        If EstSht.Range(FOStart, FOEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
        FOStart.Resize(8, 1).EntireRow.Hidden = False
    Case "8"
        If EstSht.Range(FOStart, FOEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
        FOStart.Resize(9, 1).EntireRow.Hidden = False
    Case "9"
        If EstSht.Range(FOStart, FOEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
        FOStart.Resize(10, 1).EntireRow.Hidden = False
    Case "10"
        If EstSht.Range(FOStart, FOEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(FOStart, FOEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
        FOStart.Resize(11, 1).EntireRow.Hidden = False
    Case "11"
        If FOEnd.EntireRow.Hidden = False Then FOEnd.EntireRow.Hidden = True
        FOStart.Resize(12, 1).EntireRow.Hidden = False
    Case "12"
       EstSht.Range(FOStart, FOEnd).EntireRow.Hidden = False
    End Select

'''' building length change
    'if blank building length, reset to 0
    If EstSht.Range("Building_Length").Value = "" Then
        EstSht.Range("Building_Length").Value = 0
        EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0
    End If



'''' Trim Color Bulk Change
    If EstSht.Range("All_tColors").Value = "" Then
        EstSht.Range("All_tColors").Value = "N/A"
    Else
        'change trim colors to match
        With EstSht
            .Range("Rake_tColor").Value = EstSht.Range("All_tColors").Value
            .Range("Eave_tColor").Value = EstSht.Range("All_tColors").Value
            .Range("OutsideCorner_tColor").Value = EstSht.Range("All_tColors").Value
            .Range("FO_tColor").Value = EstSht.Range("All_tColors").Value
            .Range("Base_tColor").Value = EstSht.Range("All_tColors").Value
            'change gutter/downspout colors as well
            .Range("DownspoutColor").Value = EstSht.Range("All_tColors").Value
            .Range("GutterColor").Value = EstSht.Range("All_tColors").Value
        End With
    End If


'''' overhang table clear
    'clear soffits
    For Each cell In Overhangs
            If cell.Value = "" Or cell.Value = 0 Then
                'clear soffits
                cell.offset(0, 1).Value = ""
                cell.offset(0, 2).Value = ""
                cell.offset(0, 3).Value = ""
                cell.offset(0, 4).Value = ""
                cell.offset(0, 5).Value = ""
            End If
    Next cell

'extension table clear
    'clear soffits
    For Each cell In Extensions
            If cell.Value = "" Or cell.Value = 0 Then
                'clear soffits
                cell.offset(0, 1).Value = ""
                cell.offset(0, 2).Value = ""
                cell.offset(0, 3).Value = ""
                cell.offset(0, 4).Value = ""
                cell.offset(0, 5).Value = ""
            End If
    Next cell

'Show/Hide Eave Extension Pitch and Set Intersection default values
With EstSht
    's2 eave extension
        If .Range("s2_EaveExtension").Value = "" Then
            If .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False Then .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True
            .Range("s2e1_Intersection").Value = "N/A"
            .Range("s2e3_Intersection").Value = "N/A"
        Else
            'If previously hidden, unhide and set default option to include intersection
            If .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True Then
                .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False
                'reset to defaults
                .Range("s2_EaveExtensionPitch").Value = "Match Roof"
                If .Range("e1_GableExtension").Value <> "" Then .Range("s2e1_Intersection").Value = "Include"
                If .Range("e3_GableExtension").Value <> "" Then .Range("s2e3_Intersection").Value = "Include"
            End If
        End If
    's4 eave extension
        If .Range("s4_EaveExtension").Value = "" Then
            If .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False Then .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True
            .Range("s4e1_Intersection").Value = "N/A"
            .Range("s4e3_Intersection").Value = "N/A"
        Else
            'If previously hidden, unhide and set default option to include intersection
            If .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True Then
                .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False
                .Range("s4_EaveExtensionPitch").Value = "Match Roof"
                If .Range("e1_GableExtension").Value <> "" Then .Range("s4e1_Intersection").Value = "Include"
                If .Range("e3_GableExtension").Value <> "" Then .Range("s4e3_Intersection").Value = "Include"
            End If
        End If
    'e1 gable extension Intersection Option
        If .Range("e1_GableExtension").Value = "" Then
            .Range("s2e1_Intersection").Value = "N/A"
            .Range("s2e1_Intersection").MergeArea.Locked = True
            .Range("s4e1_Intersection").Value = "N/A"
            .Range("s4e1_Intersection").MergeArea.Locked = True
        Else
            'intersection for s2 e1
            If .Range("s2e1_Intersection").Value = "N/A" Then
                .Range("s2e1_Intersection").MergeArea.Locked = False
                If .Range("s2_EaveExtension").Value <> "" Then .Range("s2e1_Intersection").Value = "Include"
            End If
            'intersection for s4 e1
            If .Range("s4e1_Intersection").Value = "N/A" Then
                .Range("s4e1_Intersection").MergeArea.Locked = False
                If .Range("s4_EaveExtension").Value <> "" Then .Range("s4e1_Intersection").Value = "Include"
            End If
        End If
    'e3 gable extension Intersection Option
        If .Range("e3_GableExtension").Value = "" Then
            .Range("s2e3_Intersection").Value = "N/A"
            .Range("s2e3_Intersection").MergeArea.Locked = True
            .Range("s4e3_Intersection").Value = "N/A"
            .Range("s4e3_Intersection").MergeArea.Locked = True
        Else
            'intersection for s2 e3
            If .Range("s2e3_Intersection").Value = "N/A" Then
                .Range("s2e3_Intersection").MergeArea.Locked = False
                .Range("s2e3_Intersection").Value = "Include"
            End If
            'intersection for s4 e3
            If .Range("s4e3_Intersection").Value = "N/A" Then
                .Range("s4e3_Intersection").MergeArea.Locked = False
                .Range("s4e3_Intersection").Value = "Include"
            End If
        End If
End With


'''' Roof/Wall Panel Shape Change - Disable Translucent Wall Panels and Skylights
With EstSht
        '''' Panel Shape
        'Hide skylights/translucent wall panels for m-loc
        If .Range("Wall_pShape").Value = "M-Loc" Or .Range("Roof_pShape").Value = "M-Loc" Then
            .Range("TranslucentWallPanelQty").Value = ""
            .Range("SkylightQty").Value = ""
            .Range("TranslucentWallPanelLength").Value = ""
            .Range("SkylightLength").Value = ""
            'hide rows if needed
            If .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = False Then .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = True
        'unhide skylights/translucent wall panels for R-loc
        ElseIf .Range("Wall_pShape").Value <> "M-Loc" And .Range("Roof_pShape").Value <> "M-Loc" Then
            'unhide rows if needed
            If .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = True Then .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = False
        End If
End With

'''''''''''''''''''''''''''''''''''''''''''''''''' Framed Opening Option Changes '''''''''''''''''''''
With EstSht
    '''''''''''''''''''''''''''''''''''''''''' Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''
        For Each cell In Range(.Range("pDoorCell1").offset(0, 1), .Range("pDoorCell12").offset(0, 1))
        If cell.Value = "4070" Then
            'Remove half glass option
            With cell.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 2).Value = "No"
            'Remove dead bolt option
            With cell.offset(0, 5).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 5).Value = "No"
        ElseIf cell.Value = "3070" Or cell.Value = "" Then
            'restore half glass, deadbolt
            With cell.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 2).Validation.Value = False Then cell.offset(0, 2).Value = ""
            With cell.offset(0, 5).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 5).Validation.Value = False Then cell.offset(0, 5).Value = ""
        End If
        Next cell


     '''''''''''''''''''''''''''''''''''''''''' Overhead Doors '''''''''''''''''''''''''''''''''''''''''''''''

        For Each cell In Range(.Range("OHDoorCell1").offset(0, 4), .Range("OHDoorCell12").offset(0, 4))
''''''''''''''''''''''''' Roll Up Doors''''''''''''''''''''''
        If cell.Value = "RUD" Then
            '''sizing options
            'width
            With cell.offset(0, -3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("RUDWidth").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'clear if invalid
            If cell.offset(0, -3).Validation.Value = False Then cell.offset(0, -3).Value = ""
            'height
            With cell.offset(0, -2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("RUDHeight").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'clear if invalid
            If cell.offset(0, -2).Validation.Value = False Then cell.offset(0, -2).Value = ""
            '''remove insulation, operation, windows, and high lift options
            'insulation
            With cell.offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="None"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 1).Value = "None"
            'Operation
            With cell.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Chain Hoist"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 2).Value = "Chain Hoist"
            'High Lift
            With cell.offset(0, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 3).Value = "No"
            'Windows
            With cell.offset(0, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="None"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            cell.offset(0, 4).Value = "None"
''''''''''''''''''''''''' Sectional Doors''''''''''''''''''''''
        ElseIf cell.Value = "Sectional" Or cell.Value = "" Then
            '''sizing options
            'width
            With cell.offset(0, -3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("SectionalOHDoorWidth").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, -3).Validation.Value = False Then cell.offset(0, -3).Value = ""
            'height
            With cell.offset(0, -2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("SectionalOHDoorHeight").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, -2).Validation.Value = False Then cell.offset(0, -2).Value = ""
            'insulation
            With cell.offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorInsulationOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 1).Validation.Value = False Then cell.offset(0, 1).Value = ""
            'Operation
            With cell.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorOperationOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 2).Validation.Value = False Then cell.offset(0, 2).Value = ""
            'High Lift
            With cell.offset(0, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 3).Validation.Value = False Then cell.offset(0, 3).Value = ""
            'Windows
            With cell.offset(0, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorWindowOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If cell.offset(0, 4).Validation.Value = False Then cell.offset(0, 4).Value = ""
            'change options

        End If
        Next cell
End With


Call AlterAvailableWalls(WallAvailability)






















UpdatesEventsProtection (True)













End Sub
    Exit Sub
End If


'ranges
Set BayStart = EstSht.Range("Building_Height").offset(2, -1)
Set BayEnd = BayStart.offset(12, 0)
Set PDoorStart = EstSht.Range("pDoorCell1").offset(-1, 0)
Set PDoorEnd = EstSht.Range("pDoorCell12")
Set OHDoorStart = EstSht.Range("OHDoorCell1").offset(-1, 0)
Set OHDoorEnd = EstSht.Range("OHDoorCell12")
Set WindowStart = EstSht.Range("WindowCell1").offset(-1, 0)
Set WindowEnd = EstSht.Range("WindowCell12")
Set FOStart = EstSht.Range("MiscFOCell1").offset(-1, 0)
Set FOEnd = EstSht.Range("MiscFOCell12")
Set OverhangTbl = EstSht.Range("e1_GableOverhang").offset(-1, -1).Resize(5, 7)
Set ExtensionTbl = EstSht.Range("e1_GableExtension").offset(-1, -1).Resize(5, 7)
Set Overhangs = Range(EstSht.Range("e1_GableOverhang"), EstSht.Range("s4_EaveOverhang"))
Set Extensions = Range(EstSht.Range("e1_GableExtension"), EstSht.Range("s4_EaveExtension"))


''Wall Availability for Liners, Wainscot, FOs
'Assume all available, change if not
WallAvailability(1) = True
WallAvailability(2) = True
WallAvailability(3) = True
WallAvailability(4) = True
If Me.Range("e1_WallStatus").Value <> "Include" Then WallAvailability(1) = False
If Me.Range("s2_WallStatus").Value <> "Include" Then WallAvailability(2) = False
If Me.Range("e3_WallStatus").Value <> "Include" Then WallAvailability(3) = False
If Me.Range("s4_WallStatus").Value <> "Include" Then WallAvailability(4) = False



'''' bay number change
If Not Intersect(Target, EstSht.Range("BayNum")) Is Nothing Then
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    'hide row under bay number box when bay number is 0
    If Target.Value = 0 Then
        If BayStart.offset(-1, 0).EntireRow.Hidden = False Then BayStart.offset(-1, 0).EntireRow.Hidden = True
    Else
        If BayStart.offset(-1, 0).EntireRow.Hidden = True Then BayStart.offset(-1, 0).EntireRow.Hidden = False
    End If
    'change all bay lengths to 0
    Application.EnableEvents = False
    EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0
    Application.EnableEvents = True
    'check target value
    Select Case Target.Value
    Case ""
        Target.Value = "0"
    Case "0"
       If EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = True
    Case "1"
        If EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = True
        BayStart.Resize(2, 1).EntireRow.Hidden = False
    Case "2"
        If EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = True
        BayStart.Resize(3, 1).EntireRow.Hidden = False
    Case "3"
        If EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = True
        BayStart.Resize(4, 1).EntireRow.Hidden = False
    Case "4"
        If EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = True
        BayStart.Resize(5, 1).EntireRow.Hidden = False
    Case "5"
        If EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = True
        BayStart.Resize(6, 1).EntireRow.Hidden = False
    Case "6"
        If EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = True
        BayStart.Resize(7, 1).EntireRow.Hidden = False
    Case "7"
        If EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = True
        BayStart.Resize(8, 1).EntireRow.Hidden = False
    Case "8"
        If EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = True
        BayStart.Resize(9, 1).EntireRow.Hidden = False
    Case "9"
        If EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = True
        BayStart.Resize(10, 1).EntireRow.Hidden = False
    Case "10"
        If EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = False Then EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = True
        BayStart.Resize(11, 1).EntireRow.Hidden = False
    Case "11"
        If BayEnd.EntireRow.Hidden = False Then BayEnd.EntireRow.Hidden = True
        BayStart.Resize(12, 1).EntireRow.Hidden = False
    Case "12"
       EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = False
    End Select
    'protect
    EstSht.Protect "WhiteTruckMafia"
End If

            ''''' Alter Walls
If Not Intersect(Target, EstSht.Range("AlterWalls")) Is Nothing Then
    Application.ScreenUpdating = False
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    'check if yes or no
    Select Case Target.Value
    Case ""
        Target.Value = "No"
    Case "No"
        'If wainscot table isn't visible, hide column
        If EstSht.Range("Wainscot").Value <> "Yes" Then
            'do nothing
        'remove row seperating alter walls and wainscot table
        ElseIf EstSht.Range("Wainscot").Value = "Yes" Then
            If EstSht.Range("LinerPanels").Value = "No" Then
                'resize section heading row
                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
            End If
        End If
        'Lock Cells
        EstSht.Range("e1_WallStatus").Resize(4, 3).Locked = True
        'unhide row above wainscot table if needed
        If EstSht.Range("Wainscot").Value = "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
        EstSht.Range("Wainscot").offset(-3, 0).EntireRow.Hidden = False

        'Set Defaults
        UpdatesEventsProtection (False)
        Range(Me.Range("e1_WallStatus"), Me.Range("s4_WallStatus")).Value = "Include"
        Range(Me.Range("e1_WallStatus"), Me.Range("s4_WallStatus")).offset(0, 2).Value = 0
        Me.Range("e1_Expandable").Value = "No"
        Me.Range("e3_Expandable").Value = "No"
        WallAvailability(1) = True
        WallAvailability(2) = True
        WallAvailability(3) = True
        WallAvailability(4) = True
        Call AlterAvailableWalls(WallAvailability)
















Private Sub AlterAvailableWalls(WallAvailability() As Boolean)
Dim WallSelection As String
Dim N As Integer

UpdatesEventsProtection (False)













For N = 1 To 4
    If WallAvailability(N) = True Then
        Select Case N
        Case 1
            WallSelection = "Endwall 1"
        Case 2
            If WallSelection <> "" Then
                WallSelection = WallSelection & "," & "Sidewall 2"
            Else
                WallSelection = "Sidewall 2"
            End If
        Case 3
            If WallSelection <> "" Then
                WallSelection = WallSelection & "," & "Endwall 3"
            Else
                WallSelection = "Endwall 3"
            End If
        Case 4
            If WallSelection <> "" Then
                WallSelection = WallSelection & "," & "Sidewall 4"
            Else
                WallSelection = "Sidewall 4"
            End If
        End Select
    End If
Next N

With Me
    ''' Liner Panels, Wainscot
    If WallAvailability(1) = False Then
        If .Range("e1_WallStatus").Value = "Exclude" Then
            .Range("e1_LinerPanels").Resize(1, 4).Value = ""
            .Range("e1_LinerPanels").Resize(1, 4).Locked = True
            .Range("e1_LinerPanels").Value = "None"
        Else
            .Range("e1_LinerPanels").Resize(1, 4).Locked = False
        End If
        .Range("e1_Wainscot").Resize(1, 3).Value = ""
        .Range("e1_Wainscot").Resize(1, 3).Locked = True
        .Range("e1_Wainscot").Value = "None"
    Else
        .Range("e1_LinerPanels").Resize(1, 4).Locked = False
        .Range("e1_Wainscot").Resize(1, 3).Locked = False
    End If
    If WallAvailability(2) = False Then
        If .Range("s2_WallStatus").Value = "Exclude" Then
            .Range("s2_LinerPanels").Resize(1, 4).Value = ""
            .Range("s2_LinerPanels").Resize(1, 4).Locked = True
            .Range("s2_LinerPanels").Value = "None"
        End If
        .Range("s2_Wainscot").Resize(1, 3).Value = ""
        .Range("s2_Wainscot").Resize(1, 3).Locked = True
        .Range("s2_Wainscot").Value = "None"
    Else
        .Range("s2_LinerPanels").Resize(1, 4).Locked = False
        .Range("s2_Wainscot").Resize(1, 3).Locked = False
    End If
    If WallAvailability(3) = False Then
        If .Range("e3_WallStatus").Value = "Exclude" Then
            .Range("e3_LinerPanels").Resize(1, 4).Value = ""
            .Range("e3_LinerPanels").Resize(1, 4).Locked = True
            .Range("e3_LinerPanels").Value = "None"
        End If
        .Range("e3_Wainscot").Resize(1, 3).Value = ""
        .Range("e3_Wainscot").Resize(1, 3).Locked = True
        .Range("e3_Wainscot").Value = "None"
    Else
        .Range("e3_LinerPanels").Resize(1, 4).Locked = False
        .Range("e3_Wainscot").Resize(1, 3).Locked = False
    End If
    If WallAvailability(4) = False Then
        If .Range("s4_WallStatus").Value = "Exclude" Then
            .Range("s4_LinerPanels").Resize(1, 4).Value = ""
            .Range("s4_LinerPanels").Resize(1, 4).Locked = True
            .Range("s4_LinerPanels").Value = "None"
        End If
        .Range("s4_Wainscot").Resize(1, 3).Value = ""
        .Range("s4_Wainscot").Resize(1, 3).Locked = True
        .Range("s4_Wainscot").Value = "None"
    Else
        .Range("s4_LinerPanels").Resize(1, 4).Locked = False
        .Range("s4_Wainscot").Resize(1, 3).Locked = False
    End If
    ''' Wainscot

    'no included walls, remove PDoors and OHDoors
    If WallSelection = "" Then
        UpdatesEventsProtection (True)
        .Range("PDoorNum").Value = 0
        .Range("OHDoorNum").Value = 0
        UpdatesEventsProtection (False)
    End If
    'if all walls excluded, remove Windows and MiscFOs
    If .Range("e1_WallStatus") = "Exclude" And _
    .Range("s2_WallStatus") = "Exclude" And _
    .Range("e3_WallStatus") = "Exclude" And _
    .Range("s4_WallStatus") = "Exclude" Then
        UpdatesEventsProtection (True)
        .Range("WindowNum").Value = 0
        .Range("MiscFONum").Value = 0
        UpdatesEventsProtection (False)
    'at least one wall is included, allow PDoors and OHDoors

    ElseIf WallSelection <> "" Then
        Dim FieldLocateWallSelection As String
        FieldLocateWallSelection = WallSelection & "," & "Field Locate"
        ''''P doors
        With .Range("pDoorCell1").offset(0, 2).Resize(12, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=FieldLocateWallSelection
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        .Range("pDoorCell1").offset(0, 2).Resize(12, 1).Value = ""
        ''''OH doors
        With .Range("OHDoorCell1").offset(0, 3).Resize(12, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=WallSelection
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        .Range("OHDoorCell1").offset(0, 3).Resize(12, 1).Value = ""
    End If
    'Reset Wall Selection, allow for Partial or Included for Windows/MisFOs
    WallSelection = ""
    If .Range("e1_WallStatus").Value <> "Exclude" Then
        WallSelection = WallSelection & "," & "Endwall 1"
    End If
    If .Range("s2_WallStatus").Value <> "Exclude" Then
        WallSelection = WallSelection & "," & "Sidewall 2"
    End If
    If .Range("e3_WallStatus").Value <> "Exclude" Then
        WallSelection = WallSelection & "," & "Endwall 3"
    End If
    If .Range("s4_WallStatus").Value <> "Exclude" Then
        WallSelection = WallSelection & "," & "Sidewall 4"
    End If
    If WallSelection = "" Then
        'Do nothing
    Else
    ''''Windows
    FieldLocateWallSelection = WallSelection & "," & "Field Locate"
        With .Range("WindowCell1").offset(0, 3).Resize(24, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=FieldLocateWallSelection
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        .Range("WindowCell1").offset(0, 3).Resize(24, 1).Value = ""
        ''''Misc FOs
        With .Range("MiscFOCell1").offset(0, 3).Resize(12, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=FieldLocateWallSelection
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        .Range("MiscFOCell1").offset(0, 3).Resize(12, 1).Value = ""
    End If
End With

UpdatesEventsProtection (True)

End Sub
















        UpdatesEventsProtection (True)

    Case "Yes"
        '''''''''''''''''''''''''''''' Sheet Formatting
        'Unlock Cells
        EstSht.Range("e1_WallStatus").Resize(4, 3).Locked = False
        'unhide last table row
        If EstSht.Range("Wainscot").Value = "Yes" Then
            If EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = True Then _
            EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = False
        End If
        ''''''''''''''''''''''''''''''' Set Defaults
'        With EstSht
'            .Range("e1_WallStatus").Value = "Include"
'            .Range("s2_WallStatus").Value = "Include"
'            .Range("e3_WallStatus").Value = "Include"
'            .Range("s4_WallStatus").Value = "Include"
'            .Range("e1_Expandable").Value = "No"
'            .Range("e3_Expandable").Value = "No"
'        End With
    End Select
    'protect
    EstSht.Protect "WhiteTruckMafia"
    Application.ScreenUpdating = True
End If

''''' Wall Status Changes
With Me
    If Not Intersect(Target, Range(.Range("e1_WallStatus"), .Range("s4_WallStatus"))) Is Nothing Then
        UpdatesEventsProtection (False)










        Private Sub UpdatesEventsProtection(Setting As Boolean)
If Setting = True Then
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Me.Protect "WhiteTruckMafia"
Else
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Me.Unprotect "WhiteTruckMafia"
End If
End Sub








        Select Case Target.Address
        Case .Range("e1_WallStatus").Address
            If .Range("e1_WallStatus").Value = "Partial" Then
                .Range("e1_WallStatus").offset(0, 2).Locked = False
                .Range("e1_WallStatus").offset(0, 2).Value = 0
            Else
                If .Range("e1_WallStatus").Value = "" Then
                    .Range("e1_WallStatus").Value = "Include"
                    WallAvailability(1) = True
                End If
                .Range("e1_WallStatus").offset(0, 2).Locked = True
                .Range("e1_WallStatus").offset(0, 2).Value = "N/A"
            End If
        Case .Range("s2_WallStatus").Address
            If .Range("s2_WallStatus").Value = "Partial" Then
                .Range("s2_WallStatus").offset(0, 2).Locked = False
                .Range("s2_WallStatus").offset(0, 2).Value = 0
            Else
                If .Range("s2_WallStatus").Value = "" Then
                    .Range("s2_WallStatus").Value = "Include"
                    WallAvailability(2) = True
                End If
                .Range("s2_WallStatus").offset(0, 2).Locked = True
                .Range("s2_WallStatus").offset(0, 2).Value = "N/A"
            End If
        Case .Range("e3_WallStatus").Address
            If .Range("e3_WallStatus").Value = "Partial" Then
                .Range("e3_WallStatus").offset(0, 2).Locked = False
                .Range("e3_WallStatus").offset(0, 2).Value = 0
            Else
                If .Range("e3_WallStatus").Value = "" Then
                    .Range("e3_WallStatus").Value = "Include"
                    WallAvailability(3) = True
                End If
                .Range("e3_WallStatus").offset(0, 2).Locked = True
                .Range("e3_WallStatus").offset(0, 2).Value = "N/A"
            End If
        Case .Range("s4_WallStatus").Address
            If .Range("s4_WallStatus").Value = "Partial" Then
                .Range("s4_WallStatus").offset(0, 2).Locked = False
                .Range("s4_WallStatus").offset(0, 2).Value = 0
            Else
                If .Range("s4_WallStatus").Value = "" Then
                    .Range("s4_WallStatus").Value = "Include"
                    WallAvailability(4) = True
                End If
                .Range("s4_WallStatus").offset(0, 2).Locked = True
                .Range("s4_WallStatus").offset(0, 2).Value = "N/A"
            End If
        End Select
        'update wall availability
        Call AlterAvailableWalls(WallAvailability)






















        UpdatesEventsProtection (True)















    End If

End With

       ''''' Liner Panels Section
If Not Intersect(Target, EstSht.Range("LinerPanels")) Is Nothing Then
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    'check if yes or no
    Select Case Target.Value
    Case ""
        Target.Value = "No"
    Case "No"
        'resize section heading row
        EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
        'hide liner panels section
        If Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False Then
            Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True
        End If
        'unhide row above wainscot table if needed
        If EstSht.Range("Wainscot").Value = "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = False

        'clear table
        UpdatesEventsProtection (False)










        Range(Me.Range("e1_LinerPanels"), Me.Range("Roof_LinerPanels")).Value = "None"
        Range(Me.Range("e1_LinerPanels"), Me.Range("Roof_LinerPanels")).offset(0, 1).Resize(5, 4).Value = ""
        UpdatesEventsProtection (True)










    Case "Yes"
        'unhide liner panels section
        If Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True Then _
        Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False
        If Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = True Then _
        Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = False
        'resize section heading row
        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.AutoFit
    End Select
    'protect
    EstSht.Protect "WhiteTruckMafia"
End If

    ''''' Liner Panels Options Change
If Not Intersect(Target, Range(EstSht.Range("e1_LinerPanels"), EstSht.Range("Roof_LinerPanels"))) Is Nothing Then
    UpdatesEventsProtection (False)










    If Target.Value = "" Then Target.Value = "None"
    If Target.Value = "None" Then
        Target.offset(0, 1).Value = ""
        Target.offset(0, 2).Value = ""
        Target.offset(0, 3).Value = ""
    End If
    UpdatesEventsProtection (True)










End If


        ''''' Wainscot Section
If Not Intersect(Target, EstSht.Range("Wainscot")) Is Nothing Then
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    'check if yes or no
    Select Case Target.Value
    Case ""
        Target.Value = "No"
    Case "No"
        'Lock Cells
        EstSht.Range("e1_Wainscot").Resize(4, 3).Locked = True
        EstSht.Range("Wainscot_tColor").Locked = True
        'If wainscot table isn't visible, hide column
        If EstSht.Range("AlterWalls").Value <> "Yes" Then
            'do nothing
            If EstSht.Range("LinerPanels").Value = "No" Then
                'resize section heading row
                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
            End If
        ElseIf EstSht.Range("AlterWalls").Value = "Yes" Then
            'Do nothing
            'hide row seperating walls and wainscot table
            If EstSht.Range("LinerPanels").Value = "No" Then
                'resize section heading row
                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15
                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = True
            End If
        End If
        ''''''''''''''''''''''''''''''' reset rows
        UpdatesEventsProtection (False)










        With EstSht
            .Range("e1_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("s2_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("e3_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("s4_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            Range(.Range("e1_Wainscot"), .Range("s4_Wainscot")).Value = "None"
            .Range("Wainscot_tColor").Value = ""
        End With
        UpdatesEventsProtection (True)









    Case "Yes"
        'Unlock Cells
        EstSht.Range("e1_Wainscot").Resize(4, 3).Locked = False
        EstSht.Range("Wainscot_tColor").Locked = False
'        'unhide row above table if needed
        If EstSht.Range("LinerPanels").Value <> "Yes" And EstSht.Range("AlterWalls").Value = "Yes" Then _
        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = False
        ''''''''''''''''''''''''''''''' Set Defaults
        UpdatesEventsProtection (False)








        With EstSht
            .Range("e1_Wainscot").Value = "None"
            .Range("s2_Wainscot").Value = "None"
            .Range("e3_Wainscot").Value = "None"
            .Range("s4_Wainscot").Value = "None"
            .Range("e1_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("s2_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("e3_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("s4_Wainscot").offset(0, 1).Resize(1, 3).Value = ""
            .Range("Wainscot_tColor").Value = "None"
        End With
        UpdatesEventsProtection (True)









    End Select
    'protect
    EstSht.Protect "WhiteTruckMafia"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End If

        ''''' Wainscot Options Change
If Not Intersect(Target, Range(EstSht.Range("e1_Wainscot"), EstSht.Range("s4_Wainscot"))) Is Nothing Then
    UpdatesEventsProtection (False)








    If Target.Value = "" Then Target.Value = "None"
    If Target.Value = "None" Then
        Target.offset(0, 1).Value = ""
        Target.offset(0, 2).Value = ""
    End If
    UpdatesEventsProtection (True)








End If


    '''' Gutter & Downspouts
If Not Intersect(Target, EstSht.Range("GutterAndDownspouts")) Is Nothing Then
    'check if yes or no
    If Target.Value = "" Then Target.Value = "No"
End If


    '''' personnel door number
If Not Intersect(Target, EstSht.Range("PDoorNum")) Is Nothing Then
    UpdatesEventsProtection (False)













     'hide row under quantity box when quantity is 0
    If Target.Value = 0 Or Target.Value = "" Then
        Target.Value = 0
        'reset table to blank
        Me.Range("pDoorCell1").offset(0, 1).Resize(12, 7).Value = ""
        'hide
        If PDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = False Then PDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = True
    Else
        If PDoorStart.offset(-1, 0).EntireRow.Hidden = True Then PDoorStart.offset(-1, 0).EntireRow.Hidden = False
        'check target value
        For i = 0 To 12
            If i <= Target.Value Then
                If PDoorStart.offset(i, 0).EntireRow.Hidden = True Then PDoorStart.offset(i, 0).EntireRow.Hidden = False
            Else
                If PDoorStart.offset(i, 0).EntireRow.Hidden = False Then PDoorStart.offset(i, 0).EntireRow.Hidden = True
            End If
        Next i
    End If


    UpdatesEventsProtection (True)












End If

    '''' OH door
If Not Intersect(Target, EstSht.Range("OHDoorNum")) Is Nothing Then
    UpdatesEventsProtection (False)












    'hide row under quantity box when quantity is 0
    If Target.Value = 0 Or Target.Value = "" Then
        Target.Value = 0
        'reset table to blank
        Me.Range("OHDoorCell1").offset(0, 1).Resize(12, 9).Value = ""
        'hide
        If OHDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = False Then OHDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = True
    Else
        If OHDoorStart.offset(-1, 0).EntireRow.Hidden = True Then OHDoorStart.offset(-1, 0).EntireRow.Hidden = False
        'check target value
        For i = 0 To 12
            If i <= Target.Value Then
                If OHDoorStart.offset(i, 0).EntireRow.Hidden = True Then OHDoorStart.offset(i, 0).EntireRow.Hidden = False
            Else
                If OHDoorStart.offset(i, 0).EntireRow.Hidden = False Then OHDoorStart.offset(i, 0).EntireRow.Hidden = True
            End If
        Next i
    End If



    UpdatesEventsProtection (True)












End If


   '''' Windows
If Not Intersect(Target, EstSht.Range("WindowNum")) Is Nothing Then
    UpdatesEventsProtection (False)












    'hide row under quantity box when quantity is 0
    If Target.Value = 0 Or Target.Value = "" Then
        Target.Value = 0
        'reset table to blank
        Me.Range("WindowCell1").offset(0, 1).Resize(24, 3).Value = ""
        Me.Range("WindowCell1").offset(0, 6).Resize(24, 1).Value = ""
        'hide
        If WindowStart.offset(-1, 0).Resize(26, 1).EntireRow.Hidden = False Then WindowStart.offset(-1, 0).Resize(26, 1).EntireRow.Hidden = True
    Else
        If WindowStart.offset(-1, 0).EntireRow.Hidden = True Then WindowStart.offset(-1, 0).EntireRow.Hidden = False
        'check target value; hide/unhide rows accordingly
        For i = 0 To 24
            If i <= Target.Value Then
                If WindowStart.offset(i, 0).EntireRow.Hidden = True Then WindowStart.offset(i, 0).EntireRow.Hidden = False
            Else
                If WindowStart.offset(i, 0).EntireRow.Hidden = False Then WindowStart.offset(i, 0).EntireRow.Hidden = True
            End If
        Next i
    End If

    UpdatesEventsProtection (True)












End If
''''Window Default Values

''''increase top edge height so that windows can't be lower than building
If Not Intersect(Target, Me.Range("WindowCell1").offset(0, 2).Resize(24, 1)) Is Nothing Then
    If Target.Value > 86 Then
        Target.offset(0, 3).Value = Target.Value / 12
    End If
End If

''''increase top edge height so that MiscFOs can't be lower than building
If Not Intersect(Target, Me.Range("MiscFOCell1").offset(0, 2).Resize(12, 1)) Is Nothing Then
    If Target.Value > (86 / 12) Then
        Target.offset(0, 5).Value = Target.Value
    End If
End If

'if MiscFO is 'field located', only allow 7'2" jambs w/ stool only
If Not Intersect(Target, Me.Range("MiscFOCell1").offset(0, 3).Resize(12, 1)) Is Nothing Then
    If Target.Value = "Field Locate" Then
        Target.offset(0, 7).Value = "7'2"" Jambs w/ Stool"
    End If
End If


   '''' Misc Framed Openings
If Not Intersect(Target, EstSht.Range("MiscFONum")) Is Nothing Then
    UpdatesEventsProtection (False)












    'hide row under quantity box when quantity is 0
    If Target.Value = 0 Or Target.Value = "" Then
        Target.Value = 0
        'reset table to blank
        Me.Range("MiscFOCell1").offset(0, 1).Resize(12, 5).Value = ""
        Me.Range("MiscFOCell1").offset(0, 8).Resize(12, 1).Value = ""
        Me.Range("MiscFOCell1").offset(0, 10).Resize(12, 1).Value = ""
        'hide
        If FOStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = False Then FOStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = True
    Else
        If FOStart.offset(-1, 0).EntireRow.Hidden = True Then FOStart.offset(-1, 0).EntireRow.Hidden = False
        'check target value; hide/unhide rows accordingly
        For i = 0 To 12
            If i <= Target.Value Then
                If FOStart.offset(i, 0).EntireRow.Hidden = True Then FOStart.offset(i, 0).EntireRow.Hidden = False
            Else
                If FOStart.offset(i, 0).EntireRow.Hidden = False Then FOStart.offset(i, 0).EntireRow.Hidden = True
            End If
        Next i
    End If

    UpdatesEventsProtection (True)












End If

'''' building length change
If Not Intersect(Target, EstSht.Range("Building_Length")) Is Nothing Then
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    'if blank building length, reset to 0
    If EstSht.Range("Building_Length").Value = "" Then EstSht.Range("Building_Length").Value = 0

    'change all bay lengths to 0
    Application.EnableEvents = False
    EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0
    Application.EnableEvents = True
    'protect
    EstSht.Protect "WhiteTruckMafia"
End If

'''' bay length change
If Not Intersect(Target, EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1)) Is Nothing Then
    'unprotect
    EstSht.Unprotect "WhiteTruckMafia"
    Application.EnableEvents = False
    'Check if Bay Length Exceeded, reset blanks to 0
    Call BayUpdate(EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1), EstSht.Range("Building_Length"), Intersect(Target, EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1)))
    Application.EnableEvents = True
    'protect
    EstSht.Protect "WhiteTruckMafia"
End If

'''' Trim Color Bulk Change
If Not Intersect(Target, EstSht.Range("All_tColors")) Is Nothing Then
    Application.EnableEvents = False
    If Target.Value = "" Then
        Target.Value = "N/A"
    Else
        'change trim colors to match
        With EstSht
            .Range("Rake_tColor").Value = Target.Value
            .Range("Eave_tColor").Value = Target.Value
            .Range("OutsideCorner_tColor").Value = Target.Value
            .Range("FO_tColor").Value = Target.Value
            .Range("Base_tColor").Value = "None"
            'change gutter/downspout colors as well
            .Range("DownspoutColor").Value = Target.Value
            .Range("GutterColor").Value = Target.Value
        End With
    End If
    Application.EnableEvents = True
End If



'''' overhang table clear
If Not Intersect(Target, OverhangTbl) Is Nothing Then
    Application.EnableEvents = False
    'clear soffits
    For Each cell In Overhangs
        If cell.Row = Target.Row Then
            If cell.Value = "" Or cell.Value = 0 Then
                'clear soffits
                cell.offset(0, 1).Value = ""
                cell.offset(0, 2).Value = ""
                cell.offset(0, 3).Value = ""
                cell.offset(0, 4).Value = ""
                cell.offset(0, 5).Value = ""
            End If
        End If
    Next cell
    Application.EnableEvents = True
End If
'extension table clear
If Not Intersect(Target, ExtensionTbl) Is Nothing Then
    Application.EnableEvents = False
    'clear soffits
    For Each cell In Extensions
        If cell.Row = Target.Row Then
            If cell.Value = "" Or cell.Value = 0 Then
                'clear soffits
                cell.offset(0, 1).Value = ""
                cell.offset(0, 2).Value = ""
                cell.offset(0, 3).Value = ""
                cell.offset(0, 4).Value = ""
                cell.offset(0, 5).Value = ""
            End If
        End If
    Next cell
    Application.EnableEvents = True
End If

'Show/Hide Eave Extension Pitch and Set Intersection default values
With EstSht
    's2 eave extension
    If Not Intersect(Target, .Range("s2_EaveExtension")) Is Nothing Then
        Application.ScreenUpdating = False
        EstSht.Unprotect "WhiteTruckMafia"
        If .Range("s2_EaveExtension").Value = "" Then
            If .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False Then .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True
            .Range("s2e1_Intersection").Value = "N/A"
            .Range("s2e3_Intersection").Value = "N/A"
        Else
            'If previously hidden, unhide and set default option to include intersection
            If .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True Then
                .Range("s2_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False
                'reset to defaults
                .Range("s2_EaveExtensionPitch").Value = "Match Roof"
                If .Range("e1_GableExtension").Value <> "" Then .Range("s2e1_Intersection").Value = "Include"
                If .Range("e3_GableExtension").Value <> "" Then .Range("s2e3_Intersection").Value = "Include"
            End If
        End If
        EstSht.Protect "WhiteTruckMafia"
        Application.ScreenUpdating = True
    End If
    's4 eave extension
    If Not Intersect(Target, .Range("s4_EaveExtension")) Is Nothing Then
        Application.ScreenUpdating = False
        EstSht.Unprotect "WhiteTruckMafia"
        If .Range("s4_EaveExtension").Value = "" Then
            If .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False Then .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True
            .Range("s4e1_Intersection").Value = "N/A"
            .Range("s4e3_Intersection").Value = "N/A"
        Else
            'If previously hidden, unhide and set default option to include intersection
            If .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = True Then
                .Range("s4_EaveExtensionPitch").offset(-1, 0).Resize(4, 1).EntireRow.Hidden = False
                .Range("s4_EaveExtensionPitch").Value = "Match Roof"
                If .Range("e1_GableExtension").Value <> "" Then .Range("s4e1_Intersection").Value = "Include"
                If .Range("e3_GableExtension").Value <> "" Then .Range("s4e3_Intersection").Value = "Include"
            End If
        End If
        EstSht.Protect "WhiteTruckMafia"
        Application.ScreenUpdating = True
    End If
    'e1 gable extension Intersection Option
    If Not Intersect(Target, .Range("e1_GableExtension")) Is Nothing Then
        UpdatesEventsProtection (False)












        If .Range("e1_GableExtension").Value = "" Then
            .Range("s2e1_Intersection").Value = "N/A"
            .Range("s2e1_Intersection").MergeArea.Locked = True
            .Range("s4e1_Intersection").Value = "N/A"
            .Range("s4e1_Intersection").MergeArea.Locked = True
        Else
            'intersection for s2 e1
            If .Range("s2e1_Intersection").Value = "N/A" Then
                .Range("s2e1_Intersection").MergeArea.Locked = False
                If .Range("s2_EaveExtension").Value <> "" Then .Range("s2e1_Intersection").Value = "Include"
            End If
            'intersection for s4 e1
            If .Range("s4e1_Intersection").Value = "N/A" Then
                .Range("s4e1_Intersection").MergeArea.Locked = False
                If .Range("s4_EaveExtension").Value <> "" Then .Range("s4e1_Intersection").Value = "Include"
            End If
        End If
        UpdatesEventsProtection (True)












    End If
    'e3 gable extension Intersection Option
    If Not Intersect(Target, .Range("e3_GableExtension")) Is Nothing Then
        UpdatesEventsProtection (False)












        If .Range("e3_GableExtension").Value = "" Then
            .Range("s2e3_Intersection").Value = "N/A"
            .Range("s2e3_Intersection").MergeArea.Locked = True
            .Range("s4e3_Intersection").Value = "N/A"
            .Range("s4e3_Intersection").MergeArea.Locked = True
        Else
            'intersection for s2 e3
            If .Range("s2e3_Intersection").Value = "N/A" Then
                .Range("s2e3_Intersection").MergeArea.Locked = False
                .Range("s2e3_Intersection").Value = "Include"
            End If
            'intersection for s4 e3
            If .Range("s4e3_Intersection").Value = "N/A" Then
                .Range("s4e3_Intersection").MergeArea.Locked = False
                .Range("s4e3_Intersection").Value = "Include"
            End If
        End If
        UpdatesEventsProtection (True)












    End If
End With


'''' Roof/Wall Panel Shape Change - Disable Translucent Wall Panels and Skylights
With EstSht
    If Not Intersect(Target, Range(.Range("Wall_pShape"), .Range("Roof_pShape"))) Is Nothing Then
        EstSht.Unprotect "WhiteTruckMafia"
        Application.EnableEvents = False
        '''' Panel Shape
        'Hide skylights/translucent wall panels for m-loc
        If .Range("Wall_pShape").Value = "M-Loc" Or .Range("Roof_pShape").Value = "M-Loc" Then
            .Range("TranslucentWallPanelQty").Value = ""
            .Range("SkylightQty").Value = ""
            .Range("TranslucentWallPanelLength").Value = ""
            .Range("SkylightLength").Value = ""
            'hide rows if needed
            If .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = False Then .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = True
        'unhide skylights/translucent wall panels for R-loc
        ElseIf .Range("Wall_pShape").Value <> "M-Loc" And .Range("Roof_pShape").Value <> "M-Loc" Then
            'unhide rows if needed
            If .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = True Then .Range(.Range("TranslucentWallPanelQty"), .Range("SkylightQty")).EntireRow.Hidden = False
        End If
        EstSht.Protect "WhiteTruckMafia"
        Application.EnableEvents = True
    End If
End With

''''''''''' only allow galvalume for panel color selection when prime acrylic galvalume panel is select
With EstSht
    'wall, roof panels
    If Not Intersect(Target, .Range("Wall_pType")) Is Nothing Then
        Call PanelColorOptionCheck(.Range("Wall_pType"), .Range("Wall_Color"))
    ElseIf Not Intersect(Target, .Range("Roof_pType")) Is Nothing Then
        Call PanelColorOptionCheck(.Range("Roof_pType"), .Range("Roof_Color"))
    End If
    'liner panels
    If Not Intersect(Target, .Range("e1_LinerPanels").offset(0, 2)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e1_LinerPanels").offset(0, 2), .Range("e1_LinerPanels").offset(0, 3))
    End If
    If Not Intersect(Target, .Range("s2_LinerPanels").offset(0, 2)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s2_LinerPanels").offset(0, 2), .Range("s2_LinerPanels").offset(0, 3))
    End If
    If Not Intersect(Target, .Range("e3_LinerPanels").offset(0, 2)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e3_LinerPanels").offset(0, 2), .Range("e3_LinerPanels").offset(0, 3))
    End If
    If Not Intersect(Target, .Range("s4_LinerPanels").offset(0, 2)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s4_LinerPanels").offset(0, 2), .Range("s4_LinerPanels").offset(0, 3))
    End If
    If Not Intersect(Target, .Range("Roof_LinerPanels").offset(0, 2)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("Roof_LinerPanels").offset(0, 2), .Range("Roof_LinerPanels").offset(0, 3))
    End If
    'wainscot
    If Not Intersect(Target, .Range("e1_Wainscot").offset(0, 1)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e1_Wainscot").offset(0, 1), .Range("e1_Wainscot").offset(0, 2))
    End If
    If Not Intersect(Target, .Range("s2_Wainscot").offset(0, 1)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s2_Wainscot").offset(0, 1), .Range("s2_Wainscot").offset(0, 2))
    End If
    If Not Intersect(Target, .Range("e3_Wainscot").offset(0, 1)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e3_Wainscot").offset(0, 1), .Range("e3_Wainscot").offset(0, 2))
    End If
    If Not Intersect(Target, .Range("s4_Wainscot").offset(0, 1)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s4_Wainscot").offset(0, 1), .Range("s4_Wainscot").offset(0, 2))
    End If
    'overhangs
    If Not Intersect(Target, .Range("e1_GableOverhang").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e1_GableOverhang").offset(0, 3), .Range("e1_GableOverhang").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("s2_EaveOverhang").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s2_EaveOverhang").offset(0, 3), .Range("s2_EaveOverhang").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("e3_GableOverhang").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e3_GableOverhang").offset(0, 3), .Range("e3_GableOverhang").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("s4_EaveOverhang").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s4_EaveOverhang").offset(0, 3), .Range("s4_EaveOverhang").offset(0, 4))
    End If
    'extensions
    If Not Intersect(Target, .Range("e1_GableExtension").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e1_GableExtension").offset(0, 3), .Range("e1_GableExtension").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("s2_EaveExtension").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s2_EaveExtension").offset(0, 3), .Range("s2_EaveExtension").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("e3_GableExtension").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("e3_GableExtension").offset(0, 3), .Range("e3_GableExtension").offset(0, 4))
    End If
    If Not Intersect(Target, .Range("s4_EaveExtension").offset(0, 3)) Is Nothing Then
        Call PanelColorOptionCheck(.Range("s4_EaveExtension").offset(0, 3), .Range("s4_EaveExtension").offset(0, 4))
    End If
End With



'''''''''''''''''''''''''''''''''''''''''''''''''' Framed Opening Option Changes '''''''''''''''''''''
With EstSht
    '''''''''''''''''''''''''''''''''''''''''' Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''
    If Not Intersect(Target, Range(.Range("pDoorCell1").offset(0, 1), .Range("pDoorCell12").offset(0, 1))) Is Nothing Then
        EstSht.Unprotect "WhiteTruckMafia"
        Application.EnableEvents = False
        If Target.Value = "4070" Then
            'Remove half glass option
            With Target.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'Remove dead bolt option
            With Target.offset(0, 5).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'Set Default Values for 4070
            Target.offset(0, 2).Value = "No"
            Target.offset(0, 5).Value = "No"
            If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "No" 'only if blank, keep selected values on change
            If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = "8.25" 'only if blank, keep selected values on change
        ElseIf Target.Value = "3070" Or Target.Value = "" Then
            'restore half glass, deadbolt
            With Target.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, 2).Validation.Value = False Then Target.offset(0, 2).Value = ""
            With Target.offset(0, 5).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'Set Default Values for 3070 and reset for blank
            If Target.Value = "3070" Then
                If Target.offset(0, 2).Value = "" Then Target.offset(0, 2).Value = "No" 'only if blank, keep selected values on change
                If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "No" 'only if blank, keep selected values on change
                If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = 8.25 'only if blank, keep selected values on change
                If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Value = "No" 'only if blank, keep selected values on change
            ElseIf Target.Value = "" Then
                Target.offset(0, 2).Value = ""
                Target.offset(0, 3).Value = ""
                Target.offset(0, 4).Value = ""
                Target.offset(0, 5).Value = ""
            End If
            If Target.offset(0, 5).Validation.Value = False Then Target.offset(0, 5).Value = ""
        End If
        EstSht.Protect "WhiteTruckMafia"
        Application.EnableEvents = True
    End If

     '''''''''''''''''''''''''''''''''''''''''' Overhead Doors '''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''Set Default Values''''''''''''''''
    'If height or width is entered, set default values; "Type" change (RUD/Sectional) will fix validation as defined below
    'Width
    If Not Intersect(Target, Range(.Range("OHDoorCell1").offset(0, 1), .Range("OHDoorCell12").offset(0, 1))) Is Nothing Then
        If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "Sectional"
        If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = "None"
        If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Value = "Manual"
        If Target.offset(0, 6).Value = "" Then Target.offset(0, 6).Value = "No"
        If Target.offset(0, 7).Value = "" Then Target.offset(0, 7).Value = "None"
        If Target.offset(0, 8).Value = "" Then Target.offset(0, 8).Value = 0
    End If
    'Height
    If Not Intersect(Target, Range(.Range("OHDoorCell1").offset(0, 2), .Range("OHDoorCell12").offset(0, 2))) Is Nothing Then
        If Target.offset(0, 2).Value = "" Then Target.offset(0, 2).Value = "Sectional"
        If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "None"
        If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = "Manual"
        If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Value = "No"
        If Target.offset(0, 6).Value = "" Then Target.offset(0, 6).Value = "None"
        If Target.offset(0, 7).Value = "" Then Target.offset(0, 7).Value = 0
    End If
    If Not Intersect(Target, Range(.Range("OHDoorCell1").offset(0, 4), .Range("OHDoorCell12").offset(0, 4))) Is Nothing Then
        EstSht.Unprotect "WhiteTruckMafia"
        Application.EnableEvents = False
''''''''''''''''''''''''' Roll Up Doors''''''''''''''''''''''
        If Target.Value = "RUD" Then
            '''sizing options
            'width
            With Target.offset(0, -3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("RUDWidth").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'clear if invalid
            If Target.offset(0, -3).Validation.Value = False Then Target.offset(0, -3).Value = ""
            'height
            With Target.offset(0, -2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("RUDHeight").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            'clear if invalid
            If Target.offset(0, -2).Validation.Value = False Then Target.offset(0, -2).Value = ""
            '''remove insulation, operation, windows, and high lift options
            'insulation
            With Target.offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="None"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            Target.offset(0, 1).Value = "None"
            'Operation
            With Target.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Chain Hoist"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            Target.offset(0, 2).Value = "Chain Hoist"
            'High Lift
            With Target.offset(0, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            Target.offset(0, 3).Value = "No"
            'Windows
            With Target.offset(0, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="None"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            Target.offset(0, 4).Value = "None"
''''''''''''''''''''''''' Sectional Doors''''''''''''''''''''''
        ElseIf Target.Value = "Sectional" Or Target.Value = "" Then
            '''sizing options
            'width
            With Target.offset(0, -3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("SectionalOHDoorWidth").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, -3).Validation.Value = False Then Target.offset(0, -3).Value = ""
            'height
            With Target.offset(0, -2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("SectionalOHDoorHeight").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, -2).Validation.Value = False Then Target.offset(0, -2).Value = ""
            'insulation
            With Target.offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorInsulationOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, 1).Validation.Value = False Then Target.offset(0, 1).Value = ""
            'Operation
            With Target.offset(0, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorOperationOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, 2).Validation.Value = False Then Target.offset(0, 2).Value = ""
            'High Lift
            With Target.offset(0, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, 3).Validation.Value = False Then Target.offset(0, 3).Value = ""
            'Windows
            With Target.offset(0, 4).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Lists!" & ListSht.Range("OHDoorWindowOptions").Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            If Target.offset(0, 4).Validation.Value = False Then Target.offset(0, 4).Value = ""
            'change options

        End If
        EstSht.Protect "WhiteTruckMafia"
        Application.EnableEvents = True
    End If

    '''''''''''''''Misc FOs '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Exhaust Fans/Louvers default to "None"'
    'Width
    If Not Intersect(Target, Range(.Range("MiscFOCell1").offset(0, 1), .Range("MiscFOCell12").offset(0, 1))) Is Nothing Then
        If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "None"
        If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = "None"
        If Target.offset(0, 6).Value = "" Then Target.offset(0, 6).Formula = "=86/12"
        If Target.offset(0, 7).Value = "" Then Target.offset(0, 7).Value = 0
    End If

    'Height
    If Not Intersect(Target, Range(.Range("MiscFOCell1").offset(0, 2), .Range("MiscFOCell12").offset(0, 2))) Is Nothing Then
        If Target.offset(0, 2).Value = "" Then Target.offset(0, 2).Value = "None"
        If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Value = "None"
        If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Formula = "=86/12"
        If Target.offset(0, 6).Value = "" Then Target.offset(0, 6).Value = 0
    End If

        '''''''''''''''Windows '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Width
    If Not Intersect(Target, Range(.Range("WindowCell1").offset(0, 1), .Range("WindowCell12").offset(0, 1))) Is Nothing Then
        If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Formula = "=86/12"
        If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Value = 0
    End If

    'Height
    If Not Intersect(Target, Range(.Range("MiscFOCell1").offset(0, 2), .Range("MiscFOCell12").offset(0, 2))) Is Nothing Then
        If Target.offset(0, 3).Value = "" Then Target.offset(0, 3).Formula = "=86/12"
        If Target.offset(0, 4).Value = "" Then Target.offset(0, 4).Value = 0
    End If

        '''''''''''''''Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Width
    If Not Intersect(Target, Range(.Range("pDoorCell1").offset(0, 1), .Range("pDoorCell12").offset(0, 1))) Is Nothing Then
        If Target.offset(0, 6).Value = "" Then Target.offset(0, 6).Value = 0
    End If

    'Height
    If Not Intersect(Target, Range(.Range("MiscFOCell1").offset(0, 2), .Range("MiscFOCell12").offset(0, 2))) Is Nothing Then
        If Target.offset(0, 5).Value = "" Then Target.offset(0, 5).Value = 0
    End If

    '''''''''''''''Overhangs and Extensions''''''''''''''''''''''''''''''''''''''''''''''
    'Overhangs Set Default Values
    If Not Intersect(Target, Range(.Range("e1_GableOverhang"), .Range("s4_EaveOverhang"))) Is Nothing Then
        If Target.offset(0, 1).Value = "" Then Target.offset(0, 1).Value = "No"
    End If
    'Extensions Set Default Values
    If Not Intersect(Target, Range(.Range("e1_GableExtension"), .Range("s4_EaveExtension"))) Is Nothing Then
        If Target.offset(0, 1).Value = "" Then Target.offset(0, 1).Value = "No"
    End If



End With





End Sub
