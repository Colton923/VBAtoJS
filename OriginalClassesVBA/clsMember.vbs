Option Explicit

Public mType As String  '''''''''''''' classified by type: column, rafter, girt, eave strut, etc.
Public Location As String   '''''''''''' Descriptive location string
Public Length As Double ''''''''''' Used for the main functional span dimension of the member
Public Depth As Double
Public Width As Double
Public tType As String 'only used for trim, but helpful here for error checking when trim is in the same collection as members (in FOs)
Public Measurement As String
Public Qty As Integer
Public CL As Double     ' position of the column's centerline as measured from the right wall corner
Public rEdgePosition As Double  ''''''''' Position of the column's right edge. Since columns are currently modeled without width, the CL, right edge, and Left edge position are all equal
Public DeleteFlag As Boolean
Public bEdgeHeight As Double
Public clsType As String
Public tEdgeHeight As Double
Public Placement As String 'string used to store misc descriptions
Public ComponentMembers As Collection
Public LoadBearing As Boolean
Public RafterLeftEdge As Double 'exclusively used for rafters
Public Size As String   '''' Size string is the specific dimensions of the mType i.e. - "W8x12" or "8" C Purlin"


Public Function lEdgePosition() As Double
    'for receiver cee's, 0 width for the purpose of positioning since purlins will essentually fit flush into it
    If InStr(1, mType, "Receiver Cee") = True Then
        'receiver cee should never have a l/r edge position even if we're tracking other column's edges because of their orintation. *This is at least true when they're functioning as jambs*
        lEdgePosition = rEdgePosition
    Else
        lEdgePosition = rEdgePosition + Width
    End If
End Function

''''''''''''''''''''''''' Sub for finding the member's size string (and width) using the structural steel lookup table
Public Sub SetSize(b As clsBuilding, ColumnOrRafter As String, Location As String, HorizontalReferenceDistance As Double, Optional CustomNonExpandable As String)
'''' Valid Location Options: "Interior", "e1","s2","e3",and "s4"
Dim LookupTbl As ListObject
Dim LookupHeight As Double
Dim LookupHorizontalIndex As Double
Dim LookupSizeString As String
Dim NearestHorizontalValue As Double


If ColumnOrRafter = "Rafter" Then
    Set LookupTbl = LookupTblMatch(b, ColumnOrRafter, Location)
    LookupHeight = Application.WorksheetFunction.RoundUp((tEdgeHeight / 12) / 10, 0) * 10
    If HorizontalReferenceDistance <= 25 * 12 Then
        LookupHorizontalIndex = 1      ' default to 30' minimum for a given horizontal distance of less than 30'
    Else
        LookupHorizontalIndex = Application.WorksheetFunction.RoundUp((HorizontalReferenceDistance / 12) / 5 - 5, 0) + 1
        'LookupHorizontalIndex = NearestHorizontalValue
        'LookupHorizontalIndex = Application.WorksheetFunction.RoundDown((((Application.WorksheetFunction.RoundUp((HorizontalReferenceDistance / 12) / 10, 0) * 10) - 25) / 10) + 1, 0)
    End If
    If LookupHeight > 80 Then GoTo BadLookupData
    If LookupHeight < 20 Then LookupHeight = 20
    If LookupHorizontalIndex > 12 Then
        If Location = "e1" Or Location = "e3" Then
            LookupHorizontalIndex = 12
        Else
            GoTo BadLookupData
        End If
    End If
    With LookupTbl
        Size = .DataBodyRange(.ListRows(CStr(LookupHorizontalIndex)).Index, .ListColumns(CStr(LookupHeight)).Index)
    End With
    If InStr(1, Size, "TS") <> 0 Then
        Width = 4
    ElseIf InStr(1, Size, "W") <> 0 Then
        Width = Right(Left(Size, InStr(1, Size, "x") - 1), Len(Left(Size, InStr(1, Size, "x") - 1)) - 1)
    End If
ElseIf ColumnOrRafter = "Column" Then
    If CustomNonExpandable = "NonExpandable" Then
        Set LookupTbl = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl")
    Else
        Set LookupTbl = LookupTblMatch(b, ColumnOrRafter, Location)
    End If
    LookupHeight = Application.WorksheetFunction.RoundUp((tEdgeHeight / 12) / 10, 0) * 10
    If HorizontalReferenceDistance < 30 * 12 Then
        LookupHorizontalIndex = 1      ' default to 30' minimum for a given horizontal distance of less than 30'
    Else
        LookupHorizontalIndex = (((Application.WorksheetFunction.RoundUp((HorizontalReferenceDistance / 12) / 10, 0) * 10) - 30) / 10) + 1
    End If
    If LookupHeight > 80 Then GoTo BadLookupData
    If LookupHeight < 20 Then LookupHeight = 20
    If LookupHorizontalIndex > 6 Then
        If Location = "e1" Or Location = "e3" Then
            LookupHorizontalIndex = 6
        Else
            'GoTo BadLookupData WHY IS S2 and S4 sending this to BADDATA??????????
            LookupHorizontalIndex = 6
        End If
    End If
    With LookupTbl
        Size = .DataBodyRange(.ListRows(CStr(LookupHorizontalIndex)).Index, .ListColumns(CStr(LookupHeight)).Index)
    End With
    If InStr(1, Size, "TS") <> 0 Then
        Width = 4
    ElseIf InStr(1, Size, "W") <> 0 Then
        Width = Right(Left(Size, InStr(1, Size, "x") - 1), Len(Left(Size, InStr(1, Size, "x") - 1)) - 1)
    End If
End If

Exit Sub

Set LookupTbl = Nothing

BadLookupData:
If LookupHorizontalIndex > 80 Then
    MsgBox "A horizontal lookup distance of greater than 80' has been calculated!", vbCritical, "Member Lookup Error"
ElseIf LookupHorizontalIndex > 80 Then
    MsgBox "A lookup height of greater than 80' has been calculated!", vbCritical, "Member Lookup Error"
End If
Stop
Exit Sub
LookupFail:
MsgBox "Member size lookup failed! Bad lookup string returned.", vbCritical, "Member Lookup Error"
Stop
End Sub

''''''''''''''''''''''''' Function that sets the correct steel lookup table
Private Function LookupTblMatch(b As clsBuilding, ColumnsOrRafters As String, Optional Wall As String) As ListObject
'Function Note: This does not properly handle non expandable endwall rafter lines, which should be set to either 8" receiver cee or 10" receiver cee depending on the length of the adjacent bay.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''Columns or Rafters
If ColumnsOrRafters = "Rafter" Then
    Set LookupTblMatch = SteelLookupSht.ListObjects("MainRafterAndExpandableEndwallRafterTbl")
ElseIf ColumnsOrRafters = "Column" Then
    '''''''''''''''''''' For columns, select table based off of walls ''''''''''''''''''''''''''
    Select Case Wall
    Case "s2", "s4"
        Set LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl")
    Case "e1"
        If b.ExpandableEndwall("e1") = True Then
            Set LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl")
        ElseIf b.ExpandableEndwall("e1") = False Then
            Set LookupTblMatch = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl")
        End If
    Case "e3"
        If b.ExpandableEndwall("e3") = True Then
            Set LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl")
        ElseIf b.ExpandableEndwall("e3") = False Then
            Set LookupTblMatch = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl")
        End If
    Case "Interior"
        Set LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl")
    End Select
End If
End Function


Public Sub SetType(mType, Optional mName As String)     '''''''''mName is nonfunctional for now- this wil be taken from the Structural Steel Lookup tables. Unsure the best way to do this yet.
Select Case mType
    Case "TS"
        Depth = 4   'this information is true, but it should be currently unused
        Width = 4
    Case "W-Beam"
        Depth = Right(Left(mName, InStr(1, mName, "x") - 1), Len(Left(mName, InStr(1, mName, "x") - 1)) - 1)
        'mid(left(activecell.Value,instr(1,activecell.Value,"x")-1),len(left(activecell.Value,instr(1,activecell.Value,"x")-1))-1)
    Case "8"" Receiver Cee"
        Width = 8
    Case "10"" Receiver Cee"
        Width = 10
    Case "C Purlin"
        'width unknown- Mr. Morgan to confirm (although widths of these members don't appear to be relevant)
    End Select
End Sub


Private Sub Class_Initialize()
'default qty to 1
Qty = 1
clsType = "Member"
Set ComponentMembers = New Collection
LoadBearing = False

End Sub
