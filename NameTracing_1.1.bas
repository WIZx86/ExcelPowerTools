Attribute VB_Name = "NameTracing"
Option Explicit

Sub NameStats()
    Dim n As Name
    Dim ws As Worksheet
    Dim c As Range
    Dim specialC As Range
    Dim searchNames As New Collection
    Dim thiswb As Workbook: Set thiswb = ThisWorkbook
    Dim nameused As Long
    Dim celladdr As String
    Dim namecells As String
    Dim iter As Variant
    Dim iter2 As Variant
    Dim matchedCells As Collection
    Dim matchedNames As New Collection
    Dim collOutput As New Collection
    Dim foundmatch As Boolean
    Dim i As Long
    With thiswb
        For Each n In .Names
            If InStr(1, n.Name, "_") <> 1 Then searchNames.Add n
        Next n
        For Each n In searchNames
            nameused = 0
            namecells = ""
            For i = 1 To searchNames.Count
            Set iter = searchNames(i)
                If CheckName(n, iter.Value) Then
                    nameused = nameused + 1
                    namecells = IIf(namecells = "", iter.Name, namecells & "," & iter.Name)
                End If
            Next i
            matchedNames.Add nameused & "|" & namecells, n.Name
            nameused = 0
            namecells = ""
            Set matchedCells = New Collection
            Set iter = Nothing
            For Each ws In .Worksheets
                With ws.Cells
                    On Error Resume Next
                    Set specialC = .SpecialCells(xlCellTypeFormulas)
                    If Err.Number Then Set specialC = Nothing
                    On Error GoTo 0
                    If Not specialC Is Nothing Then
                        For Each c In specialC
                            With c
                                If CheckName(n, .Formula2) Then
                                    nameused = nameused + 1
                                    If matchedCells.Count = 0 Then
                                        matchedCells.Add c
                                    Else
                                        foundmatch = False
                                        For i = 1 To matchedCells.Count
                                            Set iter = matchedCells(i)
                                            If iter.Parent Is ws Then
                                                Set iter = Union(iter, c)
                                                matchedCells.Add iter, after:=i
                                                matchedCells.Remove i
                                                foundmatch = True
                                                Exit For
                                            End If
                                        Next i
                                        If Not foundmatch Then matchedCells.Add c
                                    End If
                                End If
                            End With
                        Next c
                    End If
                End With
            Next ws
            For Each iter In matchedCells
                With iter
                    celladdr = "'" & .Parent.Name & "'!" & .Address
                    If namecells = "" Then
                        namecells = celladdr
                    Else
                        namecells = namecells & ";" & celladdr
                    End If
                End With
            Next iter
            collOutput.Add nameused & "|" & namecells, n.Name
            Set matchedCells = Nothing
        Next n
        Open .Path & "\nametrace[" & .Name & "].txt" For Output As #1
        For Each n In searchNames
            Dim nmdepend As String
            Dim celldepend As String
            Dim totaldepend As Long
            With n
                nmdepend = matchedNames(.Name)
                celldepend = collOutput(.Name)
                totaldepend = Split(nmdepend, "|")(0) * 1 + Split(celldepend, "|")(0) * 1
                If Left(nmdepend, 1) = "0" Then nmdepend = "none"
                If Left(celldepend, 1) = "0" Then celldepend = "none"
                Print #1, .Name & " (" & totaldepend & "):" & vbNewLine & Chr(9) & "NameDependents [" & nmdepend & "]" & _
                    vbNewLine & Chr(9) & "CellDependents [" & celldepend & "]" & vbNewLine
            End With
        Next n
        Close #1
    End With
End Sub
Function CheckName(nm As Name, frm As String) As Boolean
    Dim validchars As String: validchars = "#@& ,()%/*-+^=><"
    Dim islambda As Boolean
    Dim prevalid As Boolean
    Dim postvalid As Boolean
    Dim foundat As Long
    Dim prechar As String
    Dim postchar As String
    Dim nmstr As String: nmstr = nm.Name
    Dim nmlen As Long: nmlen = Len(nmstr)
    Dim frmlen As Long: frmlen = Len(frm)
    islambda = InStr(1, nm.Value, "=LAMBDA(") <> 0
    foundat = InStr(1, frm, nmstr & IIf(islambda, "(", ""))
    If foundat Then
        prechar = IIf(foundat <> 0, Mid(frm, foundat - 1, 1), "+")
        postchar = Mid(frm, foundat + nmlen, 1)
        prevalid = ((foundat = 1) + 1) * IIf(foundat = 2, True, InStr(2, validchars, prechar))
        postvalid = IIf(foundat + nmlen - 1 = frmlen, True, InStr(1, validchars, postchar))
    End If
    If islambda Then
        CheckName = foundat
    Else
        CheckName = foundat * prevalid * postvalid
    End If
End Function

