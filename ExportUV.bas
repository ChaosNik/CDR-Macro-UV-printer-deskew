' SPDX-License-Identifier: GPL-3.0-or-later
' =========================================================================================
'  Project   : ExportUV (CorelDRAW VBA Macro)
'  File      : ExportUV.bas
'  Version   : 1.0.0
'  Date      : 2025-08-31
'  Author    : Nikola Karpic
'
'  Summary
'    Automates a UV prepress transform on the active page:
'      • Optionally finds a 310×500 mm rectangle and renames it START_FRAME.
'      • Groups page shapes as MY_OBJECT (or reuses the single existing group).
'      • Builds a straight-edge polygon (ENV_TRANSFORM) and applies it as an Envelope
'        in Original mode (keeps lines straight).
'      • Nudges content by +0.525 mm (X) and +0.2 mm (Y).
'      • Draws a transparent registration frame MY_FRAME.
'      • Deletes START_FRAME if found anywhere (groups/PowerClips/pages).
'
'  Copyright
'    Copyright (C) 2025  Nikola Karpic
'
'  License
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'  Notes
'    • Keep a copy of the full GPL-3.0 license text in a file named LICENSE
'      alongside this source file for distribution.
'    • If you distribute modified versions, you must license them under GPL-3.0
'      (or later) and provide the corresponding source.
' =========================================================================================

Option Explicit

Public Sub ExportUV()
    Dim doc As Object, pg As Object, lyr As Object
    Dim srAll As Object, grp As Object, sEnv As Object, sFrame As Object
    Dim eff As Object
    Dim prevUnit As Long

    On Error GoTo CleanFail

    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    Set lyr = pg.ActiveLayer

    ' --- preserve user prefs
    prevUnit = doc.Unit
    doc.Unit = 3        ' cdrMillimeter
    Application.Optimization = True
    Application.EventsEnabled = False
    doc.BeginCommandGroup "Envelope + Frame macro"
	
	' 0) BEFORE FIRST STEP: rename 310x500 mm rectangle to START_FRAME
    Call TryRenameRectToSTART_FRAME

    ' 1) Group all objects (if needed) and name the group MY_OBJECT
    Set srAll = pg.Shapes.All
    If (srAll Is Nothing) Or srAll.Count = 0 Then
        Err.Raise vbObjectError + 1, , "No shapes on the active page."
    End If

    If srAll.Count = 1 And srAll(1).Type = 7 Then   ' cdrGroupShape
        Set grp = srAll(1)
    Else
        Set grp = srAll.Group
    End If
    grp.Name = "MY_OBJECT"

    ' 2) Draw closed polygon ENV_TRANSFORM (straight nodes) + CMYK blue outline
    Set sEnv = CreateClosedStraightPolygon(lyr, Array( _
        Array(-47.25, 398.496), _
        Array(261.1, 399.65), _
        Array(258.393, -100.644), _
        Array(-50.05, -102.25) _
    ))
    sEnv.Name = "ENV_TRANSFORM"

    ' CMYK “blue”: C=100 M=100 Y=0 K=0
    On Error Resume Next
    sEnv.Outline.Color.CMYKAssign 100, 100, 0, 0
    On Error GoTo CleanFail

    ' 3) Apply envelope to MY_OBJECT using ENV_TRANSFORM
    ' 4) Force ORIGINAL mapping (not Putty) and keep straight lines
    '    Mode = 1 (cdrEnvelopeOriginal), KeepLines = True
    Set eff = grp.CreateEnvelopeFromShape(sEnv, 1, True) ' returns an Effect
    ' (API ref: Shape.CreateEnvelopeFromShape)
    ' cdrEnvelopeMode docs list Original = 1

    ' 5) Delete ENV_TRANSFORM polygon
    sEnv.Delete

    ' 6) Move all objects 0.525 mm right and 0.2 mm up
    Set srAll = pg.Shapes.All
    srAll.Move 0.525, 0.2

    ' 6b) Draw transparent square MY_FRAME with opposite corners
    '     (260.0, 398.5) and (-50.0, -101.5)
    Set sFrame = lyr.CreateRectangle(-50#, 398.5, 260#, -101.5)
    sFrame.Name = "MY_FRAME"
    sFrame.Fill.ApplyNoFill
    sFrame.Outline.SetNoOutline
	
	' 7) AFTER LAST STEP: delete START_FRAME wherever it is
	Dim startSh As Object

	' First look inside the group (most likely location)
	Set startSh = FindShapeByNameRecursive(grp.Shapes, "START_FRAME")

	' Fallbacks in case it isn't inside the group
	If startSh Is Nothing Then Set startSh = FindShapeByNameRecursive(pg.Shapes, "START_FRAME")
	If startSh Is Nothing Then Set startSh = FindShapeByNameRecursive(ActiveDocument.Shapes, "START_FRAME")

	If Not startSh Is Nothing Then
		On Error Resume Next
		startSh.Delete
		On Error GoTo 0
	End If

CleanExit:
    doc.EndCommandGroup
    Application.Optimization = False
    Application.EventsEnabled = True
    doc.Unit = prevUnit
    Exit Sub

CleanFail:
    MsgBox "Macro failed: " & Err.Description, vbExclamation, "CorelDRAW Macro"
    Resume CleanExit
End Sub

' --- helper: build a closed straight-edge polygon from {x,y} points (mm)
Private Function CreateClosedStraightPolygon( _
    ByVal targetLayer As Object, _
    ByVal pts As Variant _
) As Object
    Dim crv As Object, sp As Object
    Dim i As Long

    Set crv = CreateCurve(ActiveDocument) ' Corel global function
    Set sp = crv.CreateSubPath(pts(LBound(pts))(0), pts(LBound(pts))(1))

    For i = LBound(pts) + 1 To UBound(pts)
        sp.AppendLineSegment pts(i)(0), pts(i)(1), False
    Next i

    sp.Closed = True
    Set CreateClosedStraightPolygon = targetLayer.CreateCurve(crv)
End Function

Private Sub TryRenameRectToSTART_FRAME()
    Const TARGET_W_MM As Double = 310#
    Const TARGET_H_MM As Double = 500#
    Const TOL_MM      As Double = 0.05

    Dim d As Object, p As Object
    Set d = ActiveDocument
    If d Is Nothing Then Exit Sub

    For Each p In d.Pages
        If FindAndRenameInShapes(p.Shapes, TARGET_W_MM, TARGET_H_MM, TOL_MM) Then Exit For
    Next p
End Sub

Private Function FindAndRenameInShapes(ByVal shpCol As Object, _
                                       ByVal wMM As Double, ByVal hMM As Double, ByVal tolMM As Double) As Boolean
    Dim s As Object
    For Each s In shpCol
        If IsTargetRectangle(s, wMM, hMM, tolMM) Then
            On Error Resume Next
            s.Name = "START_FRAME"
            On Error GoTo 0
            FindAndRenameInShapes = True
            Exit Function
        End If

        ' Recurse into grouped/compound shapes
        If Not s.Shapes Is Nothing Then
            If s.Shapes.Count > 0 Then
                If FindAndRenameInShapes(s.Shapes, wMM, hMM, tolMM) Then
                    FindAndRenameInShapes = True
                    Exit Function
                End If
            End If
        End If

        ' Recurse into PowerClip contents (if present)
        If Not s.PowerClip Is Nothing Then
            If FindAndRenameInShapes(s.PowerClip.Shapes, wMM, hMM, tolMM) Then
                FindAndRenameInShapes = True
                Exit Function
            End If
        End If
    Next s
End Function

Private Function IsTargetRectangle(ByVal s As Object, _
                                   ByVal wMM As Double, ByVal hMM As Double, ByVal tolMM As Double) As Boolean
    If s.Type <> cdrRectangleShape Then Exit Function

    ' Units are already mm (doc.Unit set at the start)
    Dim w As Double, h As Double
    w = s.SizeWidth
    h = s.SizeHeight

    If (NearlyEqual(w, wMM, tolMM) And NearlyEqual(h, hMM, tolMM)) _
       Or (NearlyEqual(w, hMM, tolMM) And NearlyEqual(h, wMM, tolMM)) Then
        IsTargetRectangle = True
    End If
End Function

Private Function NearlyEqual(ByVal a As Double, ByVal b As Double, ByVal tol As Double) As Boolean
    NearlyEqual = (Abs(a - b) <= tol)
End Function

' --- helper: find a shape by name inside a Shapes collection, searching groups and PowerClips
Private Function FindShapeByNameRecursive(ByVal shpCol As Object, ByVal targetName As String) As Object
    Dim s As Object
    Dim found As Object
    For Each s In shpCol
        ' compare case-insensitively
        If StrComp(s.Name, targetName, vbTextCompare) = 0 Then
            Set FindShapeByNameRecursive = s
            Exit Function
        End If

        ' dive into grouped/compound shapes safely
        On Error Resume Next
        If Not s Is Nothing Then
            ' Many “group-like” shapes expose .Shapes; accessing it on non-groups can error, so keep Resume Next on
            If Not s.Shapes Is Nothing Then
                If s.Shapes.Count > 0 Then
                    Set found = FindShapeByNameRecursive(s.Shapes, targetName)
                    If Not found Is Nothing Then
                        Set FindShapeByNameRecursive = found
                        On Error GoTo 0
                        Exit Function
                    End If
                End If
            End If
        End If

        ' dive into PowerClip contents (if any)
        Dim hasPC As Boolean
        hasPC = False
        If Err.Number <> 0 Then Err.Clear
        If Not s Is Nothing Then
            ' Access can raise on some shapes; keep Resume Next
            hasPC = Not s.PowerClip Is Nothing
        End If
        If Err.Number <> 0 Then hasPC = False: Err.Clear
        On Error GoTo 0

        If hasPC Then
            Set found = FindShapeByNameRecursive(s.PowerClip.Shapes, targetName)
            If Not found Is Nothing Then
                Set FindShapeByNameRecursive = found
                Exit Function
            End If
        End If
    Next s
End Function


