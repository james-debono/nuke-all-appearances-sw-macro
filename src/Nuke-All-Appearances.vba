'==============================================================================
' Nuke All Appearances
'
' Strips every appearance of every kind, everywhere in the document - from every
' display state and every configuration, and from every part inside an assembly.
'
' This is the blunt instrument of the appearance macros. Where Remove All
' Appearances clears one display state and stops at the assembly boundary, this
' one stops at nothing:
'
'   - Face, feature, body, part and component appearances all go.
'   - Every display state, not just the one you are looking at.
'   - Every configuration, not just the active one.
'   - In an assembly, the part files themselves are opened and cleared, so the
'     colours held inside them go too.
'
' Because it reaches into the parts an assembly references, it changes documents
' you did not open yourself. A part used in three assemblies is changed for all
' three.
'
' Nothing is written to disk. The parts are modified in memory and left unsaved,
' so SOLIDWORKS asks about them when you close, and closing without saving puts
' everything back. Saving is the point of no return, and it is yours to choose.
'
' Suppressed and lightweight components are resolved first, because an appearance
' cannot be cleared from a component that is not loaded. They stay resolved
' afterwards - which is also undone by closing without saving.
'
' There is no undo. Appearance changes leave no undo record, so the macro asks
' before it runs.
'
' To use, open a part or assembly document and run the macro.
'
'   Version   0.1.0
'   Date      2026-09-11
'   Author    James Debono
'   Licence   MIT - full text below
'   Source    https://github.com/james-debono/nuke-all-appearances-sw-macro
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.1.0   First release.
'
'------------------------------------------------------------------------------
' MIT Licence
' SPDX-License-Identifier: MIT
'
' Copyright (c) 2026 James Debono
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' This macro is deliberately built on the approach its two siblings reject.
'
' Remove All Appearances and Remove Body and Component Appearances both list
' RemoveMaterialProperty as a dead end from their 0.1.0: it is scoped by
' CONFIGURATION, and several display states live under one configuration, so it
' cleared every display state at once. That was a defect there. Here it is the
' entire point, and swAllConfiguration widens it again to cover every
' configuration as well. Do not "fix" this back to the render material API's
' per-display-state scoping - that would quietly undo what the macro is for.
'
' Two sweeps run over every document, because SOLIDWORKS scopes appearances on
' two independent axes and neither sweep provably covers both:
'
'   Sweep A - render materials, across all display states.
'       GetRenderMaterials2(swAllDisplayState) returns each appearance object;
'       RemoveAllEntities empties it. An appearance with no entities is gone.
'       Nothing is ever written back, because nothing is ever kept - which is
'       why the 161-second re-attach that sank Remove All Appearances 0.3.0
'       cannot happen here. This is also the only sweep that reaches
'       COMPONENT appearances, since an assembly has no bodies or faces.
'
'   Sweep B - RemoveMaterialProperty, across all configurations.
'       Document, body, face and feature level, each with swAllConfiguration.
'       This is CodeStack's remove-color pattern. It catches anything scoped by
'       configuration that sweep A's display-state view does not see.
'
' Running both is belt-and-braces on purpose. A macro called Nuke earns the
' redundancy, and the cost is small everywhere except the per-face loop.
'
' NUKE_FACES_INDIVIDUALLY is the expensive one: one COM call per face, and a
' large weldment has thousands. Sweep A already clears face appearances wholesale
' in one call per appearance, so turning this off is the first thing to try if a
' big assembly crawls. Measure it from the diagnostics timing before deciding.
'
' Nothing is saved. SetSaveFlag marks each modified part dirty so SOLIDWORKS
' prompts for it on close; the user chooses whether the change becomes permanent.
' Resolving suppressed components is a change to the assembly and rides on the
' same bargain - close without saving and the suppression states come back.
'
' Components are resolved by calling SetSuppression2 and then RE-READING
' GetSuppression to confirm it took. The call's return enum is not trusted,
' because a resolve that silently failed would leave a part quietly uncleared
' and the summary would still claim success.
'
' Documents are de-duplicated by lower-cased path. A part placed fifty times is
' one document and must be nuked once - doing it fifty times is not wrong, just
' fifty times slower.
'
' EditRebuild3 at the end is not optional. Without it the viewport keeps drawing
' the old colours until the display state is switched away and back.
'------------------------------------------------------------------------------

' Turning this on writes a line per document and per appearance to the VBA
' Immediate window (Ctrl+G in the editor), plus a timing breakdown in the
' completion dialog.
Const SHOW_DIAGNOSTICS As Boolean = False

' The per-face sweep. Slow on large models - see the notes above.
Const NUKE_FACES_INDIVIDUALLY As Boolean = True

' Must match the Version line in the header block above. build-library.ps1 checks
' that they agree and fails the build if they drift.
Const MACRO_VERSION As String = "0.1.0"

Dim swApp As SldWorks.SldWorks

' Running totals, gathered across every document the run touches.
Dim gDocuments As Long
Dim gAppearances As Long
Dim gEntities As Long
Dim gFaces As Long
Dim gResolved As Long
Dim gResolveFails As Long
Dim gModified As Long
Dim gDocFails As Long

'--- Entry point --------------------------------------------------------------

Sub main()
    On Error GoTo mainError

    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly document.", vbCritical
        Exit Sub
    End If

    Dim docType As Long
    docType = swModel.GetType()

    Dim isAssembly As Boolean
    If docType = swDocumentTypes_e.swDocPART Then
        isAssembly = False
    ElseIf docType = swDocumentTypes_e.swDocASSEMBLY Then
        isAssembly = True
    Else
        MsgBox "This macro only works on part or assembly documents.", vbCritical
        Exit Sub
    End If

    If Not Confirmed(isAssembly) Then Exit Sub

    NukeEverything swModel, isAssembly
    Exit Sub

mainError:
    MsgBox "An error occurred in main(): " & Err.Description & _
           " (Error " & Err.Number & ")", vbCritical
End Sub

' Asks before running, and says plainly what is about to happen. This destroys
' work that cannot be recreated and leaves no undo record - and in an assembly it
' changes documents the user never opened, which is the part most worth spelling
' out before the fact rather than after.
Function Confirmed(ByVal isAssembly As Boolean) As Boolean

    Dim prompt As String

    If isAssembly Then
        prompt = "Remove EVERY appearance from this assembly and from every" & vbCrLf & _
                 "part inside it?" & vbCrLf & vbCrLf & _
                 "This clears face, feature, body, part and component" & vbCrLf & _
                 "appearances, in every display state and every" & vbCrLf & _
                 "configuration, including any applied by hand." & vbCrLf & vbCrLf & _
                 "The referenced part files ARE modified. A part used in" & vbCrLf & _
                 "other assemblies is changed for those too." & vbCrLf & vbCrLf & _
                 "Nothing is saved to disk. The parts are left modified and" & vbCrLf & _
                 "unsaved, so closing without saving undoes all of this." & vbCrLf & vbCrLf & _
                 "Suppressed and lightweight components will be resolved." & vbCrLf & vbCrLf & _
                 "There is no undo."
    Else
        prompt = "Remove EVERY appearance from this part?" & vbCrLf & vbCrLf & _
                 "This clears face, feature, body and part appearances," & vbCrLf & _
                 "in every display state and every configuration," & vbCrLf & _
                 "including any applied by hand." & vbCrLf & vbCrLf & _
                 "Nothing is saved to disk. The part is left modified and" & vbCrLf & _
                 "unsaved, so closing without saving undoes all of this." & vbCrLf & vbCrLf & _
                 "There is no undo."
    End If

    Confirmed = (MsgBox(prompt, vbYesNo + vbExclamation, _
                        "Nuke All Appearances") = vbYes)
End Function

'--- Core ---------------------------------------------------------------------

Sub NukeEverything(swModel As SldWorks.ModelDoc2, ByVal isAssembly As Boolean)

    Dim currentStep As String
    Dim swView As SldWorks.ModelView
    On Error GoTo ErrorHandler

    Dim tStart As Single, tResolved As Single, tNuked As Single, tRefreshed As Single
    tStart = Timer

    gDocuments = 0: gAppearances = 0: gEntities = 0: gFaces = 0
    gResolved = 0: gResolveFails = 0: gModified = 0: gDocFails = 0

    If SHOW_DIAGNOSTICS Then
        Debug.Print "=== Nuke All Appearances " & MACRO_VERSION & " ==="
    End If

    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False

    If isAssembly Then
        currentStep = "Resolving components"
        ResolveEverything swModel
    End If
    tResolved = Timer

    currentStep = "Clearing the active document"
    NukeDocument swModel, (Not isAssembly)

    If isAssembly Then
        currentStep = "Clearing referenced documents"
        NukeReferencedDocuments swModel
    End If

    tNuked = Timer

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True

    currentStep = "Rebuilding"
    swModel.EditRebuild3
    tRefreshed = Timer

    Report swModel, isAssembly, tStart, tResolved, tNuked, tRefreshed
    Exit Sub

ErrorHandler:
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MsgBox "ERROR in NukeEverything at step [" & currentStep & "]" & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

' Resolves everything, because an appearance cannot be cleared from a component
' that is not loaded. Lightweight components go in one call; suppressed ones have
' to be done individually.
Sub ResolveEverything(swModel As SldWorks.ModelDoc2)

    On Error Resume Next

    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel

    swAssy.ResolveAllLightWeightComponents True

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long
    Dim swComp As SldWorks.Component2

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If swComp.GetSuppression() = _
           swComponentSuppressionState_e.swComponentSuppressed Then

            swComp.SetSuppression2 swComponentSuppressionState_e.swComponentResolved

            ' Verify the result, not the call's return value. A resolve that
            ' silently failed would leave a part uncleared while the summary
            ' still claimed success.
            If swComp.GetSuppression() = _
               swComponentSuppressionState_e.swComponentSuppressed Then
                gResolveFails = gResolveFails + 1
                If SHOW_DIAGNOSTICS Then
                    Debug.Print "  FAILED to resolve " & swComp.Name2
                End If
            Else
                gResolved = gResolved + 1
            End If
        End If
    Next i
End Sub

' Walks every component at every level, de-duplicates by path, and nukes each
' referenced document once.
Sub NukeReferencedDocuments(swModel As SldWorks.ModelDoc2)

    Dim currentStep As String
    On Error GoTo Failed

    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel

    Dim vComps As Variant
    vComps = swAssy.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    seen.Add LCase$(swModel.GetPathName()), True

    Dim i As Long
    Dim swComp As SldWorks.Component2
    Dim swRefModel As SldWorks.ModelDoc2
    Dim docKey As String

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        Set swRefModel = swComp.GetModelDoc2()

        If Not swRefModel Is Nothing Then
            docKey = LCase$(swRefModel.GetPathName())

            If Not seen.Exists(docKey) Then
                seen.Add docKey, True

                If SHOW_DIAGNOSTICS Then
                    Debug.Print "document: " & docKey
                End If

                NukeDocument swRefModel, _
                    (swRefModel.GetType() = swDocumentTypes_e.swDocPART)

                ' Marks it modified so SOLIDWORKS asks about it on close.
                ' Nothing is written to disk here by design.
                swRefModel.SetSaveFlag
                gModified = gModified + 1
            End If
        End If
    Next i
    Exit Sub

Failed:
    gDocFails = gDocFails + 1
    MsgBox "ERROR walking referenced documents: " & Err.Description & vbCrLf & _
           "Error " & Err.Number, vbCritical
End Sub

' Both sweeps over one document. See the maintenance notes for why there are two.
Sub NukeDocument(swModel As SldWorks.ModelDoc2, ByVal isPart As Boolean)

    On Error Resume Next

    '--- Sweep A: render materials, across all display states ------------------
    Dim vAppearances As Variant
    vAppearances = swModel.Extension.GetRenderMaterials2( _
        swDisplayStateOpts_e.swAllDisplayState, Empty)

    If Not IsEmpty(vAppearances) Then
        Dim i As Long
        Dim swAppearance As SldWorks.RenderMaterial
        Dim vEnts As Variant

        For i = 0 To UBound(vAppearances)
            Set swAppearance = vAppearances(i)

            vEnts = swAppearance.GetEntities
            If Not IsEmpty(vEnts) Then
                gEntities = gEntities + UBound(vEnts) + 1
            End If

            swAppearance.RemoveAllEntities
            gAppearances = gAppearances + 1
        Next i
    End If

    '--- Sweep B: RemoveMaterialProperty, across all configurations -------------
    swModel.Extension.RemoveMaterialProperty _
        swInConfigurationOpts_e.swAllConfiguration, Empty

    If isPart Then
        Dim swPart As SldWorks.PartDoc
        Set swPart = swModel

        Dim vBodies As Variant
        vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, False)

        If Not IsEmpty(vBodies) Then
            Dim b As Long
            Dim swBody As SldWorks.Body2

            For b = 0 To UBound(vBodies)
                Set swBody = vBodies(b)
                swBody.RemoveMaterialProperty _
                    swInConfigurationOpts_e.swAllConfiguration, Empty

                If NUKE_FACES_INDIVIDUALLY Then NukeFaces swBody
            Next b
        End If

        NukeFeatures swModel
    End If

    gDocuments = gDocuments + 1
End Sub

' The expensive sweep: one call per face. Sweep A already clears face appearances
' wholesale, so this is redundancy rather than the main mechanism - see
' NUKE_FACES_INDIVIDUALLY in the maintenance notes.
Sub NukeFaces(swBody As SldWorks.Body2)

    On Error Resume Next

    Dim vFaces As Variant
    vFaces = swBody.GetFaces()
    If IsEmpty(vFaces) Then Exit Sub

    Dim i As Long
    Dim swFace As SldWorks.Face2

    For i = 0 To UBound(vFaces)
        Set swFace = vFaces(i)
        swFace.RemoveMaterialProperty2 _
            swInConfigurationOpts_e.swAllConfiguration, Empty
        gFaces = gFaces + 1
    Next i
End Sub

' Feature-level appearances. Not every feature type carries one, so failures here
' are expected and ignored rather than reported.
Sub NukeFeatures(swModel As SldWorks.ModelDoc2)

    On Error Resume Next

    Dim vFeats As Variant
    vFeats = swModel.FeatureManager.GetFeatures(False)
    If IsEmpty(vFeats) Then Exit Sub

    Dim i As Long
    Dim swFeat As SldWorks.Feature

    For i = 0 To UBound(vFeats)
        Set swFeat = vFeats(i)
        swFeat.RemoveMaterialProperty2 _
            swInConfigurationOpts_e.swAllConfiguration, Empty
    Next i
End Sub

'--- Reporting ----------------------------------------------------------------

Sub Report(swModel As SldWorks.ModelDoc2, ByVal isAssembly As Boolean, _
           ByVal tStart As Single, ByVal tResolved As Single, _
           ByVal tNuked As Single, ByVal tRefreshed As Single)

    Dim msg As String

    If isAssembly Then
        msg = "Nuked " & gAppearances & " appearances across " & _
              gDocuments & " documents." & vbCrLf & vbCrLf & _
              gModified & " referenced documents are now modified and UNSAVED."

        If gResolved > 0 Then
            msg = msg & vbCrLf & gResolved & " suppressed components were resolved."
        End If

        msg = msg & vbCrLf & vbCrLf & _
              "Save them to make this permanent, or close without" & vbCrLf & _
              "saving to put everything back."

        If gResolveFails > 0 Then
            msg = msg & vbCrLf & vbCrLf & _
                  "WARNING: " & gResolveFails & " components could not be" & vbCrLf & _
                  "resolved and still have their appearances."
        End If
    Else
        msg = "Nuked " & gAppearances & " appearances." & vbCrLf & vbCrLf & _
              "The part is modified and UNSAVED. Save to make this" & vbCrLf & _
              "permanent, or close without saving to put it back."
    End If

    msg = msg & vbCrLf & vbCrLf & "Macro Version: " & MACRO_VERSION

    If SHOW_DIAGNOSTICS Then
        msg = msg & vbCrLf & vbCrLf & _
              "Documents processed: " & gDocuments & vbCrLf & _
              "Appearances cleared: " & gAppearances & vbCrLf & _
              "Entities detached: " & gEntities & vbCrLf & _
              "Faces swept: " & gFaces & vbCrLf & _
              "Components resolved: " & gResolved & vbCrLf & _
              "Resolve failures: " & gResolveFails & vbCrLf & _
              "Document failures: " & gDocFails & vbCrLf & vbCrLf & _
              "Resolve: " & Format(tResolved - tStart, "0.00") & " s" & vbCrLf & _
              "Nuke:    " & Format(tNuked - tResolved, "0.00") & " s" & vbCrLf & _
              "Rebuild: " & Format(tRefreshed - tNuked, "0.00") & " s" & vbCrLf & _
              "Total:   " & Format(tRefreshed - tStart, "0.00") & " s"

        Debug.Print "documents=" & gDocuments & _
                    "  appearances=" & gAppearances & _
                    "  entities=" & gEntities & _
                    "  faces=" & gFaces & _
                    "  resolved=" & gResolved & _
                    "  resolveFails=" & gResolveFails
        Debug.Print "=== end ==="
    End If

    MsgBox msg, vbInformation
End Sub
