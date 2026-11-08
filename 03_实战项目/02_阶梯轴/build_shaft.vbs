' Stepped shaft: phi25/40/30/20 segments, keyway 12 x 5 x 26 on the phi40 section
' Output files are written next to this script.
' Verified volume: 116541.0 mm3 (analytical 116540.7)

Set fso = CreateObject("Scripting.FileSystemObject")
Set sw = CreateObject("SldWorks.Application")
sw.Visible = True

tmpl = "C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\gb_part.prtdot"
Set part = sw.NewDocument(tmpl, 0, 0, 0)
If part Is Nothing Then
  WScript.Echo "FAIL: cannot create new document"
  WScript.Quit
End If

Set fe = part.FirstFeature()
pn1 = "" : pn2 = "" : idx = 0
Do While Not fe Is Nothing
  If fe.GetTypeName2() = "RefPlane" Then
    idx = idx + 1
    If idx = 1 Then pn1 = fe.Name
    If idx = 2 Then pn2 = fe.Name
  End If
  Set fe = fe.GetNextFeature()
Loop

' 1) Revolve around a horizontal centerline
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set cl = part.SketchManager.CreateCenterLine(-0.010, 0, 0, 0.180, 0, 0)
Set l1 = part.SketchManager.CreateLine(0,0,0, 0.170,0,0)
Set l2 = part.SketchManager.CreateLine(0.170,0,0, 0.170,0.010,0)
Set l3 = part.SketchManager.CreateLine(0.170,0.010,0, 0.150,0.010,0)
Set l4 = part.SketchManager.CreateLine(0.150,0.010,0, 0.150,0.015,0)
Set l5 = part.SketchManager.CreateLine(0.150,0.015,0, 0.080,0.015,0)
Set l6 = part.SketchManager.CreateLine(0.080,0.015,0, 0.080,0.020,0)
Set l7 = part.SketchManager.CreateLine(0.080,0.020,0, 0.050,0.020,0)
Set l8 = part.SketchManager.CreateLine(0.050,0.020,0, 0.050,0.0125,0)
Set l9 = part.SketchManager.CreateLine(0.050,0.0125,0, 0,0.0125,0)
Set l10 = part.SketchManager.CreateLine(0,0.0125,0, 0,0,0)
part.SketchManager.InsertSketch True
Set rf = part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "S1_REVOLVE vol=" & Round(mp.Volume*1e9,1) & " expect=118005.8"

' 2) Keyway: rectangle on front plane, double ended blind cut 6 mm each side
part.ClearSelection2 True
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set rc = part.SketchManager.CreateCornerRectangle(0.052, 0.015, 0, 0.078, 0.025, 0)
part.SketchManager.InsertSketch True
Set cf = part.FeatureManager.FeatureCut4(False, False, True, 0, 0, 0.006, 0.006, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
Set mp = part.Extension.CreateMassProperty2()
If cf Is Nothing Then
  WScript.Echo "S2_KEYWAY FAIL vol=" & Round(mp.Volume*1e9,1)
Else
  WScript.Echo "S2_KEYWAY vol=" & Round(mp.Volume*1e9,1) & " expect=116540.7"
End If

part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\shaft_steps.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\shaft_steps.STEP", 0, False, False
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\shaft_steps.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\shaft_steps.STEP")
