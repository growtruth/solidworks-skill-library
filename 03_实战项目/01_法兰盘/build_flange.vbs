' Flange: OD 120, bore 40, flange thickness 16, boss phi70 x 8, 6 x phi14 holes on phi90
' Output files are written next to this script.
' Verified volume: 166806.0 mm3 (analytical 166806.7)

Set fso = CreateObject("Scripting.FileSystemObject")
Set sw = CreateObject("SldWorks.Application")
sw.Visible = True

tmpl = "C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\gb_part.prtdot"
Set part = sw.NewDocument(tmpl, 0, 0, 0)
If part Is Nothing Then
  WScript.Echo "FAIL: cannot create new document"
  WScript.Quit
End If

' Locate the first two reference planes without hardcoding localized names
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

' 1) Revolved base: half section in mm, API units are meters
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set cl = part.SketchManager.CreateCenterLine(0, -0.010, 0, 0, 0.030, 0)
Set l1 = part.SketchManager.CreateLine(0.020,0,0, 0.060,0,0)
Set l2 = part.SketchManager.CreateLine(0.060,0,0, 0.060,0.016,0)
Set l3 = part.SketchManager.CreateLine(0.060,0.016,0, 0.035,0.016,0)
Set l4 = part.SketchManager.CreateLine(0.035,0.016,0, 0.035,0.024,0)
Set l5 = part.SketchManager.CreateLine(0.035,0.024,0, 0.020,0.024,0)
Set l6 = part.SketchManager.CreateLine(0.020,0.024,0, 0.020,0,0)
part.SketchManager.InsertSketch True
Set rf = part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "S1_REVOLVE vol=" & Round(mp.Volume*1e9,1) & " expect=181584.1"

' 2) Six phi14 through holes, sketch on top plane uses (x, z, 0)
part.ClearSelection2 True
part.Extension.SelectByID2 pn2, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Dim i, a, x, z
For i = 0 To 5
  a = i * 3.14159265358979 / 3
  x = 0.045 * Cos(a)
  z = 0.045 * Sin(a)
  Set cir = part.SketchManager.CreateCircle(x, z, 0, x + 0.007, z, 0)
Next
part.SketchManager.InsertSketch True
' Cut default direction is opposite the sketch normal, so Dir must be True here
Set cf = part.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
Set mp = part.Extension.CreateMassProperty2()
If cf Is Nothing Then
  WScript.Echo "S2_HOLES FAIL vol=" & Round(mp.Volume*1e9,1)
  WScript.Quit
End If
WScript.Echo "S2_HOLES vol=" & Round(mp.Volume*1e9,1) & " expect=166806.7"

' 3) Save next to this script
part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\flange_D120_d40.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\flange_D120_d40.STEP", 0, False, False
' SaveAs2 return value is unreliable in VBScript, verify by file existence
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\flange_D120_d40.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\flange_D120_d40.STEP")
