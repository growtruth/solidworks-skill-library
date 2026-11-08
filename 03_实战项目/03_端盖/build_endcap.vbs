' End cap: OD 90, thickness 10, boss phi60 x 8, bore 30, 4 x phi9 holes on phi72
' Output files are written next to this script.
' Verified volume: 70968.6 mm3 (analytical 70968.6)

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

' 1) Revolved base
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set cl = part.SketchManager.CreateCenterLine(0, -0.010, 0, 0, 0.025, 0)
Set l1 = part.SketchManager.CreateLine(0.015,0,0, 0.045,0,0)
Set l2 = part.SketchManager.CreateLine(0.045,0,0, 0.045,0.010,0)
Set l3 = part.SketchManager.CreateLine(0.045,0.010,0, 0.030,0.010,0)
Set l4 = part.SketchManager.CreateLine(0.030,0.010,0, 0.030,0.018,0)
Set l5 = part.SketchManager.CreateLine(0.030,0.018,0, 0.015,0.018,0)
Set l6 = part.SketchManager.CreateLine(0.015,0.018,0, 0.015,0,0)
part.SketchManager.InsertSketch True
Set rf = part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "S1_REVOLVE vol=" & Round(mp.Volume*1e9,1) & " expect=73513.3"

' 2) Four phi9 through holes on phi72
part.ClearSelection2 True
part.Extension.SelectByID2 pn2, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Dim i, a, x, z
For i = 0 To 3
  a = i * 3.14159265358979 / 2
  x = 0.036 * Cos(a)
  z = 0.036 * Sin(a)
  Set cir = part.SketchManager.CreateCircle(x, z, 0, x + 0.0045, z, 0)
Next
part.SketchManager.InsertSketch True
Set cf = part.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
Set mp = part.Extension.CreateMassProperty2()
If cf Is Nothing Then
  WScript.Echo "S2_HOLES FAIL vol=" & Round(mp.Volume*1e9,1)
Else
  WScript.Echo "S2_HOLES vol=" & Round(mp.Volume*1e9,1) & " expect=70968.6"
End If

part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\endcap_D90_d30.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\endcap_D90_d30.STEP", 0, False, False
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\endcap_D90_d30.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\endcap_D90_d30.STEP")
