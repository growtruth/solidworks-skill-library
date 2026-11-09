' Flange R2: revolved base with 2x45 deg section chamfer, single hole + 6 instance circular pattern
' Output files are written next to this script.
' Verified volumes: base 180092.9 | one hole 177630.0 | patterned 165315.5 mm3

Set fso = CreateObject("Scripting.FileSystemObject")
Set sw = CreateObject("SldWorks.Application")
sw.Visible = True
sw.UserControl = False

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

' 1) Revolve base, outer rim already chamfered 2x45 in the section
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set cl = part.SketchManager.CreateCenterLine(0, -0.010, 0, 0, 0.030, 0)
Set l1 = part.SketchManager.CreateLine(0.020,0,0, 0.058,0,0)
Set l2 = part.SketchManager.CreateLine(0.058,0,0, 0.060,0.002,0)
Set l3 = part.SketchManager.CreateLine(0.060,0.002,0, 0.060,0.014,0)
Set l4 = part.SketchManager.CreateLine(0.060,0.014,0, 0.058,0.016,0)
Set l5 = part.SketchManager.CreateLine(0.058,0.016,0, 0.035,0.016,0)
Set l6 = part.SketchManager.CreateLine(0.035,0.016,0, 0.035,0.024,0)
Set l7 = part.SketchManager.CreateLine(0.035,0.024,0, 0.020,0.024,0)
Set l8 = part.SketchManager.CreateLine(0.020,0.024,0, 0.020,0,0)
part.SketchManager.InsertSketch True
Set rf = part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "S1_REVOLVE vol=" & Round(mp.Volume*1e9,1) & " expect=180092.9"

' 2) Single hole phi14 at r=45, sketch on top plane uses (x, z, 0)
part.ClearSelection2 True
part.Extension.SelectByID2 pn2, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set c1 = part.SketchManager.CreateCircle(0.045, 0, 0, 0.052, 0, 0)
part.SketchManager.InsertSketch True
Set cf = part.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
Set mp = part.Extension.CreateMassProperty2()
If cf Is Nothing Then
  WScript.Echo "S2_HOLE FAIL vol=" & Round(mp.Volume*1e9,1)
  sw.CloseDoc part.GetTitle()
  sw.ExitApp
  WScript.Quit
End If
WScript.Echo "S2_HOLE vol=" & Round(mp.Volume*1e9,1) & " expect=177630.0"

' 3) Circular pattern: seed feature mark 4, direction axis (circular edge) mark 1
part.ClearSelection2 True
okf = part.Extension.SelectByID2(cf.Name, "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
oke = part.Extension.SelectByID2("", "EDGE", 0.020, 0.024, 0.0, True, 1, Nothing, 0)
WScript.Echo "SEL_FEATURE=" & okf & " SEL_AXIS=" & oke
Set pat = part.FeatureManager.FeatureCircularPattern4(6, 6.2831853071796, False, "NULL", False, True, False)
Set mp = part.Extension.CreateMassProperty2()
If pat Is Nothing Then
  WScript.Echo "S3_PATTERN FAIL vol=" & Round(mp.Volume*1e9,1)
Else
  WScript.Echo "S3_PATTERN vol=" & Round(mp.Volume*1e9,1) & " expect=165315.5 name=" & pat.Name
End If
WScript.Echo "FEATURES=" & part.GetFeatureCount()

' 4) Save and exit to release memory
part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\flange_r2_D120_d40.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\flange_r2_D120_d40.STEP", 0, False, False
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\flange_r2_D120_d40.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\flange_r2_D120_d40.STEP")
sw.CloseDoc part.GetTitle()
sw.ExitApp
WScript.Echo "SW_EXITED"
