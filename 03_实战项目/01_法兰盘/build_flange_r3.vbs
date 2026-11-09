' Flange R3: R2 geometry + material assignment experiment (material API does not work here)
' Output files are written next to this script.
' Verified volume: 165314.8 mm3; analytical mass for plain carbon steel (7.845 g/cm3) = 1296.9 g

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

' 1) Revolve base with 2x45 deg section chamfer
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
part.FeatureManager.FeatureRevolve2 True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True

' 2) Single hole + circular pattern
part.ClearSelection2 True
part.Extension.SelectByID2 pn2, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set c1 = part.SketchManager.CreateCircle(0.045, 0, 0, 0.052, 0, 0)
part.SketchManager.InsertSketch True
Set cf = part.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
If cf Is Nothing Then
  WScript.Echo "FAIL: hole cut"
  sw.CloseDoc part.GetTitle()
  sw.ExitApp
  WScript.Quit
End If
part.ClearSelection2 True
okf = part.Extension.SelectByID2(cf.Name, "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
oke = part.Extension.SelectByID2("", "EDGE", 0.020, 0.024, 0.0, True, 1, Nothing, 0)
Set pat = part.FeatureManager.FeatureCircularPattern4(6, 6.2831853071796, False, "NULL", False, True, False)
If pat Is Nothing Then WScript.Echo "FAIL: circular pattern"

' 3) Material assignment attempts
On Error Resume Next
db = "D:\FlashGo\sw2025\SOLIDWORKS Corp 2025\SOLIDWORKS\lang\chinese-simplified\sldmaterials\solidworks materials.sldmat"
cfg = part.ConfigurationManager.ActiveConfiguration.Name
' Material name "Plain Carbon Steel" built with ChrW so this file stays ASCII
matName = ChrW(26222) & ChrW(36890) & ChrW(30923) & ChrW(38050)
part.SetMaterialPropertyName2 cfg, db, matName
part.EditRebuild3
WScript.Echo "MATERIAL_USER=" & part.MaterialUserName
Set mpA = part.Extension.CreateMassProperty2()
WScript.Echo "A_DENSITY=" & mpA.Density

Set bodies = part.GetBodies2(0, True)
For Each b In bodies
  b.SetMaterialProperty "solidworks materials.sldmat", matName
Next
Set mpB = part.Extension.CreateMassProperty2()
WScript.Echo "B_DENSITY=" & mpB.Density

' 4) Mass properties
Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "VOLUME_MM3=" & Round(mp.Volume*1e9,1)
WScript.Echo "DENSITY=" & mp.Density
WScript.Echo "MASS=" & mp.Mass

' 5) Save and exit
part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\flange_r3_D120_d40.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\flange_r3_D120_d40.STEP", 0, False, False
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\flange_r3_D120_d40.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\flange_r3_D120_d40.STEP")
sw.CloseDoc part.GetTitle()
sw.ExitApp
WScript.Echo "SW_EXITED"
