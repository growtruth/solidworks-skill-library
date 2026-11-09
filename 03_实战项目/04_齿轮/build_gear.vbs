' Gear R4: module 2, teeth 20, pressure angle 20 deg, face width 10 mm, bore 10 mm
' Involute flanks approximated by polylines; volume verified by polygon area x width.
' Output files are written next to this script.

Const PI = 3.14159265358979
Const M = 2
Const Z = 20
Const ALPHA = 0.349065850398866
Const WIDTH = 10
Const BORE_R = 5

Dim rp, rb, ra, rf, invA, pitch, half, phiB, phiA
rp = M * Z / 2
rb = rp * Cos(ALPHA)
ra = rp + M
rf = rp - 1.25 * M
invA = Tan(ALPHA) - ALPHA
pitch = 2 * PI / Z
half = PI / (2 * Z)
phiB = half + invA
Dim xA, tA
xA = rb / ra
tA = Atn(-xA / Sqr(-xA * xA + 1)) + 2 * Atn(1)
phiA = half + invA - (Tan(tA) - tA)

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
pn1 = "" : idx = 0
Do While Not fe Is Nothing
  If fe.GetTypeName2() = "RefPlane" Then
    idx = idx + 1
    If idx = 1 Then pn1 = fe.Name
  End If
  Set fe = fe.GetNextFeature()
Loop

part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True

Dim area, px, py, nx, ny, k, i, r, ang, th, NF, NT, NR
Dim xx, tt, phi
NF = 10
NT = 6
NR = 6
area = 0
px = 0 : py = 0
first = True

For k = 0 To Z - 1
  th = k * pitch
  nx = rf * Cos(th - phiB)
  ny = rf * Sin(th - phiB)
  If first Then
    px = nx : py = ny
    first = False
  Else
    part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
    area = area + (px * ny - nx * py)
    px = nx : py = ny
  End If
  nx = rb * Cos(th - phiB)
  ny = rb * Sin(th - phiB)
  part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
  area = area + (px * ny - nx * py)
  px = nx : py = ny
  For i = 1 To NF - 1
    r = rb + (ra - rb) * i / (NF - 1)
    xx = rb / r
    tt = Atn(-xx / Sqr(-xx * xx + 1)) + 2 * Atn(1)
    phi = half + invA - (Tan(tt) - tt)
    ang = th - phi
    nx = r * Cos(ang)
    ny = r * Sin(ang)
    part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
    area = area + (px * ny - nx * py)
    px = nx : py = ny
  Next
  For i = 1 To NT
    ang = (th - phiA) + 2 * phiA * i / NT
    nx = ra * Cos(ang)
    ny = ra * Sin(ang)
    part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
    area = area + (px * ny - nx * py)
    px = nx : py = ny
  Next
  For i = NF - 1 To 1 Step -1
    r = rb + (ra - rb) * i / (NF - 1)
    xx = rb / r
    tt = Atn(-xx / Sqr(-xx * xx + 1)) + 2 * Atn(1)
    phi = half + invA - (Tan(tt) - tt)
    ang = th + phi
    nx = r * Cos(ang)
    ny = r * Sin(ang)
    part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
    area = area + (px * ny - nx * py)
    px = nx : py = ny
  Next
  nx = rf * Cos(th + phiB)
  ny = rf * Sin(th + phiB)
  part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
  area = area + (px * ny - nx * py)
  px = nx : py = ny
  For i = 1 To NR
    ang = (th + phiB) + (pitch - 2 * phiB) * i / NR
    nx = rf * Cos(ang)
    ny = rf * Sin(ang)
    part.SketchManager.CreateLine px/1000, py/1000, 0, nx/1000, ny/1000, 0
    area = area + (px * ny - nx * py)
    px = nx : py = ny
  Next
Next

part.SketchManager.InsertSketch True
area = Abs(area) / 2
WScript.Echo "PROFILE_AREA_MM2=" & Round(area, 1)
WScript.Echo "EXPECT_VOL_MM3=" & Round((area - PI * BORE_R * BORE_R) * WIDTH, 1)

Set ex = part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, WIDTH/1000, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)
If ex Is Nothing Then WScript.Echo "FAIL: extrude"
Set mp0 = part.Extension.CreateMassProperty2()
WScript.Echo "EXTRUDE_VOL_MM3=" & Round(mp0.Volume*1e9,1)

part.ClearSelection2 True
part.Extension.SelectByID2 pn1, "PLANE", 0, 0, 0, False, 0, Nothing, 0
part.SketchManager.InsertSketch True
Set c = part.SketchManager.CreateCircle(0, 0, 0, BORE_R/1000, 0, 0)
part.SketchManager.InsertSketch True
Set cf = part.FeatureManager.FeatureCut4(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
If cf Is Nothing Then WScript.Echo "FAIL: bore cut"

Set mp = part.Extension.CreateMassProperty2()
WScript.Echo "ACTUAL_VOL_MM3=" & Round(mp.Volume*1e9,1)

WScript.Echo "OD_EXPECT=44 FACE_EXPECT=10"

part.ViewZoomtofit2
outDir = fso.GetParentFolderName(WScript.ScriptFullName)
part.SaveAs2 outDir & "\gear_m2_z20.SLDPRT", 0, False, False
part.SaveAs2 outDir & "\gear_m2_z20.STEP", 0, False, False
WScript.Echo "PART_EXISTS=" & fso.FileExists(outDir & "\gear_m2_z20.SLDPRT")
WScript.Echo "STEP_EXISTS=" & fso.FileExists(outDir & "\gear_m2_z20.STEP")
sw.CloseDoc part.GetTitle()
sw.ExitApp
WScript.Echo "SW_EXITED"
