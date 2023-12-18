Imports System
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module Program

    Dim acadApp As New AcadApplication

    Sub Main(args As String())
        acadApp.Visible = True ' Optionally, set AutoCAD to be visible

        ' Call the boquillaSWsol function with sample values
        boquillaSWsol(0, 5, 10, 15, 20, 0, 5, 10, 15, 20)
        'tapaTorCorte(1, 0.5, 0.2, 0.1, 0.05, 0.1, 45, 0.02, 1, 0.15)

        'coupling(1, 2, 3, 4, 5, 6, 7, "Sólido")

        Console.WriteLine("Function executed successfully.")
    End Sub

    Function boquillaSWsol(x0 As Double, x1 As Double, x2 As Double, x3 As Double, x4 As Double, y0 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Object
        Dim pIni(0 To 2) As Double
        Dim pFin(0 To 2) As Double
        Dim rect As AcadLWPolyline
        Dim region As AcadRegion
        Dim aux3D As Acad3DSolid
        Dim pF(0 To 21) As Double
        Dim pii As Double = Math.PI
        Dim ang As Double = 2 * pii
        Dim z0 As Double = 0 ' Assuming z0 needs to be defined

        ' 1. Creación de la boquilla
        pF(0) = x0 : pF(1) = y0
        pF(2) = x0 : pF(3) = y1
        pF(4) = x1 : pF(5) = y1
        pF(6) = x1 : pF(7) = y2
        pF(8) = x2 : pF(9) = y2
        pF(10) = x2 : pF(11) = y4
        pF(12) = x3 : pF(13) = y4
        pF(14) = x3 : pF(15) = y3
        pF(16) = x4 : pF(17) = y3
        pF(18) = x4 : pF(19) = y0
        pF(20) = x0 : pF(21) = y0

        pIni(0) = x0 : pIni(1) = y0 : pIni(2) = z0
        pFin(0) = 1 : pFin(1) = y0 : pFin(2) = z0

        rect = acadApp.ActiveDocument.ModelSpace.AddLightWeightPolyline(pF)


        region = acadApp.ActiveDocument.ModelSpace.AddRegion(rect.Document.)                                                      )



        'region = acadApp.ActiveDocument.ModelSpace.AddRegion(rect)
        aux3D = acadApp.ActiveDocument.ModelSpace.AddRevolvedSolid(region, pIni, pFin, ang)

        'rect.Delete()
        'region.Delete()

        Return aux3D
    End Function



    Function tapaTorCorte(rAb, rCu, By, Cy, tc, aX, alfa, fPes, D, riBC)
        Dim pIni(0 To 2) As Double
        Dim pFin(0 To 2) As Double
        Dim pCap(0 To 2) As AcadEntity
        Dim pF(0 To 7) As Double
        Dim rect(0 To 0) As AcadEntity
        Dim pii As Double = Math.PI
        Dim ang As Double = 2 * pii
        Dim lineaH As AcadEntity
        Dim arcoAI As AcadEntity
        Dim arcoTI As AcadEntity
        Dim region As AcadEntity
        Dim aux3D As AcadEntity

        ' 1. Cierre de forma linea Horizontal
        Dim y2 As Double = Cy + fPes
        Dim y1 As Double = Cy
        Dim y0 As Double = 0
        Dim x0 As Double = -rCu / 2
        Dim x1 As Double = 0

        pF(0) = x0 : pF(1) = y1
        pF(2) = x0 : pF(3) = y2
        pF(4) = x1 : pF(5) = y2
        pF(6) = x1 : pF(7) = y0

        lineaH = acadApp.ActiveDocument.ModelSpace.AddLightWeightPolyline(pF)

        Dim z1 As Double = Cy
        Dim x1a As Double = -aX
        pIni(0) = x1a : pIni(1) = y1 : pIni(2) = z1
        Dim angI1 As Double = pii
        Dim angF1 As Double = pii + alfa
        arcoAI = acadApp.ActiveDocument.ModelSpace.AddArc(pIni, rAb, angI1, angF1)

        ' 5. Arco de cap Interno
        z1 = By + Cy
        x1 = 0
        pIni(0) = x1 : pIni(1) = y1 : pIni(2) = z1
        Dim angI As Double = pii + alfa
        Dim angF As Double = 270 / 180 * pii
        arcoTI = acadApp.ActiveDocument.ModelSpace.AddArc(pIni, rCu, angI, angF)

        ' 4.2.3.1 Convertir líneas en region
        pCap(0) = lineaH
        pCap(1) = arcoAI
        pCap(2) = arcoTI

        region = acadApp.ActiveDocument.ModelSpace.AddRegion(pCap)
        lineaH.Delete()
        arcoAI.Delete()
        arcoTI.Delete()

        'pIni(0) = x2C : pIni(1) = 0 : pIni(2) = 0
        pFin(0) = 0 : pFin(1) = 0 : pFin(2) = 1

        aux3D = acadApp.ActiveDocument.ModelSpace.AddRevolvedSolid(region, pIni, pFin, ang)
        region.Delete()
    End Function
End Module
