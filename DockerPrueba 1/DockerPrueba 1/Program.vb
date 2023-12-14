Imports System
Imports Autodesk.AutoCAD.Interop

Module Program

    Dim acadApp As New AcadApplication

    Sub Main(args As String())
        acadApp.Visible = True ' Optionally, set AutoCAD to be visible

        ' Call the boquillaSWsol function with sample values
        boquillaSWsol(0, 5, 10, 15, 20, 0, 5, 10, 15, 20)

        Console.WriteLine("Function executed successfully.")
    End Sub

    Function boquillaSWsol(x0 As Double, x1 As Double, x2 As Double, x3 As Double, x4 As Double, y0 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double) As Object
        Dim pIni(0 To 2) As Double
        Dim pFin(0 To 2) As Double
        Dim rect(0 To 0) As AcadEntity
        Dim region As AcadEntity
        Dim aux3D As AcadEntity
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
        region = acadApp.ActiveDocument.ModelSpace.AddRegion(rect)
        aux3D = acadApp.ActiveDocument.ModelSpace.AddRevolvedSolid(region, pIni, pFin, ang)

        rect.Delete()
        region.Delete()

        Return aux3D
    End Function
End Module
