Sub OptimizePipeAndTank()
    ' Declare And initialize variables
    Dim D As Variant, Q As Double, v1 As Double, v2 As Double, Re As Double, Lambda As Double
    Dim hf As Double, PipeCost As Double, BalancingStorage As Double, ExtraStorage As Double
    Dim TotalStorage As Double, StorageCost As Double, TotalCost As Double
    Dim PipeUnitCost As Variant, PipeDiameter As Variant, PipeLength As Double, Viscosity As Double
    Dim Gravity As Double, Pi As Double, StorageUnitCost As Double, Roughness As Double
    Dim EmergencyHours As Double, BottomHours As Double, MaxHeadloss As Double
    Dim Time As Variant, VolumeDemanded As Variant, TankVolume As Double

    ' Transfer input data from "Data" worksheet
    With Worksheets("Data")
        PipeUnitCost = .Range("G2:G7").Value
        PipeDiameter = .Range("H2:H7").Value
        PipeLength = .Cells(3, 11).Value
        Viscosity = .Cells(2, 11).Value
        Gravity = .Cells(4, 11).Value
        Pi = .Cells(5, 11).Value
        StorageUnitCost = .Cells(6, 11).Value
        Roughness = .Cells(7, 11).Value
        EmergencyHours = .Cells(8, 11).Value
        BottomHours = .Cells(9, 11).Value
        MaxHeadloss = .Cells(10, 11).Value
        Time = .Range("B3:B170").Value
        VolumeDemanded = .Range("C3:C170").Value
    End With

    ' Create worksheet "Balancing storage calculations"
    Worksheets.Add.Name = "Balancing storage calculations"

    ' Loop through each pipe diameter
    For Each D In PipeDiameter
        ' Initialize tank volume
        TankVolume = 0

        ' Add headers To "Balancing storage calculations" worksheet
        Worksheets("Balancing storage calculations").Cells(1, 1).Value = "Time"
        Worksheets("Balancing storage calculations").Cells(1, 2).Value = "Volume Demanded"
        Worksheets("Balancing storage calculations").Cells(1, 3).Value = "Tank Volume"
        Worksheets("Balancing storage calculations").Cells(1, 4).Value = "Pipe Supply (Q)"

        ' Loop through each hour
        For i = 1 To UBound(Time)


            ' Assume initial velocity
            v1 = 1

            ' Calculate Pipe supply (Q)
            Do
                ' Calculate Reynolds number
                Re = (v1 * D) / Viscosity

                ' Calculate pipe headloss coefficient (Lambda)
                Dim LogCalc As Double
                LogCalc = Log((Roughness / (3.7 * D)) + (5.74 / (Re ^ 0.9))) * Log(10)
                Lambda = (1 / (-2 * LogCalc)) ^ 2

                ' Calculate New velocity
                v2 = (2 * Gravity * D * MaxHeadloss / (Lambda * PipeLength)) ^ 0.5

                ' Update v1 If Not converged
                If v1 <> v2 Then
                    v1 = v2
                End If
            Loop Until v1 = v2

            ' Calculate Pipe supply (Q)
            Q = v2 * (Pi * D ^ 2 / 4)

            ' Calculate Tank volume
            TankVolume = TankVolume + (Q * 3600) - VolumeDemanded(i, 1)  ' Convert Q from m^3/s To m^3/hr

            ' Write data To "Balancing storage calculations" worksheet
            Worksheets("Balancing storage calculations").Cells(i + 1, 1).Value = CStr(Time(i, 1))
            Worksheets("Balancing storage calculations").Cells(i + 1, 2).Value = VolumeDemanded(i, 1)
            Worksheets("Balancing storage calculations").Cells(i + 1, 3).Value = TankVolume
            Worksheets("Balancing storage calculations").Cells(i + 1, 4).Value = Q * 3600 ' Convert Q from m^3/s To m^3/hr
        Next i

        ' Create worksheet "Optimization"
        Worksheets.Add.Name = "Optimization"

        ' Add headers To "Optimization" worksheet
        Worksheets("Optimization").Cells(1, 1).Value = "Pipe Diameter (D)"
        Worksheets("Optimization").Cells(1, 2).Value = "Pipe Unit Cost"
        Worksheets("Optimization").Cells(1, 3).Value = "Pipe Supply (Q)"
        Worksheets("Optimization").Cells(1, 4).Value = "Pipe Headloss"
        Worksheets("Optimization").Cells(1, 5).Value = "Pipe Cost"
        Worksheets("Optimization").Cells(1, 6).Value = "Balancing Storage"
        Worksheets("Optimization").Cells(1, 7).Value = "Extra Storage"
        Worksheets("Optimization").Cells(1, 8).Value = "Total Storage"
        Worksheets("Optimization").Cells(1, 9).Value = "Storage Cost"
        Worksheets("Optimization").Cells(1, 10).Value = "Total Cost"

        ' Write data To "Optimization" worksheet
        Worksheets("Optimization").Cells(2, 1).Value = D
        Dim matchResult As Variant
        matchResult = Application.Match(D, PipeDiameter, 0)

        If IsError(matchResult) Then
            Worksheets("Optimization").Cells(2, 2).Value = 0  ' Assign a default value If no match is found
        Else
            Worksheets("Optimization").Cells(2, 2).Value = PipeUnitCost(matchResult)
        End If
        Worksheets("Optimization").Cells(2, 3).Value = Q
        Worksheets("Optimization").Cells(2, 4).Value = hf
        Worksheets("Optimization").Cells(2, 5).Value = PipeCost
        Worksheets("Optimization").Cells(2, 6).Value = BalancingStorage
        Worksheets("Optimization").Cells(2, 7).Value = ExtraStorage
        Worksheets("Optimization").Cells(2, 8).Value = TotalStorage
        Worksheets("Optimization").Cells(2, 9).Value = StorageCost
        Worksheets("Optimization").Cells(2, 10).Value = TotalCost

        ' Calculate additional optimization parameters
        Worksheets("Optimization").Cells(2, 4).Value = Lambda * PipeLength * v2 ^ 2 / (2 * Gravity * D) ' Pipe headloss
        Worksheets("Optimization").Cells(2, 5).Value = PipeLength * PipeUnitCost(Application.Match(D, PipeDiameter, 0)) ' Pipe cost
        Worksheets("Optimization").Cells(2, 6).Value = Application.WorksheetFunction.Max(Worksheets("Balancing storage calculations").Range("C2:C195")) + Application.WorksheetFunction.Max(-Worksheets("Balancing storage calculations").Range("C2:C195")) ' Balancing storage
        Worksheets("Optimization").Cells(2, 7).Value = EmergencyHours * Application.WorksheetFunction.Average(VolumeDemanded) + BottomHours * Application.WorksheetFunction.Average(VolumeDemanded) ' Extra storage
        Worksheets("Optimization").Cells(2, 8).Value = Worksheets("Optimization").Cells(2, 6).Value + Worksheets("Optimization").Cells(2, 7).Value ' Total storage
        Worksheets("Optimization").Cells(2, 9).Value = Worksheets("Optimization").Cells(2, 8).Value * StorageUnitCost ' Storage cost
        Worksheets("Optimization").Cells(2, 10).Value = Worksheets("Optimization").Cells(2, 9).Value + Worksheets("Optimization").Cells(2, 5).Value ' Total cost

        ' Find row With minimum Total Cost
        Dim OptimumRow As Long
        OptimumRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Worksheets("Optimization").Range("J2:J" & Worksheets("Optimization").Cells(Rows.Count, "J").End(xlUp).Row)), Worksheets("Optimization").Range("J2:J" & Worksheets("Optimization").Cells(Rows.Count, "J").End(xlUp).Row), 0)

        ' Highlight row With optimum solution
        If Worksheets("Optimization").Cells(OptimumRow, 4).Value <= MaxHeadloss Then
            Worksheets("Optimization").Rows(OptimumRow).Interior.ColorIndex = 4 ' Green color
        Else
            Worksheets("Optimization").Cells(2, 10).Value = "The maximum headloss constraint (hfmax) was violated And no optimum pipe And tank combination can be calculated."
        End If

        ' Delete "Optimization" worksheet For Next iteration
        Application.DisplayAlerts = False
        Worksheets("Optimization").Delete
        Application.DisplayAlerts = True
    Next D

    ' Delete "Balancing storage calculations" worksheet
    Application.DisplayAlerts = False
    Worksheets("Balancing storage calculations").Delete
    Application.DisplayAlerts = True


End Sub

