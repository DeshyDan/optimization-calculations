Sub OptimizePipeAndTank()
    ' Declare And initialize variables
    Dim D As Variant, Q As Double, v1 As Double, v2 As Double, Re As Double, Lambda As Double
    Dim hf As Double, PipeCost As Double, BalancingStorage As Double, ExtraStorage As Double
    Dim TotalStorage As Double, StorageCost As Double, TotalCost As Double
    Dim PipeUnitCost As Variant, PipeDiameter As Variant, PipeLength As Double, Viscosity As Double
    Dim Gravity As Double, Pi As Double, StorageUnitCost As Double, Roughness As Double
    Dim EmergencyHours As Double, BottomHours As Double, MaxHeadloss As Double
    Dim Time As Variant, VolumeDemanded As Variant, TankVolume As Double

    'Create worksheet "Balancing storage calculations"
    Worksheets.Add.Name = "Balancing storage calculations"

    ' Add headers To "Balancing storage calculations" worksheet
    Worksheets("Balancing storage calculations").Cells(1, 1).Value = "Time"
    Worksheets("Balancing storage calculations").Cells(1, 2).Value = "Volume Demanded"
    Worksheets("Balancing storage calculations").Cells(1, 3).Value = "Tank Volume"
    Worksheets("Balancing storage calculations").Cells(1, 4).Value = "Pipe Supply (Q)"


    ' Create worksheet "Optimization"
    Worksheets.Add.Name = "Optimization"

    ' Add headers To "Optimization" worksheet
    Worksheets("Optimization").Cells(1, 1).Value = "Pipe Diameter (m)"
    Worksheets("Optimization").Cells(1, 2).Value = "Pipe Unit Cost (Rand/m)"
    Worksheets("Optimization").Cells(1, 3).Value = "Pipe Supply (m^3/s)"
    Worksheets("Optimization").Cells(1, 4).Value = "Pipe Headloss (m)"
    Worksheets("Optimization").Cells(1, 5).Value = "Pipe Cost (Rand)"
    Worksheets("Optimization").Cells(1, 6).Value = "Balancing Storage (m^3)"
    Worksheets("Optimization").Cells(1, 7).Value = "Extra Storage (m^3)"
    Worksheets("Optimization").Cells(1, 8).Value = "Total Storage (m^3)"
    Worksheets("Optimization").Cells(1, 9).Value = "Storage Cost (Rand)"
    Worksheets("Optimization").Cells(1, 10).Value = "Total Cost (Rand)"


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
    Dim k As Integer

    k = 1
    ' Loop through each pipe diameter
    For Each D In PipeDiameter
        ' Initialize tank volume
        TankVolume = 0



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
            With Worksheets("Balancing storage calculations").Cells(i + 1, 1)
                .Value = CStr(Time(i, 1))
                .NumberFormat = "hh:mm"
            End With
            Worksheets("Balancing storage calculations").Cells(i + 1, 2).Value = VolumeDemanded(i, 1)
            Worksheets("Balancing storage calculations").Cells(i + 1, 3).Value = TankVolume
            Worksheets("Balancing storage calculations").Cells(i + 1, 4).Value = Q * 3600 ' Convert Q from m^3/s To m^3/hr
        Next i


        ' Calculate pipe headloss
        hf = (Lambda * PipeLength * v2 ^ 2) / (2 * Gravity * D)

        ' Check If headloss constraint is violated
        If hf > MaxHeadloss Then
            MsgBox "The maximum headloss constraint (h_fmax) was violated For pipe diameter " & D & " m And no optimum pipe And tank combination can be calculated."

        End If

        ' Calculate pipe cost
        Dim pipeUnitCostValue As Double
        Dim pipeDiameterValue As Double

        ' Find the row number of the current pipe diameter in the "Data" worksheet
        Dim rowNum As Long
        rowNum = Application.WorksheetFunction.Match(D, Worksheets("Data").Range("H2:H7"), 0) + 1

        ' Get the corresponding pipe unit cost And pipe diameter values
        pipeUnitCostValue = Worksheets("Data").Cells(rowNum, 7).Value
        pipeDiameterValue = D

        PipeCost = PipeLength * pipeUnitCostValue

        ' Find balancing storage (highest cumulative surplus And deficit)
        Dim HighestSurplus As Double, HighestDeficit As Double
        HighestSurplus = 0
        HighestDeficit = 0

        For i = 2 To Worksheets("Balancing storage calculations").UsedRange.Rows.Count
            If Worksheets("Balancing storage calculations").Cells(i, 3).Value > 0 Then
                HighestSurplus = IIf(Worksheets("Balancing storage calculations").Cells(i, 3).Value > HighestSurplus, Worksheets("Balancing storage calculations").Cells(i, 3).Value, HighestSurplus)
            Else
                HighestDeficit = IIf(Abs(Worksheets("Balancing storage calculations").Cells(i, 3).Value) > HighestDeficit, Abs(Worksheets("Balancing storage calculations").Cells(i, 3).Value), HighestDeficit)
            End If
        Next i

        BalancingStorage = HighestSurplus + HighestDeficit

        ' Calculate extra storage
        ExtraStorage = (EmergencyHours * Application.WorksheetFunction.Average(VolumeDemanded)) + (BottomHours * Application.WorksheetFunction.Average(VolumeDemanded))

        ' Calculate total storage And storage cost
        TotalStorage = BalancingStorage + ExtraStorage
        StorageCost = TotalStorage * StorageUnitCost

        ' Calculate total cost
        TotalCost = StorageCost + PipeCost

        ' Write data To "Optimization" worksheet
        Worksheets("Optimization").Cells(k + 1, 1).Value = pipeDiameterValue
        Worksheets("Optimization").Cells(k + 1, 2).Value = pipeUnitCostValue
        Worksheets("Optimization").Cells(k + 1, 3).Value = Q
        Worksheets("Optimization").Cells(k + 1, 4).Value = hf
        Worksheets("Optimization").Cells(k + 1, 5).Value = PipeCost
        Worksheets("Optimization").Cells(k + 1, 6).Value = BalancingStorage
        Worksheets("Optimization").Cells(k + 1, 7).Value = ExtraStorage
        Worksheets("Optimization").Cells(k + 1, 8).Value = TotalStorage
        Worksheets("Optimization").Cells(k + 1, 9).Value = StorageCost
        Worksheets("Optimization").Cells(k + 1, 10).Value = TotalCost
        k = k+1
    Next D

    ' Highlight the row With the optimum pipe And storage combination
    Dim OptimumRow As Long
    OptimumRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Worksheets("Optimization").Range("J2:J7")), Worksheets("Optimization").Range("J2:J7"), 0)
    Worksheets("Optimization").Range("A" & OptimumRow + 1 & ":J" & OptimumRow + 1).Interior.ColorIndex = 4

End Sub
