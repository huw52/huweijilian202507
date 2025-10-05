'***********************************************************************************
'James Martin, AEA Technology, Engineering Software, Hyprotech UK Technical Services
'Email: support@hyprotech.com
'***********************************************************************************

'Note:  Hysys Type library must be referenced under Tools ... References

'Require explicit variable declaration
Option Explicit

Const EMPTYTEXT As String = "<empty>"
Const NATEXT As String = "---"

Public Sub GetFeedAndResultInfo()
'
'Description:    Demonstrates how to link to a PFR in Hysys
'                and obtain results info via back door variable methods
'
'    NB Backdoor methods are only recommended when there is no other alternative, the internal
'    Hysys "monikers" may not remain constant between versions
'    so care should be exercised when upgrading
'
    'Declare Variables----------------------------------------------------------------------------
    Dim hyApp As HYSYS.Application          'Hysys Application
    Dim hyCase As HYSYS.SimulationCase      'Hysys Case
    Dim hyStream As HYSYS.ProcessStream     'Process stream variable
    Dim CompTransferVar As Variant          'For retrieving composition
    Dim Counter As Integer
    Dim Counter2 As Integer
    Dim AppBuildNo As String                'Declare AppBuildNo variable
    
    Dim hyPFR As HYSYS.PFReactor    'PRF object
    Dim hyPFRbd As BackDoor         'Back door Object - Must be declared as BackDoor (not As Object) to force the QueryInterface call

    Dim PipeLengths As Variant  'Array holding the "Length" results
    
    Dim HeaderData As Variant   'Heading data for rxn rates
    Dim ResultData As Variant   'results for rxn rates
    
    'Procedure--------------------------------------------------------------------------------------

    'Setup Error Handler
    On Error GoTo ErrorHandler
    
    'Link to Hysys---------------------------------------------------
    Set hyApp = GetObject(, "HYSYS.Application")    'Only works if Hysys is open
    'Get the currently open case
    Set hyCase = hyApp.ActiveDocument
    If hyCase Is Nothing Then
        MsgBox "Make sure Hysys is open with a case containing a PFR of the specified name", , "Error"
        Exit Sub
    End If

    'Link to the PFR object
    Set hyPFR = hyCase.Flowsheet.Operations.Item(Range("PFRName").Value)
    
    'Get a backdoor variable for it
    Set hyPFRbd = hyPFR
    
    'Link to the first Feed stream
    Set hyStream = hyPFR.Feeds.Item(0)

    'Get Feed Properties
    With hyStream
        If .Temperature.IsKnown = True Then
            Range("FeedTemp").Value = .Temperature.GetValue("C")
        Else
            Range("FeedTemp").Value = EMPTYTEXT
        End If
        
        If .Pressure.IsKnown = True Then
            Range("FeedPres").Value = .Pressure.GetValue("bar")
        Else
            Range("FeedPres").Value = EMPTYTEXT
        End If
    
        If .MolarFlow.IsKnown = True Then
            Range("FeedFlow").Value = .MolarFlow.GetValue("kgmole/h")
        Else
            Range("FeedFlow").Value = EMPTYTEXT
        End If
        
        'Clear any existing component info
        With Range("FeedCompStart")
            Counter = 0
            Do While .Offset(Counter, -1).Value <> ""
                For Counter2 = -1 To 0
                    .Offset(Counter, Counter2).Value = ""
                Next 'Counter2
                Counter = Counter + 1
            Loop
        End With
            
        'List the components in the feed stream and their mole fractions
        CompTransferVar = .ComponentMolarFractionValue
        If CompTransferVar(0) <> -32767 Then
            For Counter = 0 To UBound(CompTransferVar)
                Range("FeedCompStart").Offset(Counter, -1).Value = hyCase.Flowsheet.FluidPackage.Components.Item(Counter)
                Range("FeedCompStart").Offset(Counter, 0).Value = CompTransferVar(Counter)
            Next 'Counter
        Else
            For Counter = 0 To UBound(CompTransferVar)
                Range("FeedCompStart").Offset(Counter, -1).Value = hyCase.Flowsheet.FluidPackage.Components.Item(Counter)
                Range("FeedCompStart").Offset(Counter, 0).Value = EMPTYTEXT
            Next 'Counter
        End If
    End With
    
    'Clear any existing pipe segment info
    With Range("OutputStart")
        Counter = 0
        Do While .Offset(Counter, 0).Value <> ""
            For Counter2 = 0 To 85
                .Offset(Counter, Counter2).Value = ""
            Next 'Counter2
            Counter = Counter + 1
        Loop
    End With
    
    'Now get the PFR result info
    
    'Initialise counter
    Counter = 0
    
    'Get the build number of Hysys
    AppBuildNo = hyApp.ProgramBuild
    
    'Not all the properties of the PRF are "wrapped" for access via normal OLE as yet.
    '
    'However properties of unwrapped objects can be obtained via a "backdoor" variable
    'if you know the "Moniker" - (Hysys internal identifier) for the property
    '
    'Monikers can most easily be identified by recording a script while you import the property
    'of interest to a spreadsheet
    'examination of the script file generated using notepad will show the moniker names
    '
    'Monikers can also be extracted from the rdf (report definition file) for the operation
    '
    'For the PFR length data is stored at the following Moniker
    '   :Length.550.#
    'Where # is an integer starting from 0 and refers to the position
    'in the array of results data
    '
    'Other relevant Monikers:
    'Temperature        :Temperature.550.#
    'Pressure           :Pressure.550.#
    'etc...
    
    'First get an array carrying the lengths at which results are reported
    'Since a [] is used instead of a specific number then retrieve the whole array
    PipeLengths = hyPFRbd.BackDoorVariable(":Length.550.[]").Variable.GetValues("m")
    
    'From the size of the array we hence know how entries there are in the results arrays
    For Counter = 0 To UBound(PipeLengths)
        
        With Range("OutputStart")

            'Now get some of the other properties
            
            'Conditions -------------------------------

            'Length
            .Offset(Counter, 0).Value = PipeLengths(Counter)

            'Temperature - call a procedure to deal with possible empty values
            .Offset(Counter, 1).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":Temperature.550." & CStr(Counter)).Variable, "C")

            'Pressure
            .Offset(Counter, 2).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":Pressure.550." & CStr(Counter)).Variable, "bar")
            
            'Vapour Fraction
            .Offset(Counter, 3).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":VapourFraction.550." & CStr(Counter)).Variable, "")
            
            'Duty
            .Offset(Counter, 4).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":HeatFlow.551." & CStr(Counter)).Variable, "kcal/h")
            
            'Enthalpy
            .Offset(Counter, 5).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":Enthalpy.550." & CStr(Counter)).Variable, "kcal/kgmole")
            
            'Enthalpy
            .Offset(Counter, 6).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":Entropy.550." & CStr(Counter)).Variable, "kJ/kgmole-C")
            
            'Inside HTC
            .Offset(Counter, 7).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":HeatTranCoeff.550." & CStr(Counter)).Variable, "kJ/h-m2-C")
            
            'Outside HTC
            .Offset(Counter, 8).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":HeatTranCoeff.551." & CStr(Counter)).Variable, "kJ/h-m2-C")

            'Flows -------------------------------

            'Length
            .Offset(Counter, 10).Value = PipeLengths(Counter)
            
            'Molar Flow
            .Offset(Counter, 11).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":MolarFlow.550." & CStr(Counter)).Variable, "kgmole/h")
            
            'Mass Flow
            .Offset(Counter, 12).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":MassFlow.550." & CStr(Counter)).Variable, "kg/h")
            
            'Volumetric Flow
            .Offset(Counter, 13).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":ActVolFlow.550." & CStr(Counter)).Variable, "m3/h")
            
            'Heat Flow
            .Offset(Counter, 14).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":HeatFlow.550." & CStr(Counter)).Variable, "kcal/h")

        End With

    Next 'Counter
    
    'Overall Reaction Rates----------------------------------
    
    'Now get the array of reaction names
    HeaderData = hyPFRbd.BackDoorVariable(":RxnRateName.550.[1-]").Variable.Values
    
    'and the matrix of reaction results
    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.505.[1-].[]").Variable.GetValues("kgmole/m3-s")
    
    'Output these names
    For Counter = 1 To UBound(HeaderData)
        
        With Range("OutputStart")
            
            'Write the reaction name
            .Offset(-1, 16 + Counter).Value = HeaderData(Counter)
            
            'Write the ...
            For Counter2 = 0 To UBound(PipeLengths)
                
                '... Lengths
                .Offset(Counter2, 16).Value = PipeLengths(Counter2)
                '... Rxn Rates
                .Offset(Counter2, 16 + Counter).Value = ResultData(Counter2, Counter)
            
            Next 'Counter2

        End With

    Next 'Counter
    
    'NB There is the possbility that if there are lots of rxns this data could run into
    'that in the next cols along - obviously this could be fixed with a little more coding ...
    
    'Component Production Rates----------------------------------
    
    'Get an array of component rates for the table
    HeaderData = hyCase.Flowsheet.FluidPackage.Components.Names
    
    'Cpt reaction data
    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.506.[1-].[]").Variable.GetValues("kgmole/m3-s")

    'Output these names
    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            
            'Write the component name
            .Offset(-1, 22 + Counter).Value = HeaderData(Counter)
            
            'Write the ...
            For Counter2 = 0 To UBound(PipeLengths)
                
                '... Lengths
                .Offset(Counter2, 21).Value = PipeLengths(Counter2)
                '... Rxn Rates
                .Offset(Counter2, 22 + Counter).Value = ResultData(Counter2, Counter + 1)
            
            Next 'Counter2

        End With

    Next 'Counter
    
    'Transport Properties -------------------------------

    For Counter = 0 To UBound(PipeLengths)
        
        With Range("OutputStart")

            'Length
            .Offset(Counter, 29).Value = PipeLengths(Counter)

            .Offset(Counter, 30).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":Viscosity.550." & CStr(Counter)).Variable, "cP")
            .Offset(Counter, 31).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":MoleWeight.550." & CStr(Counter)).Variable, "")
            .Offset(Counter, 32).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":MassDensity.550." & CStr(Counter)).Variable, "kg/m3")
            .Offset(Counter, 33).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":HeatCapacity.550." & CStr(Counter)).Variable, "kJ/kgmole-C")
            .Offset(Counter, 34).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":SurfaceTen.550." & CStr(Counter)).Variable, "dyne/cm")
            .Offset(Counter, 35).Value = ReturnValOrEmpty(hyPFRbd.BackDoorVariable(":ZFactor.550." & CStr(Counter)).Variable, "")
                       
        End With
    
    Next 'Counter
    
    'Component Flow Rates----------------------------------

    'Molar Flows

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.511.[1-].[]").Variable.GetValues("kgmole/h")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 38 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 37).Value = PipeLengths(Counter2)
                .Offset(Counter2, 38 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter
    
    'Molar Fractions

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.507.[1-].[]").Variable.GetValues("")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 46 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 45).Value = PipeLengths(Counter2)
                .Offset(Counter2, 46 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter

    'Mass Flows

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.512.[1-].[]").Variable.GetValues("kg/h")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 54 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 53).Value = PipeLengths(Counter2)
                .Offset(Counter2, 54 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter

    'Mass Fractions

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.508.[1-].[]").Variable.GetValues("")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 62 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 61).Value = PipeLengths(Counter2)
                .Offset(Counter2, 62 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter
    
    'LiqVol Flows

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.513.[1-].[]").Variable.GetValues("m3/h")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 70 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 69).Value = PipeLengths(Counter2)
                .Offset(Counter2, 70 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter

    'LiqVol Fractions

    ResultData = hyPFRbd.BackDoorVariable(":ExtraData.509.[1-].[]").Variable.GetValues("")

    For Counter = 0 To UBound(HeaderData)
        
        With Range("OutputStart")
            .Offset(-1, 78 + Counter).Value = HeaderData(Counter)
            For Counter2 = 0 To UBound(PipeLengths)
                .Offset(Counter2, 77).Value = PipeLengths(Counter2)
                .Offset(Counter2, 78 + Counter).Value = ResultData(Counter2, Counter + 1)
            Next 'Counter2
        End With

    Next 'Counter

    'Get rid of all our object vars
    Set hyApp = Nothing
    Set hyCase = Nothing
    Set hyStream = Nothing
    Set hyPFR = Nothing
    Set hyPFRbd
