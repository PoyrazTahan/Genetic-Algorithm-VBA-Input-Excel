Global p  As Integer 'this variable holds the value of number of open store you can have
Global size  As Integer ' this value holds the value of number of cities are there


Sub geneticAlgo()





size = Worksheets("Prepare Sheet").Range("c1") ' determining how many cities are there?

Dim individuals() As Boolean
ReDim individuals(1 To size)


Dim generationStart() As Integer   ' this variable will hold the generation at hand at any time
    
generationSize = Worksheets("Model").Range("b2") ' this variable will hold the population size taken from Worksheet("Model")
ReDim generationStart(1 To generationSize, 1 To size)


p = Worksheets("Model").Range("b3") ' taken from worksheet("Model")
Dim sum As Integer

Dim fitness() As Integer ' this array will hold the fitness level calculated from the model // notice that indexes represent indiviuals
ReDim fitness(1 To generationSize)

Dim PastPlacemnet As Integer ' this variable will hold how many individuals have been written into worksheet("NextGen")
PastPlacemnet = 0


Dim SameError As Boolean
Dim SameCounter As Integer
SameError = True

Dim pointPlacer As Integer 'We will sort while we are writing the generation into NextGen
Dim Test As Integer

Dim indexNumber As Integer



For i = 1 To generationSize
    Randomize
    SameError = True
    sum = 0
    ' Creating a creature
    For j = 1 To size
        If isOpen() Then
            generationStart(i, j) = 1
            sum = sum + 1
        Else
            generationStart(i, j) = 0
        End If
    Next j
    ' Checking if creature is valid
    If sum <> 5 Then
        i = i - 1
    Else
    
        'First Put it into Model created creature
        For j = 1 To size
            Worksheets("Model").Range("E5").Offset(j) = generationStart(i, j)
        Next j
        'Determine its Fitness
        fitness(i) = Worksheets("Model").Range("Y21")
        
        'Records To NextGen
        If PastPlacemnet = 0 Then ' If this is the first recordmend
            ' This Part only works in the first step of the loop
            Worksheets("NextGen").Range("A1").Offset(0, i - 1) = i
            Worksheets("NextGen").Range("A2").Offset(0, i - 1) = fitness(i)
            For j = 1 To size
                Worksheets("NextGen").Range("A3").Offset(j - 1) = generationStart(i, j)
            Next j
            PastPlacemnet = PastPlacemnet + 1
        ElseIf PastPlacemnet <> 0 Then ' If this is not the first recordment check if same individual exist
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'This parts check if individuals are the same
            For j = 1 To PastPlacemnet ' look at each recorded fitness to give you hint of if it is the same
                If fitness(i) = Worksheets("NextGen").Range("A2").Offset(0, j - 1) Then ' if the fitness levels are same check ind.
                    
                    SameCounter = 0
                    For k = 1 To size ' if You Get An Alert Compare individulas
                        indexNumber = Worksheets("NextGen").Range("A1").Offset(0, j - 1)
                        If generationStart(i, k) = generationStart(indexNumber, k) Then ' if any city of a creature is diferent the alarm will go of
                            SameCounter = SameCounter + 1
                        End If
                    Next k
                    If SameCounter = size Then 'if the last matching returns all same it will go kill the item
                        SameError = True
                        Exit For
                    End If
                ElseIf fitness(i) <> Worksheets("NextGen").Range("A2").Offset(0, j - 1) Then ' if the fitness level is not the same they are the same
                   SameError = False
                End If
            Next j
            
            If SameError = True Then
                i = i - 1 'Kils creature if same
            ElseIf SameError = False Then
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' This Parts insert the individual in a sorted way
                            
                For k = 1 To PastPlacemnet
                    If fitness(i) < Worksheets("NextGen").Range("A2").Offset(0, k - 1) Then
                        pointPlacer = k
                        Exit For
                    End If
            
                Next k
                If k <> PastPlacemnet + 1 Then
                    Columns(pointPlacer).Insert
                    Worksheets("NextGen").Range("A1").Offset(0, pointPlacer - 1) = i
                    Worksheets("NextGen").Range("A2").Offset(0, pointPlacer - 1) = fitness(i)
                    For k = 1 To size
                        Worksheets("NextGen").Range("A2").Offset(k, pointPlacer - 1) = generationStart(i, k)
                    Next k
                    PastPlacemnet = PastPlacemnet + 1
                Else
                    Worksheets("NextGen").Range("A1").Offset(0, PastPlacemnet) = i
                    Worksheets("NextGen").Range("A2").Offset(0, PastPlacemnet) = fitness(i)
                    For k = 1 To size
                        Worksheets("NextGen").Range("A2").Offset(k, PastPlacemnet) = generationStart(i, k)
                    Next k
                    PastPlacemnet = PastPlacemnet + 1
                End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
            End If
        End If
        
        
        
        
    End If
    
    
Next i



' until her was the initilization part
' Now starts the iterative part of genetic algo


Dim iterationLimit As Integer
iterationLimit = Worksheets("Model").Range("B1")
For Iteration = 1 To iterationLimit
                'finding the best and the worst of a generation
                Dim bestfit As Integer
                bestfit = 1
                worstfit = 1
                
                For i = 2 To generationSize
                    If fitness(bestfit) > fitness(i) Then
                        bestfit = i
                    End If
                    If fitness(worstfit) < fitness(i) Then
                        worstfit = i
                    End If
                Next i
                
                Dim generationSpan As Integer
                generationSpan = fitness(worstfit) - fitness(bestfit)
                
                'start kiling
                Dim deathCount As Integer
                deathCount = 0
                
                
                Dim deadIndividuals() As Boolean
                ReDim deadIndividuals(1 To generationSize)
                
                For i = 1 To generationSize
                
                    If fitness(bestfit) + generationSpan * 0.2 > fitness(i) Then
                        ' best 20% fears no harm
                    ElseIf fitness(bestfit) + generationSpan * 0.7 > fitness(i) Then
                        ' 40% of next best 50% dies
                        If Rnd < 0.6 Then
                            For j = 1 To size
                                generationStart(i, j) = 0
                                
                            Next j
                            deadIndividuals(i) = True
                            deathCount = deathCount + 1
                        End If
                    Else
                        ' last 30% dies wtih certainty because they are not adaptive enough
                        For j = 1 To size
                            generationStart(i, j) = 0
                        Next j
                        deadIndividuals(i) = True
                        deathCount = deathCount + 1
                    End If
                Next i
                
                
                
                Dim childIndex As Integer             'This index will look at Next Gen bestfit to worstfit if that indexed individual died or not
                Dim mother, motherIndex As Integer     'These are the parents
                Dim father, fatherIndex As Integer
                Dim MutationGen As Integer
                Dim MutaionRandom As Double
                Dim IndividualZeroArray() As Integer      ' For mutation to know which can be opened
                Dim IndividualOneArray() As Integer       ' For mutation to know which can be closed
                
                ReDim IndividualZeroArray(1 To size - p)
                ReDim IndividualOneArray(1 To p)
                
                'Start breeding new gen on top of the dead ones of the last gen
                For i = 1 To generationSize ' i here represents the column which we are looking at worksheet("NextGen")
                    Randomize
                    'Determine the mother & father of the new breed
                    childIndex = Worksheets("NextGen").Range("A1").Offset(0, i - 1)
                    If deadIndividuals(childIndex) Then
                        mother = Int((i - 1 + 1) * Rnd + 1) ' determine the palce in the NextGen sheet of mother
                        Do
                            father = Int((i - 1 + 1) * Rnd + 1) ' determine the palce in the NextGen sheet of father //notice it can not be same creature with mother
                        Loop Until mother <> father
                        motherIndex = Worksheets("NextGen").Range("A1").Offset(0, mother - 1) 'Finds in which index does the mother contained
                        fatherIndex = Worksheets("NextGen").Range("A1").Offset(0, father - 1) 'Finds in which index does the father contained
                        'mother father determined
                    
                        'start creating the new breed
                        sum = 0
                        For j = 1 To size
                            
                            If generationStart(motherIndex, j) = generationStart(fatherIndex, j) Then
                                generationStart(childIndex, j) = generationStart(fatherIndex, j)
                            Else
                                'if mother and father don't have the same characteristics in the same chose with 50% to get one
                                If Rnd < 0.5 Then
                                    generationStart(childIndex, j) = generationStart(fatherIndex, j)
                                Else
                                    generationStart(childIndex, j) = generationStart(motherIndex, j)
                                End If
                                
                            End If
                            
                            If generationStart(childIndex, j) = 1 Then
                                sum = sum + 1
                            End If
                        Next j
                        
                        If sum <> 5 Then
                            i = i - 1
                        Else
                            Dim mutationzero1, mutationzero2, mutationzero3 As Integer
                            Dim mutationone1, mutationone2, mutationone3 As Integer
                            'mutaion dynamic
                            For j = 1 To size - p
                                IndividualZeroArray(j) = 0
                                
                            Next j
                            For j = 1 To p
                                IndividualOneArray(j) = 0
                                
                            Next j
                            
                            zerocounter = 1
                            onercounter = 1
                            For j = 1 To size
                                If generationStart(childIndex, j) = 0 Then
                                    IndividualZeroArray(zerocounter) = j
                                    zerocounter = zerocounter + 1
                                Else
                                    IndividualOneArray(onercounter) = j
                                    onercounter = onercounter + 1
                                End If
                            Next j
                            MutationGen = 0
                            MutaionRandom = Rnd
                            If MutaionRandom < 0.001 Then
                                MutationGen = 3
                                
                                ''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Determining which are going to mutate which means change behaviour
                                mutationzero1 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                Do
                                    mutationzero2 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                Loop Until mutationzero2 <> mutationzero1
                                Do
                                    mutationzero3 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                Loop Until mutationzero2 <> mutationzero3 And mutationzero3 <> mutationzero1
                                '''''''''''''
                                mutationone1 = Int((p - 1 + 1) * Rnd + 1)
                                Do
                                    mutationone2 = Int((p - 1 + 1) * Rnd + 1)
                                Loop Until mutationone1 <> mutationone2
                                Do
                                    mutationone3 = Int((p - 1 + 1) * Rnd + 1)
                                Loop Until mutationone3 <> mutationone1 And mutationone3 <> mutationone2
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' make the changes according to determined values
                                
                                generationStart(childIndex, IndividualZeroArray(mutationzero1)) = 1
                                generationStart(childIndex, IndividualZeroArray(mutationzero2)) = 1
                                generationStart(childIndex, IndividualZeroArray(mutationzero3)) = 1
                                
                                generationStart(childIndex, IndividualOneArray(mutationone1)) = 0
                                generationStart(childIndex, IndividualOneArray(mutationone2)) = 0
                                generationStart(childIndex, IndividualOneArray(mutationone3)) = 0
                                
                                
                            ElseIf MutaionRandom < 0.01 Then
                                MutationGen = 2
                                                               ''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Determining which are going to mutate which means change behaviour
                                mutationzero1 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                Do
                                    mutationzero2 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                Loop Until mutationzero2 <> mutationzero1

                                '''''''''''''
                                mutationone1 = Int((p - 1 + 1) * Rnd + 1)
                                Do
                                    mutationone2 = Int((p - 1 + 1) * Rnd + 1)
                                Loop Until mutationone1 <> mutationone2
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' make the changes according to determined values
                                
                                generationStart(childIndex, IndividualZeroArray(mutationzero1)) = 1
                                generationStart(childIndex, IndividualZeroArray(mutationzero2)) = 1
                                
                                generationStart(childIndex, IndividualOneArray(mutationone1)) = 0
                                generationStart(childIndex, IndividualOneArray(mutationone2)) = 0
    
                                
                                
                            ElseIf MutaionRandom < 0.1 Then
                                MutationGen = 1
                               ''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Determining which are going to mutate which means change behaviour
                                mutationzero1 = Int(((size - p) - 1 + 1) * Rnd + 1)
                                '''''''''''''
                                mutationone1 = Int((p - 1 + 1) * Rnd + 1)
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' make the changes according to determined values
                                
                                generationStart(childIndex, IndividualZeroArray(mutationzero1)) = 1
                                
                                generationStart(childIndex, IndividualOneArray(mutationone1)) = 0
                            
                            End If
                            
                            
                            
                            'Now you are sure at this point there is a new breed at on top of the dead one
                            'New breed will enter the model to generate its fitness level
                            
                            For j = 1 To size
                                Worksheets("Model").Range("E5").Offset(j) = generationStart(childIndex, j)
                            Next j
                            fitness(childIndex) = Worksheets("Model").Range("Y21")
    
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'new breed could be same with the existing creatures so this needs to be checked
                            SameError = True
                            sum = 0
                                 'This parts check if individuals are the same
                                For k = 1 To PastPlacemnet ' look at each recorded fitness to give you hint of if it is the same
                                    ' pastPlacement here should be 100
                                    If fitness(childIndex) = Worksheets("NextGen").Range("A2").Offset(0, k - 1) Then ' if the fitness levels are same check ind.

                                        SameCounter = 0
                                        For l = 1 To size ' if You Get An Alert Compare individulas
                                            indexNumber = Worksheets("NextGen").Range("A1").Offset(0, k - 1)
                                            If generationStart(childIndex, l) = generationStart(indexNumber, l) Then ' if any city of a creature is diferent the alarm will go of
                                                SameCounter = SameCounter + 1
                                            End If
                                        Next l
                                        If SameCounter = size Then 'if the last matching returns all same it will go kill the item
                                            SameError = True
                                            Exit For
                                        End If
                                    ElseIf fitness(childIndex) <> Worksheets("NextGen").Range("A2").Offset(0, k - 1) Then ' if the fitness level is not the same they are the same
                                       SameError = False
                                    End If
                                Next k

                                If SameError = True Then
                                    i = i - 1 'Kils creature if same
                                End If
                                
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' Write the creature down to the NextGen
                            Worksheets("NextGen").Range("A1").Offset(0, i - 1) = childIndex
                            Worksheets("NextGen").Range("A2").Offset(0, i - 1) = fitness(childIndex)
                            For j = 1 To size
                                Worksheets("NextGen").Range("A2").Offset(j, i - 1) = generationStart(childIndex, j)
                            Next j
                            
                        End If
 
                    End If
    
                Next i
                
                'all creatures now back alive with the new breed
                For i = 1 To generationSize
                    deadIndividuals(i) = False
                Next i
                
                
                
                Worksheets("NextGen").Range(Range("A1"), Range("A1").End(xlDown).End(xlToRight)).Clear
                'Building the next Gen
                
                flag = False
                sum = 0
                
                
                PastPlacemnet = 0
                
                For i = 1 To generationSize
                    sum = 0
                    For j = 1 To size
                        sum = sum + generationStart(i, j)
                    Next j
                    pointPlacer = 1
                    If sum = 5 Then
                        Worksheets("NextGen").Range("A1").Offset(0, pointPlacer - 1) = i
                        Worksheets("NextGen").Range("A2").Offset(0, pointPlacer - 1) = fitness(i)
                        For j = 1 To size
                             Worksheets("NextGen").Range("A2").Offset(j, pointPlacer - 1) = generationStart(i, j)
                        Next j
                        PastPlacemnet = PastPlacemnet + 1
                        Exit For
                    End If
                Next i
                
                
                
                For i = i + 1 To generationSize
                    sum = 0
                    For j = 1 To size
                        sum = sum + generationStart(i, j)
                 
                    Next j
                
                    If sum = 5 Then
                        pointPlacer = 0
                        For j = 1 To PastPlacemnet
                            If fitness(i) < Worksheets("NextGen").Range("A2").Offset(0, j - 1) Then
                                pointPlacer = j
                                Exit For
                            End If
                            
                        Next j
                        If j <> PastPlacemnet + 1 Then
                            Columns(pointPlacer).Insert
                            Worksheets("NextGen").Range("A1").Offset(0, pointPlacer - 1) = i
                            Worksheets("NextGen").Range("A2").Offset(0, pointPlacer - 1) = fitness(i)
                            For j = 1 To size
                                Worksheets("NextGen").Range("A2").Offset(j, pointPlacer - 1) = generationStart(i, j)
                            Next j
                            PastPlacemnet = PastPlacemnet + 1
                        Else
                            Worksheets("NextGen").Range("A1").Offset(0, PastPlacemnet) = i
                            Worksheets("NextGen").Range("A2").Offset(0, PastPlacemnet) = fitness(i)
                            For j = 1 To size
                                Worksheets("NextGen").Range("A2").Offset(j, PastPlacemnet) = generationStart(i, j)
                            Next j
                            PastPlacemnet = PastPlacemnet + 1
                        End If
                    End If
                Next i

Next Iteration

End Sub

Public Function isOpen() As Boolean
    Randomize
    If Rnd < (p / size) Then
        isOpen = True
        Exit Function
    End If
    
    isOpen = False
    
End Function

