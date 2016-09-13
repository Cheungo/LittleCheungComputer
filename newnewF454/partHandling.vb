Module partHandling

    'Determine which machine code is data or not
    Public IsDAT(35) As Boolean

    'Determines number of parts in current line and what they are
    Public Component As ReturnComponent

    Structure ReturnComponent
        Dim ComponentParts() As String
        Dim ComponentLength As Integer
    End Structure

    'Splits current line into individual components to be processed
    Sub SplitIntoComponents(ByVal CurrentLine As String)
        'Declare Component Array for Parts
        ReDim Component.ComponentParts(2)
        Component.ComponentLength = 0

        'Changes line to upper case for validation
        Dim UpperCaseCurrentLine As String = ChangeInsToUpperCase(CurrentLine)
        'Split function splits line based on the space character
        Dim SplitLine() As String = Split(UpperCaseCurrentLine)
        Dim LastNonEmpty As Integer = -1

        'Loops through line and stores components in array in structure
        For i As Integer = 0 To SplitLine.Length - 1
            'Stores number of parts in line
            If SplitLine(i) <> "" Then
                LastNonEmpty += 1
                'Stores actual component in array(i)
                Component.ComponentParts(LastNonEmpty) = SplitLine(i)
                Component.ComponentLength += 1
            End If
        Next
    End Sub

    'Handles One Part Instructions
    Function OnePart(ByVal Component As String, ByVal Linenumber As Integer) As String
        DecodeResults = Decode(Component)

        If DecodeResults.Foundit = True Then
            'Set Label values to Null
            TempUnresolvedLabel.LabelText = ""
            TempUnresolvedLabel.Instruction = ""
            TempUnresolvedLabel.LabelAddress = Linenumber
            UnresolvedLabelList.Add(TempUnresolvedLabel)

            TempLabelRef.LabelText = ""
            TempLabelRef.LabelAddress = Linenumber
            LabelRefList.Add(TempLabelRef)

            'Validation for One Part Instructions
            Select Case DecodeResults.Output
                Case "901"
                    Return DecodeResults.Output

                Case "902"
                    Return DecodeResults.Output

                Case "000"
                    Return DecodeResults.Output

                'Other Instruction not accepted for one part
                Case Else
                    MsgBox("Error: " & Component & " cannot be used indivdually." & vbNewLine & "Please add labels or addresses to the instruction." & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                    Return "000"
            End Select
        Else
            MsgBox("Error: Instruction Not Found for " & Component & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
            Return "000"
        End If

    End Function

    'Handles Two Part Instructions
    Function TwoParts(ByVal Component0 As String, ByVal Component1 As String, ByVal Linenumber As Integer) As String

        Dim IsNumber As Boolean
        Dim IsInteger As Integer


        'Pass to Decode Function to Check for matches
        DecodeResults = Decode(Component0)

        'Sort and store in known and unknown Label arrays
        '================================================
        Select Case DecodeResults.Foundit '1st Part Instruction or Label?
            Case True 'Instruction is 1st, Label/Address/DAT is 2nd

                If Component0 = "INP" Or Component0 = "OUT" Or Component0 = "HLT" Then
                    MsgBox("Invalid Label/Address placement for " & Component0 & vbNewLine & "INP, OUT, HLT cannot be used in this format." & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                    Return "000"
                Else
                    IsNumber = IsNumeric(Component1) 'Checks 2nd Part 
                    Select Case IsNumber '2nd Part Address/DAT or Label?
                        Case True 'Address/DAT is 2nd

                            TempUnresolvedLabel.LabelText = ""
                            TempUnresolvedLabel.Instruction = ""
                            TempUnresolvedLabel.LabelAddress = Linenumber
                            UnresolvedLabelList.Add(TempUnresolvedLabel)

                            TempLabelRef.LabelText = ""
                            TempLabelRef.LabelAddress = Linenumber
                            LabelRefList.Add(TempLabelRef)


                            'Check if integer instead of any other data types
                            If Integer.TryParse(Component1, IsInteger) Then
                                'Storing DAT values
                                If Component0 = "DAT" Then
                                    Dim FormatDAT As Integer = Component1
                                    Return DATvalidation(FormatDAT, Linenumber)

                                Else
                                    'Not DAT therefore 2nd part is address
                                    If Component1 >= 0 And Component1 <= 35 Then
                                        Dim FormatAddress As Integer = Component1
                                        Return DecodeResults.Output & FormatAddress.ToString("D2")
                                    Else
                                        MsgBox("Error: Mailbox " & Component1 & " Does Not Exist (0 To 35 available registers)" & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                                        Return "000"
                                    End If
                                End If

                                'Value is not integer data type (e.g. decimal numbers)
                            Else
                                MsgBox("Error: Data value is not an integer between 0 to 999." & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                                IsDAT(Linenumber) = False
                                Return "000"
                            End If

                        Case False 'Label is 2nd
                            'Set Label values 
                            TempUnresolvedLabel.LabelText = Component1
                            TempUnresolvedLabel.Instruction = Component0
                            TempUnresolvedLabel.LabelAddress = Linenumber
                            UnresolvedLabelList.Add(TempUnresolvedLabel)

                            TempLabelRef.LabelText = ""
                            TempLabelRef.LabelAddress = Linenumber
                            LabelRefList.Add(TempLabelRef)

                            Return DecodeResults.Output
                    End Select
                End If

            Case False 'Label/Address is 1st (Address error), instruction is 2nd
                DecodeResults = Decode(Component1) 'Checks 2nd to see if actually instruction
                IsNumber = IsNumeric(Component0) 'Checks 1st to see if Label or Address

                Select Case DecodeResults.Foundit
                    Case True 'Instruction is 2nd

                        Select Case IsNumber '1st Part Address or Label/DAT?
                            Case True 'Address is 1st
                                MsgBox("Error: Invalid Address Placement for " & Component1 & vbNewLine & "Address cannot be placed infront of instruction." & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                                Return "000"
                            Case False 'Label is 1st
                                TempUnresolvedLabel.LabelText = ""
                                TempUnresolvedLabel.Instruction = ""
                                TempUnresolvedLabel.LabelAddress = Linenumber
                                UnresolvedLabelList.Add(TempUnresolvedLabel)

                                TempLabelRef.LabelText = Component0
                                TempLabelRef.LabelAddress = Linenumber
                                LabelRefList.Add(TempLabelRef)

                                'Storing DAT values
                                If Component1 = "DAT" Then
                                    IsDAT(Linenumber) = True
                                    Return "000"
                                Else
                                    'Not DAT so return machine code
                                    Return DecodeResults.Output.PadRight(3, "0"c)
                                End If

                        End Select

                    Case False
                        MsgBox("Error: Instruction Not Found for " & Component1 & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                        Return "000"
                End Select

        End Select
    End Function

    'Handles Three Part Instructions
    Function ThreeParts(ByVal Component0 As String, ByVal Component1 As String, ByVal Component2 As String, ByVal Linenumber As Integer) As String

        'Check if number
        Dim IsNumber As Boolean = IsNumeric(Component2)
        Dim IsInteger As Integer

        'Pass to Decode Function to Check for matches
        DecodeResults = Decode(Component1)

        'Sort and store in known and unknown Label arrays
        '================================================
        Select Case DecodeResults.Foundit
            Case True
                Select Case IsNumber '3rd Part Address/DAT or Label?
                    Case True 'Address/DAT is 3rd Part
                        '1st Part is Label
                        TempUnresolvedLabel.LabelText = ""
                        TempUnresolvedLabel.Instruction = ""
                        TempUnresolvedLabel.LabelAddress = Linenumber
                        UnresolvedLabelList.Add(TempUnresolvedLabel)

                        TempLabelRef.LabelText = Component0
                        TempLabelRef.LabelAddress = Linenumber
                        LabelRefList.Add(TempLabelRef)

                        'Check if integer instead of any other data types
                        If Integer.TryParse(Component2, IsInteger) Then
                            'Storing DAT values
                            If Component1 = "DAT" Then
                                Dim FormatDAT As Integer = Component2
                                Return DATvalidation(FormatDAT, Linenumber)

                            Else
                                'Not DAT therefore 2nd part is address
                                If Component2 >= 0 And Component2 <= 35 Then
                                    Dim FormatAddress As Integer = Component2
                                    Return DecodeResults.Output & FormatAddress.ToString("D2")
                                Else
                                    MsgBox("Error: Mailbox " & Component1 & " Does Not Exist (0 To 35 available registers)" & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                                    Return "000"
                                End If
                            End If

                            'Value is not integer data type (e.g. decimal numbers)
                        Else
                            MsgBox("Error: Data value is not an integer between 0 to 999." & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                            IsDAT(Linenumber) = False
                            Return "000"
                        End If

                    Case False 'Label is 3rd
                        '3rd Part is Unresolved
                        TempUnresolvedLabel.LabelText = Component2
                        TempUnresolvedLabel.Instruction = Component1
                        TempUnresolvedLabel.LabelAddress = Linenumber
                        UnresolvedLabelList.add(TempUnresolvedLabel)

                        '1st Part is Label
                        TempLabelRef.LabelText = Component0
                        TempLabelRef.LabelAddress = Linenumber
                        LabelRefList.Add(TempLabelRef)

                        'Return Instruction
                        Return DecodeResults.Output
                End Select

            Case False
                MsgBox("Error: Instruction Not Found for " & Component1 & vbNewLine & "On Line " & Linenumber, MsgBoxStyle.Exclamation)
                Return "000"

        End Select
    End Function

    'Handles DAT Validation
    Function DATvalidation(ByVal FormatDAT As Integer, ByVal Linenumber As Integer) As String
        IsDAT(Linenumber) = True
        Select Case FormatDAT
            Case > 999
                MsgBox("Data entered at line " & Linenumber & " is greater than 999. Please enter an integer within the range of 0 and 999", MsgBoxStyle.Exclamation)
                IsDAT(Linenumber) = False
                Return "000"
            Case < 0
                MsgBox("Data entered at line " & Linenumber & " is a negative number. Please enter an integer within the range of 0 and 999", MsgBoxStyle.Exclamation)
                IsDAT(Linenumber) = False
                Return "000"
            Case Else
                Return FormatDAT.ToString("D3")

        End Select

    End Function

    'Converts Instructions to Uppercase
    Function ChangeInsToUpperCase(ByVal Ins As String) As String
        If TypeOf Ins Is String Then
            Return UCase(Ins)
        End If
    End Function

End Module
