Module decodeIns

    'Holds machine code output & boolean
    Public DecodeResults As ReturnDecode

    Structure ReturnDecode
        Dim Output As String
        Dim Foundit As Boolean
    End Structure

    'Handles Mnemonic(Instruction) and Machine Code Matchup & Returns Machine Code
    Function Decode(ByVal CurrentLine As String) As Object
        Dim Mnemonic(10) As String
        Dim MachineCode(10) As String

        Mnemonic(0) = "LDA"
        Mnemonic(1) = "STA"
        Mnemonic(2) = "ADD"
        Mnemonic(3) = "SUB"
        Mnemonic(4) = "INP"
        Mnemonic(5) = "OUT"
        Mnemonic(6) = "HLT"
        Mnemonic(7) = "BRA"
        Mnemonic(8) = "BRZ"
        Mnemonic(9) = "BRP"
        Mnemonic(10) = "DAT"

        MachineCode(0) = "5"
        MachineCode(1) = "3"
        MachineCode(2) = "1"
        MachineCode(3) = "2"
        MachineCode(4) = "901"
        MachineCode(5) = "902"
        MachineCode(6) = "000"
        MachineCode(7) = "6"
        MachineCode(8) = "7"
        MachineCode(9) = "8"
        MachineCode(10) = ""

        Dim FoundMachineCode As Integer

        'Matches Mnemonic to Machine Code
        For i = 0 To Mnemonic.Length - 1
            If CurrentLine = Mnemonic(i) Then
                FoundMachineCode = i
                DecodeResults.Foundit = True
                Exit For

                'If Instruction not found
            ElseIf i = 10 And CurrentLine <> Mnemonic(Mnemonic.Length - 1) Then
                DecodeResults.Foundit = False
            End If
        Next

        DecodeResults.Output = MachineCode(FoundMachineCode)
        Return DecodeResults
    End Function
End Module
