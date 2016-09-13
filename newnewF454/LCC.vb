Imports System.IO

Public Class LCC
    'Final Decoded instructions are stored in this array
    Dim BuildInstruction(35) As String

    'For Execute Instruction
    Dim ProgCount As Integer = 0

    'For more than one input
    Dim MoreINP As Boolean = False

    'Submit Button executes program
    Private Sub BtnSubmit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnSubmit.Click
        'If more than one input prompt / Continue progress from last time
        If MoreINP = True Then
            MoreINP = False
            ExecuteIns()
        Else
            'Start normally
            ResetProgram()
            main()
            ExecuteIns()
        End If
    End Sub

    'Resets all variables and text boxes to initial values
    Sub ResetProgram()
        'Clear All Variables
        Array.Clear(BuildInstruction, 0, BuildInstruction.Length)
        Array.Clear(IsDAT, False, IsDAT.Length)
        Component.ComponentLength = 0
        DecodeResults.Output = 0
        DecodeResults.Foundit = False
        ProgCount = 0
        TempLabelRef.LabelText = ""
        TempLabelRef.LabelAddress = 0
        TempUnresolvedLabel.LabelText = ""
        TempUnresolvedLabel.Instruction = ""
        TempUnresolvedLabel.LabelAddress = 0
        LabelRefList.Clear()
        UnresolvedLabelList.Clear()

        'Reset Line number box
        txtNum.Text = ""

        'Reset Accumulator
        txtAccum.Text = "000"

        'Reset Output Box
        txtOutput.Text = ""

        'Reset ProgCount
        txtProgCount.Text = ""

        'Reset Mail Boxes
        For i = 0 To 35
            BuildInstruction(i) = "000"
            Me.Controls("reg" & i).Text = BuildInstruction(i)
            Me.Controls("reg" & i).BackColor = Color.White
            Me.Controls("lab" & i).ForeColor = Color.Black
        Next
    End Sub

    'Adds line numbers next to code
    Sub AddLineNum(ByVal Linenumber As Integer)
        txtNum.Text = txtNum.Text & Linenumber & vbNewLine
    End Sub

    'Trim leading spaces in user program
    Function TrimSpaces(ByVal Line As String)
        Return LTrim(Line)
    End Function

    ''' <summary>
    ''' Links everything together
    ''' </summary>
    ''' 
    Sub main()
        Dim CurrentLine As String = ""
        Dim NumOfLines As Integer = txtUserCode.Lines.Length - 1

        'Reads the current line and passes to handling functions depending on Parts
        For linenumber = 0 To NumOfLines
            AddLineNum(linenumber) 'Add line numbers next to code
            CurrentLine = TrimSpaces(txtUserCode.Lines(linenumber)) 'Stores current line in variable
            If CurrentLine = "" Then
                'Do nothing
            Else
                SplitIntoComponents(CurrentLine) 'Current line is split into components to be processed
                'Too many lines (mail box limitations)
                If NumOfLines > 35 Then
                    MsgBox("Error: Insufficient Mailbox Capacity - Please reduce the number of lines in your program (" & NumOfLines & " Lines).", MsgBoxStyle.Exclamation)
                    Exit For
                Else
                    'Components of current line is passed to functions to be processed based on Num of Components
                    'Buildinstruction stores complete set of Machine Code at end
                    Select Case Component.ComponentLength
                        Case 1
                            BuildInstruction(linenumber) = OnePart(Component.ComponentParts(0), linenumber)
                        Case 2
                            BuildInstruction(linenumber) = TwoParts(Component.ComponentParts(0), Component.ComponentParts(1), linenumber)
                        Case 3
                            BuildInstruction(linenumber) = ThreeParts(Component.ComponentParts(0), Component.ComponentParts(1), Component.ComponentParts(2), linenumber)
                        Case Else
                            MsgBox("Error: Invalid Instruction - Instruction too long: " & Component.ComponentLength & " parts" & vbNewLine & "On Line " & linenumber, MsgBoxStyle.Exclamation)
                            Exit For
                    End Select

                End If
            End If

        Next linenumber
        'Unresolved Labels matched with Known Labels
        ResolveLabel()
    End Sub

    'Performs actual instruction tasks
    Sub ExecuteIns()
        Dim WhichIns As Integer
        Dim WhichAddress As Integer
        Dim numOfIterations As Integer = 0
        Dim Halt As Boolean = False 'Stop Condition

        Do Until Halt = True

            If IsDAT(ProgCount) = True Then
                'Increment one if instruction on line is Data
            Else
                'Stores 1st digit of 3 digit Machine Code
                WhichIns = BuildInstruction(ProgCount).ToString().Substring(0, 1)
                'Store 2nd & 3rd digit of 3 digit Machine Code
                WhichAddress = BuildInstruction(ProgCount).Substring(BuildInstruction(ProgCount).Length - 2)

                txtProgCount.Text = ProgCount
                Select Case WhichIns
                    Case 1 'ADD Adds value to accumulator value
                        Dim AccumValue As Integer = CInt(txtAccum.Text)
                        Dim AddValue As Integer = CInt(BuildInstruction(WhichAddress))
                        Dim Answer As Integer = AccumValue + AddValue
                        txtAccum.Text = Answer.ToString("D3")

                    Case 2 'SUB Subtracts value from accumulator value
                        Dim AccumValue As Integer = CInt(txtAccum.Text)
                        Dim SubValue As Integer = CInt(BuildInstruction(WhichAddress))
                        Dim Answer As Integer = AccumValue - SubValue
                        txtAccum.Text = Answer.ToString("D3")
                    Case 3 'STA Store the contents of the accumulator to address xx
                        Dim STAValue As Integer = txtAccum.Text
                        BuildInstruction(WhichAddress) = STAValue.ToString("D3")
                    Case 5 'LDA Loads contents of address xx onto accum
                        Dim LDAValue As Integer = BuildInstruction(WhichAddress)
                        txtAccum.Text = LDAValue.ToString("D3")
                    Case 6 'BRA Set program counter to address xx
                        ProgCount = WhichAddress - 1
                    Case 7 'BRZ If the contents of the accumulator are ZERO , set the program counter to address xx.
                        If CInt(txtAccum.Text) = 0 Then
                            ProgCount = WhichAddress - 1
                        End If
                    Case 8 'BRP If the contents of the accumulator are ZERO or positive (i.e. the negative flag is not set), set the program counter to address xx.
                        If CInt(txtAccum.Text) >= 0 Then
                            ProgCount = WhichAddress - 1
                        End If
                    Case 9 'INP or OUT
                        If BuildInstruction(ProgCount) = 901 Then 'INP Adds value from Input Box to Accumulator
                            If txtInput.Text = "" Then
                                MsgBox("Please enter a number in the Input Box")
                                'There are more INP
                                MoreINP = True
                                Exit Do
                                'Validation for INP
                            ElseIf txtInput.Text > 999 Then
                                MsgBox("The input is greater than 999." & vbNewLine & "Please enter an integer within the range of 0 and 999.")
                                Exit Do
                            Else
                                'Success Case
                                txtAccum.Text = txtInput.Text
                                txtInput.Text = ""
                            End If
                        Else 'OUT Copies value in accumulator to Output Box
                            txtOutput.Text = txtOutput.Text & txtAccum.Text & vbNewLine
                        End If

                    Case 0 'HLT Stops Program
                        Halt = True
                        AssignToMailBox()
                        MsgBox("Program Halted")
                End Select

            End If
            ProgCount += 1
            numOfIterations += 1

            'If all mailboxes are full
            If ProgCount > 35 Then
                Halt = True
                AssignToMailBox()
                MsgBox("Program Halted")
            End If

            'If loop gets too big (ie. infinite loop)
            If numOfIterations > 100 Then
                MsgBox("Error: Infinite Loop created", MsgBoxStyle.Exclamation)
                Halt = True
            End If
        Loop
    End Sub

    'Places Machine Code into Mail Boxes
    Sub AssignToMailBox()
        For i = 0 To BuildInstruction.Count - 1 Step 1
            Me.Controls("reg" & i).Text = BuildInstruction(i)
            If BuildInstruction(i) <> "000" Then
                Me.Controls("lab" & i).ForeColor = Color.Red
                Me.Controls("reg" & i).BackColor = Color.Orange
            End If
        Next

    End Sub

    'Prepares Machine Code to be placed into Mail Boxes
    Public Sub BuildInstructionFormatting(ByVal Where As Integer, ByVal Operand As Integer, ByVal FoundLabel As Boolean)
        Dim MemoryAddress As Integer = Operand
        Dim FormatMemoryAddress As String = MemoryAddress.ToString("D2")

        'Found label or not
        Select Case FoundLabel
            Case True
                'If Machine Code is not already 3 digits (e.g. not INP)
                If Len(BuildInstruction(Where)) <> 3 Then
                    BuildInstruction(Where) = BuildInstruction(Where) & FormatMemoryAddress
                End If
            Case False
                'Halt Program
                BuildInstruction(Where) = "000"

        End Select
    End Sub

    'Additional Features
    '===================

    'Validation for Input
    Private Sub txtInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter an integer within the range of 0 and 999")
            e.Handled = True
        End If
    End Sub

    'New Program (reset all)
    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        txtUserCode.Text = ""
        ResetProgram()
    End Sub

    Dim openFileDiag As OpenFileDialog = New OpenFileDialog()
    Dim saveFileDiag As New SaveFileDialog()

    'Opens files to load into program
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim myStream As Stream = Nothing
        openFileDiag.Title = "Open File"
        openFileDiag.InitialDirectory = "C:\"
        openFileDiag.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        openFileDiag.FilterIndex = 1
        openFileDiag.RestoreDirectory = True

        If openFileDiag.ShowDialog() = DialogResult.OK Then
            Try
                myStream = openFileDiag.OpenFile()
                If (myStream IsNot Nothing) Then
                    ReadFromFile(openFileDiag.FileName)
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If

    End Sub

    'Reads content of file specified by path and writes to user code text box
    Sub ReadFromFile(ByVal FilePath As String)
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(FilePath)
        txtUserCode.Text = fileReader
    End Sub

    'Saves program to a file
    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If saveFileDiag.FileName = "" Then
            SaveAs()
        Else
            My.Computer.FileSystem.WriteAllText(saveFileDiag.FileName, txtUserCode.Text, False)
        End If
    End Sub

    'Saves program to a new file
    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveAs()
    End Sub

    'File dialog for user to save file in path
    Sub SaveAs()
        Dim myStream As Stream
        saveFileDiag.Title = "Save As File"
        saveFileDiag.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        saveFileDiag.FilterIndex = 1
        saveFileDiag.RestoreDirectory = True

        If saveFileDiag.ShowDialog() = DialogResult.OK Then
            myStream = saveFileDiag.OpenFile()
            If (myStream IsNot Nothing) Then
                myStream.Close()
                My.Computer.FileSystem.WriteAllText(saveFileDiag.FileName, txtUserCode.Text, False)
            End If
        End If
    End Sub

    'Cut text
    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        If txtUserCode.SelectedText <> "" Then
            txtUserCode.Cut()
        End If
    End Sub

    'Copy text
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If txtUserCode.SelectionLength > 0 Then
            txtUserCode.Copy()
        End If
    End Sub

    'Paste text
    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        If Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) = True Then
            txtUserCode.Paste()
        End If
    End Sub

    'Undo last action
    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        If txtUserCode.CanUndo = True Then
            txtUserCode.Undo()
            txtUserCode.ClearUndo()
        End If
    End Sub

    'Select all text
    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        txtUserCode.SelectionStart = 0
        txtUserCode.SelectionLength = Len(txtUserCode.Text)
    End Sub

    'About Button
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutLMC.Show()
    End Sub

    'Shows list of available instructions
    Private Sub InstructionSetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstructionSetToolStripMenuItem.Click
        InstructionSet.Show()
    End Sub

    'Exits program
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

End Class
