Module labelHandling

    'Temp values to be put in Unresolved & Label list
    Public TempLabelRef As LabelRef
    Public TempUnresolvedLabel As UnresolvedLabel

    Public LabelRefList = New List(Of LabelRef)
    Public UnresolvedLabelList = New List(Of UnresolvedLabel)

    Structure LabelRef
        Dim LabelText As String
        Dim LabelAddress As Integer
    End Structure

    Structure UnresolvedLabel
        Dim LabelText As String
        Dim Instruction As String
        Dim LabelAddress As Integer
    End Structure

    'Handles Label (Unresolved and LabelRef) Matchup
    Sub ResolveLabel()
        Dim UnresolvedCount As Integer
        Dim EndofUnresolved As Integer = UnresolvedLabelList.count - 1
        Dim LabelRefCount As Integer
        Dim EndofLabelRef As Integer = LabelRefList.count - 1

        Dim increment As Boolean

        'Search through Unresolved Labels and LabelRef to find matches
        While UnresolvedCount <= EndofUnresolved

            If UnresolvedLabelList(UnresolvedCount).labeltext = "" Then
                UnresolvedCount += 1
            Else
                'Search for Unresolved(i) with all of LabelRef
                For LabelRefCount = 0 To EndofLabelRef

                    'Match Found
                    If UnresolvedLabelList(UnresolvedCount).labeltext = LabelRefList(LabelRefCount).labeltext Then
                        increment = True
                        LCC.BuildInstructionFormatting(UnresolvedCount, LabelRefCount, True)
                    End If

                    'Skip to next Unresolved if at end
                    If increment = True And LabelRefCount = EndofLabelRef Then
                        UnresolvedCount += 1
                    End If

                    'Error if not found Label
                    If UnresolvedLabelList(EndofUnresolved).labeltext <> LabelRefList(EndofLabelRef).labeltext And LabelRefCount = EndofLabelRef And increment = False Then
                        MsgBox("Error: Not Found For " & UnresolvedLabelList(UnresolvedCount).labeltext & vbNewLine & "On Line " & UnresolvedCount, MsgBoxStyle.Exclamation)
                        LCC.BuildInstructionFormatting(UnresolvedCount, LabelRefCount, False)
                        Exit While
                    End If

                Next LabelRefCount
            End If

        End While

    End Sub
End Module
