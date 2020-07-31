Option Strict On
Imports System.IO
Imports Microsoft.Win32

'' Name: Hoang Quoc Bao Nguyen 
'' Date: 2020-07-30
'' Course: NETD-2202
'' Description: A program that is used to edit text

Public Class textEditorForm
    ''' <summary>
    ''' A method that copies the selected text into the clipboard for either copy or cut operation 
    ''' </summary>
    Sub CopyMethod()
        Dim txtBox As TextBox = TryCast(Me.ActiveControl, TextBox)
        If txtBox.SelectedText <> String.Empty Then
            Clipboard.SetData(DataFormats.Text, txtBox.SelectedText)
        End If
    End Sub

    ''' <summary>
    ''' A function that pop up a message box to ask if the user want to close the text editor 
    ''' </summary>
    ''' <returns></returns>
    Function ConfirmClose() As Boolean
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to close this ?", "Confirm Close", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Return True
        Else
            Return False
        End If
    End Function

#Region "File"
    ''' <summary>
    ''' Create a new page of text 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub NewClick(sender As Object, e As EventArgs) Handles mnuFileNew.Click
        If ConfirmClose() = True Then
            Call SaveClick(Me, e)
            txtTextArea.Text = String.Empty
        End If
    End Sub

    ''' <summary>
    ''' Open a saved text document from file explorer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OpenClick(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        If ConfirmClose() = True Then
            Call SaveClick(Me, e)
            Dim filestream As StreamReader

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                filestream = New StreamReader(openFileDialog.FileName)
                txtTextArea.Text = filestream.ReadToEnd()

                filestream.Close()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Save the text file after editing
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SaveClick(sender As Object, e As EventArgs) Handles mnuFileSave.Click
        saveDialog.Title = "Save Text Files"
        saveDialog.CheckFileExists = True
        saveDialog.CheckPathExists = True
        saveDialog.DefaultExt = "txt"
        saveDialog.RestoreDirectory = True

        Try
            txtTextArea.SaveFile(saveDialog.FileName, RichTextBoxStreamType.PlainText)
        Catch ex As Exception
            Call SaveAsClick(Me, e)
        End Try
    End Sub

    ''' <summary>
    ''' Save the text file as a new file 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SaveAsClick(sender As Object, e As EventArgs) Handles mnuFileSaveAs.Click
        saveDialog.Title = "Save Text Files"
        saveDialog.CheckFileExists = False
        saveDialog.CheckPathExists = False
        saveDialog.DefaultExt = "txt"
        saveDialog.RestoreDirectory = True

        Dim outputstream As StreamWriter

        If saveDialog.ShowDialog() = DialogResult.OK Then
            outputstream = New StreamWriter(saveDialog.FileName)
            outputstream.Write(txtTextArea.Text)

            outputstream.Close()
        End If
    End Sub

    ''' <summary>
    ''' Close the window form 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CloseClick(sender As Object, e As EventArgs) Handles mnuFileClose.Click
        If ConfirmClose() = True Then
            Call SaveClick(Me, e)
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' Exit from everything 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ExitClick(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        If ConfirmClose() = True Then
            Call SaveClick(Me, e)
            Application.Exit()
        End If

    End Sub
#End Region

#Region "Edit"
    ''' <summary>
    ''' Copy a selected text using CopyMethod()
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CopyClick(sender As Object, e As EventArgs) Handles mnuEditCopy.Click
        CopyMethod()
    End Sub

    ''' <summary>
    ''' Cut a selected text using CopyMethod() and delete it 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CutClick(sender As Object, e As EventArgs) Handles mnuEditCut.Click
        CopyMethod()
        txtTextArea.SelectedText = String.Empty
    End Sub

    ''' <summary>
    ''' Paste the former cut/copy selected text into the text editor
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PasteClick(sender As Object, e As EventArgs) Handles mnuEditPaste.Click
        Dim position As Integer = (CType(Me.ActiveControl, TextBox)).SelectionStart
        Me.ActiveControl.Text = Me.ActiveControl.Text.Insert(position, Clipboard.GetText())
    End Sub
#End Region

#Region "Help"
    ''' <summary>
    ''' Giving the information regarding information of course, Lab assignment, and student name 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutClick(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        Dim aboutModal As New aboutForm
        aboutModal.ShowDialog()
    End Sub

#End Region

End Class
