Imports System.Xml
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Runtime.Serialization




Public Class Form1

    Private ds As DataSet
    Private boolChangesMade As Boolean = False

    Private strXmlFile As String


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Application.ProductName


    End Sub




    Private Sub LoadFile()
        If boolChangesMade Then CloseFile()

        If Not IO.File.Exists(strXmlFile) Then
            MsgBox("Could not find file:  " & strXmlFile, MsgBoxStyle.Exclamation, Application.ProductName)
            Exit Sub
        End If

        Dim fi As IO.FileInfo = New IO.FileInfo(strXmlFile)
        If Not fi.Extension.ToLower = ".xml" Then
            MsgBox("Please select a file with an xml extension.", MsgBoxStyle.Exclamation, Application.ProductName)
            Exit Sub
        End If


        ds = New DataSet

        Dim xmlFile As XmlReader

        xmlFile = XmlReader.Create(strXmlFile, New XmlReaderSettings)

        Try
            ds.ReadXml(xmlFile)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, Application.ProductName)
            xmlFile.Close()
            Exit Sub
        End Try


        xmlFile.Close()

        Dim intTableNum As Integer = 0

        If ds.Tables.Count > 1 Then
            MsgBox("Sorry, XML file has more than 1 table in it." & vbCrLf & vbCrLf & "This app can only work with 1 table.", MsgBoxStyle.Information, Application.ProductName)
            Exit Sub
        End If

        BindingSource1.DataSource = ds.Tables(intTableNum)
        DataGridView1.DataSource = BindingSource1

        Me.Text = Application.ProductName & " - " & fi.FullName

        boolChangesMade = False

    End Sub

    Private Sub CloseFile()
        If boolChangesMade Then
            Dim ret As Integer = MsgBox("Save changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Save changes?")
            If ret = vbYes Then SaveFile()
        End If

        DataGridView1.DataSource = Nothing
        BindingSource1.DataSource = Nothing
        ds = Nothing

        boolChangesMade = False
        Me.Text = Application.ProductName
    End Sub


    Private Sub SaveFile()

        Dim dt As New DataTable

        BindingSource1.EndEdit()

        Dim settings As XmlWriterSettings = New XmlWriterSettings()
        settings.Indent = True
        Dim writer As XmlWriter = XmlWriter.Create(strXmlFile, settings)
        ds.WriteXml(writer)
        writer.Close()

        boolChangesMade = False

    End Sub



    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If boolChangesMade Then
            Dim ret As Integer = MsgBox("Exit without saving?", _
                                        MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Exit " & Application.ProductName & "?")
            If ret <> vbYes Then e.Cancel = True
        End If
    End Sub
    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        boolChangesMade = True
        Debug.WriteLine("cell end edit")
    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        Debug.WriteLine("row added")
        boolChangesMade = True
    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        Debug.WriteLine("row removed")
        boolChangesMade = True
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        CloseFile()
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            strXmlFile = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If

        LoadFile()
    End Sub

    
    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveFile()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            strXmlFile = SaveFileDialog1.FileName
            SaveFile()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        CloseFile()
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            If boolChangesMade Then CloseFile()
            strXmlFile = path
            LoadFile()
            Exit For
        Next
    End Sub

End Class
