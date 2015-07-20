Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
''' <summary>
''' Coded by Caleb Woolley, uploaded via GitHub 2015.
''' </summary>
''' <remarks></remarks>
Public Class Form1
    Dim FilePath As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CheckForIllegalCrossThreadCalls = False
        Dim FSize As String = GetFileSize(Application.ExecutablePath)
        GroupBox1.Text = "Checksums - " & Application.ExecutablePath.Substring(Application.ExecutablePath.LastIndexOf("\") + 1, Application.ExecutablePath.Length - (Application.ExecutablePath.LastIndexOf("\") + 1)) & " (" & If(Format(FSize / 1024, "#0.00") = 0, FSize & " bytes", Format(FSize / 1024, "#0.00") & " KB") & ")"
        Label2.Text = "MD5: " & MD5CalcFile(Application.ExecutablePath)
        Label3.Text = "SHA1: " & SHA1File(Application.ExecutablePath)
        Label4.Text = "SHA256: " & SHA256File(Application.ExecutablePath)

        ' Second app
        Dim doneWithInit As New Threading.EventWaitHandle(False, Threading.EventResetMode.ManualReset, "MyWaitHandle")

        ' Here, the second application initializes what it needs to.
        ' When it's done, it signals the wait handle:
        doneWithInit.[Set]()
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        On Error Resume Next
        Dim file_names As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
        FilePath = file_names(0)
        If IsDirectory(FilePath) Then Exit Sub
        Dim T As New Threading.Thread(AddressOf GatherInfo)
        T.Start()
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

#Region " Threading "
    Public Sub GatherInfo()
        Me.Text = "File Checksum Generator"
        Label2.Text = "MD5: Loading..."
        Label3.Text = "SHA1: Loading..."
        Label4.Text = "SHA256: Loading..."
        Dim FSize As String = GetFileSize(FilePath)
        GroupBox1.Text = "Checksums - " & FilePath.Substring(FilePath.LastIndexOf("\") + 1, FilePath.Length - (FilePath.LastIndexOf("\") + 1)) & " (" & If(Format(FSize / 1024, "#0.00") = 0, FSize & " bytes", Format(FSize / 1024, "#0.00") & " KB") & ")"

        Dim MD5 As New Threading.Thread(AddressOf MD5Thread)
        Dim SHA1 As New Threading.Thread(AddressOf SHA1Thread)
        Dim SHA256 As New Threading.Thread(AddressOf SHA256Thread)
        MD5.Start()
        SHA1.Start()
        SHA256.Start()

        Me.Text = "File Checksum Generator"
    End Sub

    Public Sub MD5Thread()
        Label2.Text = "MD5: " & MD5CalcFile(FilePath)
    End Sub
    Public Sub SHA1Thread()
        Label3.Text = "SHA1: " & SHA1File(FilePath)
    End Sub
    Public Sub SHA256Thread()
        Label4.Text = "SHA256: " & SHA256File(FilePath)
    End Sub
#End Region

#Region " Copying "
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click, Label3.Click, Label4.Click
        If MessageBox.Show("Copy this hash?", "Checksum Copying", MessageBoxButtons.YesNo) = vbYes Then
            Clipboard.SetText(sender.Text.SubString(sender.Text.IndexOf(" ") + 1))
        End If
    End Sub
#End Region

#Region " Functions "
    Public Function IsDirectory(ByRef Path As String) As Boolean
        Dim attr As FileAttributes = File.GetAttributes(Path)
        If (attr And FileAttributes.Directory) = FileAttributes.Directory Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MD5CalcFile(ByVal filepath As String) As String
        Using reader As New System.IO.FileStream(filepath, IO.FileMode.Open, IO.FileAccess.Read)
            Using md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
                Dim hash() As Byte = md5.ComputeHash(reader)
                Return ByteArrayToString(hash)
            End Using
        End Using
    End Function
    Public Function SHA1File(ByRef Path As String) As String
        Dim SHA1 As New SHA1CryptoServiceProvider()
        Dim myfile As Byte() = File.ReadAllBytes(Path)
        Return ByteArrayToString(SHA1.ComputeHash(myfile))
    End Function
    Public Function SHA256File(ByRef Path As String) As String
        Dim sha As New SHA256Managed()
        Dim myfile As Byte() = File.ReadAllBytes(Path)
        Return ByteArrayToString(sha.ComputeHash(myfile))
    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        Dim sb As New System.Text.StringBuilder(arrInput.Length * 2)
        For i As Integer = 0 To arrInput.Length - 1
            sb.Append(arrInput(i).ToString("X2"))
        Next
        Return sb.ToString().ToLower
    End Function

    Private Function GetFileSize(ByVal MyFilePath As String) As Long
        Dim MyFile As New FileInfo(MyFilePath)
        Dim FileSize As Long = MyFile.Length
        Return FileSize
    End Function
#End Region


    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub
End Class


