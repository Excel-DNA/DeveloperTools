Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

<ComVisible(True)>
Public Class MyRibbon
    Inherits ExcelRibbon

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Using stream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("$safeprojectname$.MyRibbon.xml")
            Using reader As New StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function

    Public Sub OnButtonPressed(control As IRibbonControl)
        System.Windows.Forms.MessageBox.Show("Hello!")
    End Sub

End Class