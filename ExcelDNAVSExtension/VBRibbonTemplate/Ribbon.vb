Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

<ComVisible(True)>
Public Class $safeitemname$
    Inherits ExcelRibbon

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Return $fileinputname$Resources.Ribbon
    End Function

    Public Overrides Function LoadImage(imageId As String) As Object
        ' This will return the image resource with the name specified in the image='xxxx' tag
        Return $fileinputname$Resources.ResourceManager.GetObject(imageId)
    End Function

    Public Sub OnButtonPressed(control As IRibbonControl)
        System.Windows.Forms.MessageBox.Show("Hello!")
    End Sub

End Class