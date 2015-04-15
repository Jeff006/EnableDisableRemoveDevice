Public Class Form1

    Private Sub btnEnable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnable.Click
        DeviceHelper.SetDeviceEnabled("COM" + CStr(ComboBox1.SelectedItem), True)
    End Sub

    Private Sub btnDisable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisable.Click
        DeviceHelper.SetDeviceEnabled("COM" + CStr(ComboBox1.SelectedItem), False)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        DeviceHelper.SetDeviceRemove("COM" + CStr(ComboBox1.SelectedItem))
    End Sub

End Class
