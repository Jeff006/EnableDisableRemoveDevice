Imports Microsoft.Win32.SafeHandles
Imports EnableDisableRemoveDevice.NativeMethods

Friend Class SafeDeviceInfoSetHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    Sub New()
        MyBase.New(True)
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Return NativeMethods.SetupDiDestroyDeviceInfoList(Me.handle)
    End Function

End Class
