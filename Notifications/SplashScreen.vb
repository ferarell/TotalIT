﻿Public Class SplashScreen
    Sub New
        InitializeComponent()
        TextToSpeak("Iniciando el notificador, espere por favor...")
    End Sub

    Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
        MyBase.ProcessCommand(cmd, arg)
    End Sub

    Public Enum SplashScreenCommand
        SomeCommandId
    End Enum
End Class