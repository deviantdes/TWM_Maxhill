'//  SAP MANAGE UI API 2005 SDK Sample
'//****************************************************************************
'//
'//  File:      SubMain.vb
'//
'//  Copyright (c) SAP MANAGE
'//
'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'// PARTICULAR PURPOSE.
'//
'//****************************************************************************
Option Strict Off
Option Explicit On 

Module SubMain
    Public Key As Byte() = New Byte() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 10, 73, 1, 5, 75, 1, 8}
    Public Sub Main()
        Try
            If checkInstance(True) = True Then
                Dim oTWM_customization As TWM_Maxhill
                oTWM_customization = New TWM_Maxhill()
                System.Windows.Forms.Application.Run()
            End If
        Catch ex As Exception
            MsgBox(ex.StackTrace & ":" & ex.Message)
        End Try
    End Sub
End Module
