Attribute VB_Name = "mdlSAPConecta"
Function SapConecta()

'========= Abre conexão com o SAP
            On Error Resume Next
            strSAPID = session.ID
            If Err.Number <> 0 Then SAPConnected = False
            On Error GoTo ErrChk
            If Not (SAPConnected) Then
                On Error Resume Next
                k = 0
                While Not (SAPConnected) And k < 10
RetrySAPConn:
                    iErr = 0
                    Set SapGuiAuto = GetObject("SAPGUI")
                    Set SAPApp = SapGuiAuto.GetScriptingEngine
                    iErr = Err.Number
                    Set Connection = SAPApp.Children(0)
                    Set session = Connection.Children(0)
                    session.createsession
                    Set session2 = Connection.Children(1)
                    k = k + 1
                    If (k > 10) Then Stop 'SAP está encerrado?
                    If iErr <> 70 Then
                        SAPConnected = True
                    Else
                        Stop
                    End If
                Wend
                On Error GoTo ErrChk
            End If 'Finaliza abertura de conexão com o SAP
    

End Function
