Attribute VB_Name = "Module1"
Option Explicit
'Variables for SAP GUI Tool
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
'Variables for Functions
Public Plant, SAP_CODE, Listing_Procedure As String
Dim W_System
Dim iCtr As Integer
'Transactions Code
Const tcode = "WSM3"

'Function to Connect with SAP GUI Sessions
Function Create_SAP_Session() As Boolean
    'Variables for Session Creation
    Dim il, it
    Dim W_conn, W_Sess, tcode, Transac, Info_System
    Dim N_Gui As Integer
    Dim A1, A2 As String

    'Get Transaction Code
    tcode = Sheets(1).Range("B3")

    'Get System Name in Cell(2,1) of Sheet1
    If mysystem = "" Then
        W_System = Sheets(1).Cells(2, 2)
    Else
        W_System = mysystem
    End If

    'If we are already connected to a Session we exit this function
    If W_System = "" Then
    Create_SAP_Session = False
    Exit Function
    End If

    'If Object Session is not null and the system is matching with the one we target: we use this object
    If Not session Is Nothing Then
        If session.Info.SystemName & session.Info.Client = W_System Then
            Create_SAP_Session = True
            Exit Function
        End If
    End If

    'If we are not connected to anything and GUI Object is Nothing we create one
    If objGui Is Nothing Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    End If

    'Loop through all SAP GUI Sessions to find the one with the right transaction
    For il = 0 To objGui.Children.Count - 1
        Set W_conn = objGui.Children(il + 0)
        
        For it = 0 To W_conn.Children.Count - 1
            Set W_Sess = W_conn.Children(it + 0)
            Transac = W_Sess.Info.Transaction
            Info_System = W_Sess.Info.SystemName & W_Sess.Info.Client
            
            'Check if Session Name and Transaction Code are matching then connect to it
            If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System Then
            'If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System And W_Sess.Info.Transaction = tcode Then
                Set objConn = objGui.Children(il + 0)
                Set session = objConn.Children(it + 0)
                Exit For
            End If
            
        Next
        
    Next

    ' If we can't find Session with the right System Name and Transaction Code: display error message
    If session Is Nothing Then
    MsgBox "No active session to system " + W_System + " with transaction " + tcode + ", or scripting is not enabled.", vbCritical + vbOKOnly
    Create_SAP_Session = False
    Exit Function
    End If

    ' Turn on scripting mode
    If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject objGui, "on"
    End If

    'Confirm connection to a session
    Create_SAP_Session = True
End Function

' ---
'Procedure to perform Listing
Public Sub Listing_Process()
'Variable for Listing Process
Dim W_Src_Ord
Dim W_Ret As Boolean
Dim N As Integer
Dim N_max As Integer
Set objSheet = ActiveWorkbook.Sheets(1)
' Connect to a system
W_Ret = Create_SAP_Session()

'If Create_SAP_Session Return Nothing message
If Not W_Ret Then
    MsgBox "Not connected to client"
End If

'Loop Line by Line
N = 2
While Not (Sheets("Listing").Cells(N + 1, 1) = "")
    'Get Parameters from Excel Sheet
    SAP_CODE = Sheets("Listing").Cells(N, 1)
    Plant = Sheets("Listing").Cells(N, 2)
    Listing_Procedure = Sheets("Listing").Cells(N, 3)
    
    'Call the function
    Call Listing_Function(Plant, SAP_CODE, Listing_Procedure, N)
    
    'Confirm with a "V"
    Sheets("Listing").Cells(N, 4) = "V"
    N = N + 1 
Wend

End Sub

' ---- 
'Function for Listing
Function Listing_Function(Plant, SAP_CODE, Listing_Procedure, N)
    'session.findById("wnd[0]").Maximize: if you want to maximize the screen
    'Call the transaction
    session.findById("wnd[0]/tbar[0]/okcd").Text = "wsm3"
    session.findById("wnd[0]").sendVKey 0
    'Ticking on Listing
    session.findById("wnd[0]/usr/chkLSTFLMAT").Selected = True
    session.findById("wnd[0]/usr/chkLIEFWERK").Selected = True
    'Filling Plant, SAP Code and Listing Procedure
    session.findById("wnd[0]/usr/ctxtASORT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtASORT-LOW").Text = Plant
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = SAP_CODE
    session.findById("wnd[0]/usr/ctxtLSTFL").Text = ""
    session.findById("wnd[0]/usr/ctxtLSTFL").Text = Listing_Procedure
    'Validate
    session.findById("wnd[0]/usr/chkLIEFWERK").SetFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    'Getting Start Date from Window
    Sheets("Listing").Cells(N, 5) = session.findById("wnd[0]/usr/lbl[0,0]").Text
    Application.Wait (Now + TimeValue("0:00:1"))
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    Application.Wait (Now + TimeValue("0:00:2"))
End Function

