VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Enum eChannel
    time = 0
    Incoming = 1
    Outgoing = 2
    Cursor = 3
    Spell = 4
    Inventory = 5
    FormEvents = 6
    Inventariovisible = 7
    Move = 8
    CotrolSM = 9
    EndRecording = 10
End Enum

Public Enum eFormEvents
    CloseForm = 0
    OpenForm = 1
End Enum

Private Frames() As tFrame
Public freq As Currency
Public RecordFile As Integer
Public isRecording As Boolean
Public startedTime As Long

Private CursorPos As Position
Private pItemSelected As Byte
Private pListIndex As Integer
Private pInventoryVisible As Boolean
Private FormMain As Form

Private Type tMiniForm
    Action As Byte
    LenghtName As Integer
    FormName As String
End Type

Private Type tFrame
    Channel As Byte
    Timing As Long
    X As Long
    Y As Long
    lenght As Integer
    data As String
    ListIndex As Integer
    ItemSelected As Byte
    miniForm As tMiniForm
    Visible As Boolean
    Direction As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                    ByRef destination As Any, _
                                    ByRef source As Any, _
                                    ByVal length As Long)
                                    
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private lastCursorTime As Long
Private lastSavedTime As Long
Private Buffer As String
Private writer As clsByteBuffer
Private reader As clsByteBuffer
Private handler As Integer

Private Sub Class_Initialize()
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
QueryPerformanceFrequency freq
Set writer = New clsByteBuffer
End Sub

Public Sub stopRecording()
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.EndRecording)
Call writer.saveBuffer
Close handler
isRecording = False
RecordFile = 0
End Sub

Public Sub Initialize(ByRef PrincipalForm As Form)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
'Seteo el form principal, lo hice por una buena razón que en este preciso no puedo recordar!
Set FormMain = PrincipalForm
End Sub

Public Function startRecording(ByVal lastChar As Integer, _
                            ByVal consoleText As String, _
                            ByVal UserMap As Integer, _
                            ByVal UserCharIndex As Integer) As Boolean
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Dim tmpChar As String
Dim i, X, Y As Integer
Dim temp As String
Dim length As Integer

'RecordFile = FreeFile()
handler = FreeFile()

isRecording = True
Me.startedTime = getHighFreqTime()

Open App.path & "\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".vao" For Binary As handler 'RecordFile
    Call writer.initializeWriter(handler) '(RecordFile)
    Call writer.putInteger(lastChar)
    
    For i = LBound(charlist) To lastChar
        With charlist(i)
            'position
            Call writer.putByte(.Active)
            
            If .Active Then
                Call writer.putLong(.Heading)
                Call writer.putLong(.Pos.X)
                Call writer.putLong(.Pos.Y)
                
                'char
                Call writer.putInteger(.iHead)
                Call writer.putInteger(.iBody)
                Call writer.putInteger(.Casco.Head(E_Heading.NORTH).GrhIndex)
                Call writer.putInteger(.Escudo.ShieldWalk(E_Heading.NORTH).GrhIndex)
                Call writer.putInteger(.Arma.WeaponWalk(E_Heading.NORTH).GrhIndex)
                
                Call writer.putBoolean(.UsandoArma)
                Call writer.putInteger(.FxIndex)
                
                Call writer.putByte(.Criminal)
                Call writer.putBoolean(.Atacable)

                Call writer.putString(.Nombre)
                
                Call writer.putByte(.Moving)
                Call writer.putSingle(.MoveOffsetX)
                Call writer.putSingle(.MoveOffsetY)
                    
                Call writer.putBoolean(.pie)
                Call writer.putBoolean(.muerto)
                Call writer.putBoolean(.invisible)
                Call writer.putByte(.priv)
            End If
        End With
    Next i
    
    'Put RecordFile, , UserMap
    Call writer.putInteger(UserMap)
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            With MapData(X, Y)
                Call writer.putByte(.Blocked)
                Call writer.putInteger(.CharIndex)
                Call writer.putInteger(.NPCIndex)
                Call writer.putInteger(.ObjGrh.GrhIndex)
                Call writer.putInteger(.Trigger)
            End With
        Next Y
    Next X
    
    For i = 1 To MAX_INVENTORY_SLOTS
        With Inventario
            Call writer.putInteger(.OBJIndex(i))
            
            If .OBJIndex(i) Then
                Call writer.putLong(.Amount(i))
                Call writer.putBoolean(.Equipped(i))
                Call writer.putInteger(.GrhIndex(i))
                Call writer.putInteger(.OBJType(i))
                Call writer.putInteger(.MaxHit(i))
                Call writer.putInteger(.MinHit(i))
                Call writer.putInteger(.MaxDef(i))
                Call writer.putInteger(.MinDef(i))
                Call writer.putSingle(.Valor(i))
                Call writer.putString(.ItemName(i))
            End If
        End With
    Next i
    
    Call writer.putByte(CByte(Inventario.SelectedItem))
    
    Call writer.putString(consoleText)
    Call writer.putInteger(UserCharIndex)
    
    Call writer.putInteger(UserMinSTA)
    Call writer.putInteger(UserMaxSTA)
    Call writer.putInteger(UserMinHP)
    Call writer.putInteger(UserMaxHP)
    Call writer.putInteger(UserMinMAN)
    Call writer.putInteger(UserMaxMAN)
    Call writer.putByte(UserMinHAM)
    Call writer.putByte(UserMaxHAM)
    Call writer.putByte(UserMinAGU)
    Call writer.putByte(UserMaxAGU)
    
    Call writer.putByte(UserFuerza)
    Call writer.putByte(UserAgilidad)
    
    Call writer.putLong(UserGLD)
    Call writer.putLong(UserExp)
    Call writer.putLong(UserPasarNivel)
    Call writer.putString(UserName)
    
    For i = 0 To 3
        Call writer.putBoolean(SMStatus(i))
    Next i
    
    startRecording = True
End Function

Public Sub getVideoFile(ByVal FileName As String)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Dim nFile, i, j, X, Y As Integer
Dim length As Integer
Dim richText As String
Dim buf As String

Dim OBJIndex As Integer
Dim Amount As Long
Dim Name As String
Dim Equipped As Boolean
Dim GrhIndex As Integer
Dim OBJType As Integer
Dim MaxHit As Integer
Dim MinHit As Integer
Dim MaxDef As Integer
Dim MinDef As Integer
Dim value As Single
Dim slot As Byte
Dim SR As RECT
Dim DR As RECT
Dim data() As Byte
nFile = FreeFile
i = 0
Open FileName For Binary As nFile
    Seek nFile, 1
    ReDim data(LOF(nFile)) As Byte
    Get nFile, , data
    
    Set reader = New clsByteBuffer
    Call reader.initializeReader(data)

    'Get nFile, , lastChar
    lastChar = reader.getInteger
    
    For j = LBound(charlist) To lastChar
        With charlist(j)
            'position
            'Get nFile, , .Active
            .Active = reader.getByte
            If .Active Then
                .Heading = reader.getLong
                .Pos.X = reader.getLong
                .Pos.Y = reader.getLong
                
                'char
                .iHead = reader.getInteger
                .iBody = reader.getInteger
                .Casco.Head(E_Heading.NORTH).GrhIndex = reader.getInteger
                .Escudo.ShieldWalk(E_Heading.NORTH).GrhIndex = reader.getInteger
                .Arma.WeaponWalk(E_Heading.NORTH).GrhIndex = reader.getInteger
                .UsandoArma = reader.getBoolean
                .FxIndex = reader.getInteger

                .Criminal = reader.getByte
                .Atacable = reader.getBoolean
                
                .Nombre = reader.getString()
                
                .Moving = reader.getByte
                .MoveOffsetX = reader.getSingle
                .MoveOffsetY = reader.getSingle
                              
                .pie = reader.getBoolean
                .muerto = reader.getBoolean
                .invisible = reader.getBoolean
                .priv = reader.getByte
                
                'cargo los graficos en el char
                If .iHead > 0 Then .Head = HeadData(.iHead)
                If .iBody > 0 Then .Body = BodyData(.iBody)
                
                .Arma = WeaponAnimData(getWeaponByGrhIndex(.Arma.WeaponWalk(E_Heading.NORTH).GrhIndex, E_Heading.NORTH))
                .Escudo = ShieldAnimData(getShieldByGrhIndex(.Escudo.ShieldWalk(E_Heading.NORTH).GrhIndex, E_Heading.NORTH))
                .Casco = CascoAnimData(getCascoByGrhIndex(.Casco.Head(E_Heading.NORTH).GrhIndex, E_Heading.NORTH))
                
                If .FxIndex > 0 Then
                    Call InitGrh(.fX, .FxIndex)
                End If
            End If
        End With
    Next j
    
    'usermap
    UserMap = reader.getInteger
    
    Call SwitchMap(UserMap)
        
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            With MapData(X, Y)
                .Blocked = reader.getByte
                .CharIndex = reader.getInteger
                .NPCIndex = reader.getInteger
                .ObjGrh.GrhIndex = reader.getInteger
                .Trigger = reader.getInteger
            End With
        Next Y
    Next X
    
    For slot = 1 To MAX_INVENTORY_SLOTS
        OBJIndex = reader.getInteger
        
        If OBJIndex Then
            Amount = reader.getLong
            Equipped = reader.getBoolean
            GrhIndex = reader.getInteger
            OBJType = reader.getInteger
            MaxHit = reader.getInteger
            MinHit = reader.getInteger
            MaxDef = reader.getInteger
            MinDef = reader.getInteger
            value = reader.getSingle
            Name = reader.getString
            
            If Equipped Then
                Select Case OBJType
                    Case eObjType.otWeapon
                        frmMain.lblWeapon = MinHit & "/" & MaxHit
                        UserWeaponEqpSlot = slot
                    Case eObjType.otArmadura
                        frmMain.lblArmor = MinDef & "/" & MaxDef
                        UserArmourEqpSlot = slot
                    Case eObjType.otescudo
                        frmMain.lblShielder = MinDef & "/" & MaxDef
                        UserHelmEqpSlot = slot
                    Case eObjType.otcasco
                        frmMain.lblHelm = MinDef & "/" & MaxDef
                        UserShieldEqpSlot = slot
                End Select
            Else
                Select Case slot
                    Case UserWeaponEqpSlot
                        frmMain.lblWeapon = "0/0"
                        UserWeaponEqpSlot = 0
                    Case UserArmourEqpSlot
                        frmMain.lblArmor = "0/0"
                        UserArmourEqpSlot = 0
                    Case UserHelmEqpSlot
                        frmMain.lblShielder = "0/0"
                        UserHelmEqpSlot = 0
                    Case UserShieldEqpSlot
                        frmMain.lblHelm = "0/0"
                        UserShieldEqpSlot = 0
                End Select
            End If
            
            Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, Name)
        End If
    Next slot
    
    slot = reader.getByte
    Call Inventario.setSelectedItem(slot)
    
    'console text
    frmMain.RecTxt.TextRTF = reader.getString
    
    UserCharIndex = reader.getInteger
    UserMinSTA = reader.getInteger
    UserMaxSTA = reader.getInteger
    UserMinHP = reader.getInteger
    UserMaxHP = reader.getInteger
    UserMinMAN = reader.getInteger
    UserMaxMAN = reader.getInteger
    UserMinHAM = reader.getByte
    UserMaxHAM = reader.getByte
    UserMinAGU = reader.getByte
    UserMaxAGU = reader.getByte
        
    Call UpdateBar(frmMain.shpEnergia, frmMain.lblEnergia, UserMinSTA, UserMaxSTA)
    Call UpdateBar(frmMain.shpVida, frmMain.lblVida, UserMinHP, UserMaxHP)
    Call UpdateBar(frmMain.shpMana, frmMain.lblMana, UserMinMAN, UserMaxMAN)
    Call UpdateBar(frmMain.shpHambre, frmMain.lblHambre, UserMinHAM, UserMaxHAM)
    Call UpdateBar(frmMain.shpSed, frmMain.lblSed, UserMinAGU, UserMaxAGU)
    
    UserFuerza = reader.getByte
    UserAgilidad = reader.getByte
    UserGLD = reader.getLong
    UserExp = reader.getLong
    UserPasarNivel = reader.getLong
    UserName = reader.getString
    
    frmMain.lblExp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
    frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
    
    If UserGLD >= CLng(UserLvl) * 10000 Then
        'Changes color
        frmMain.GldLbl.ForeColor = &HFF& 'Red
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow
    End If
    
    frmMain.GldLbl.Caption = UserGLD
    frmMain.lblName.Caption = UserName
    
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    frmMain.lblDext.ForeColor = getDexterityColor()
    
    For i = 0 To 3
        'Get nFile, , SMStatus(i)
        SMStatus(i) = reader.getBoolean
        
        GrhIndex = GRH_INI_SM + i + SM_CANT * (CInt(SMStatus(i)) + 1)
        
        With GrhData(GrhIndex)
            SR.left = .sX
            SR.right = SR.left + .pixelWidth
            SR.top = .sY
            SR.bottom = SR.top + .pixelHeight
            
            DR.left = 0
            DR.right = .pixelWidth
            DR.top = 0
            DR.bottom = .pixelHeight
        End With
        
        Call DrawGrhtoHdc(frmMain.picSM(i).HDC, GrhIndex, SR, DR)
        frmMain.picSM(i).Refresh
    Next i
    
    UserPos.X = charlist(UserCharIndex).Pos.X
    UserPos.Y = charlist(UserCharIndex).Pos.Y
    
    frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
    
    Do While Not reader.EOF 'EOF(nFile)
        ReDim Preserve Frames(i) As tFrame
        
        With Frames(i)
            .Channel = reader.getByte
            .Timing = reader.getLong
            
            Select Case .Channel
                Case eChannel.Incoming, eChannel.Outgoing
                    Frames(i).data = reader.getString
                    
                Case eChannel.Cursor
                    .X = reader.getLong
                    .Y = reader.getLong
                    
                Case eChannel.Spell
                    .ListIndex = reader.getInteger
                    
                Case eChannel.Inventory
                    .ItemSelected = reader.getByte
                    
                Case eChannel.FormEvents
                    .miniForm.Action = reader.getByte
                    .miniForm.FormName = reader.getString
                    
                Case eChannel.Inventariovisible
                    .Visible = reader.getBoolean
                    
                Case eChannel.Move
                    .Direction = reader.getByte
                    
                Case eChannel.CotrolSM
                    .ItemSelected = reader.getByte
                    .Visible = reader.getBoolean
            End Select
        End With
        
        i = i + 1
    Loop

    Set reader = Nothing
Close nFile

End Sub

'Devuelvo el tiempo actual de la reproducción.
Public Function getCurrentTime() As Long
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
getCurrentTime = getHighFreqTime() - Me.startedTime
End Function

Public Sub putBuffer(ByVal Channel As Byte, _
                    Optional ByVal data As String, _
                    Optional ByVal MousePosX As Long, Optional ByVal MousePosY As Long, _
                    Optional ByVal ListIndex As Integer, _
                    Optional ByVal ItemSelected As Byte, _
                    Optional ByVal Action As Byte, Optional ByVal FormName As String, _
                    Optional ByVal Visible As Boolean, _
                    Optional ByVal Direction As Byte)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Dim lenght As Integer
Static con As Integer

'On Error Resume Next
If handler <= 0 Then Exit Sub

'El canal y el timestamp DEBEN estar en todos los paquetes.
Call writer.putByte(Channel)
Call writer.putLong(getCurrentTime())

Select Case Channel
    Case eChannel.Incoming, eChannel.Outgoing
        Call writer.putString(data)
        
    Case eChannel.Cursor
        Debug.Print con
        con = con + 1
        Call writer.putLong(MousePosX)
        Call writer.putLong(MousePosY)
        
    Case eChannel.Spell
        Call writer.putInteger(ListIndex)
        
    Case eChannel.Inventory
        Call writer.putByte(ItemSelected)
        
    Case eChannel.FormEvents
        Call writer.putByte(Action)
        Call writer.putString(FormName)
        
    Case eChannel.Inventariovisible
        Call writer.putBoolean(Visible)
        
    Case eChannel.Move
        Call writer.putByte(Direction)
        
    Case eChannel.CotrolSM
        Call writer.putByte(ItemSelected)
        Call writer.putBoolean(Visible)
End Select
End Sub

Public Sub addDataPacket(ByVal Channel As Byte, ByVal pData As String)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(Channel, pData)
Debug.Print pData
End Sub

Public Sub addCursor(ByVal X As Long, ByVal Y As Long)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.Cursor, , X, Y)
End Sub

Public Sub addSpellIndex(ByVal ListIndex As Integer)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.Spell, , , , ListIndex)
End Sub

Public Sub addItemSelected(ByVal slot As Byte)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.Inventory, , , , , slot)
End Sub

Public Sub addFormEvent(ByVal Action As Byte, ByVal FormName As String)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.FormEvents, , , , , , Action, FormName)
End Sub

Public Sub addInventoryVisible(ByVal Visible As Boolean)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.Inventariovisible, , , , , , , , Visible)
End Sub

Public Sub addMove(ByVal Direction As Byte)
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Call putBuffer(eChannel.Move, , , , , , , , , Direction)
End Sub

Public Sub addControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
'***************************************************
'Author: Pato
'Last Modify Date: 06/14/10
'
'***************************************************
Call putBuffer(eChannel.CotrolSM, , , , , Index, , , Mostrar)
End Sub

Public Sub checkChanges()
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
If handler <= 0 Then Exit Sub
'If GetTickCount() - lastCursorTime > 50 Then
'    If Not CursorPos.X = FormMain.MouseX Or Not CursorPos.Y = FormMain.MouseY Then
'        CursorPos.X = FormMain.MouseX
'        CursorPos.Y = FormMain.MouseY
'        Call addCursor(CursorPos.X, CursorPos.Y)
'        lastCursorTime = GetTickCount()
'    End If
'End If
If Not pItemSelected = Inventario.SelectedItem Then
    pItemSelected = Inventario.SelectedItem
    Call addItemSelected(pItemSelected)
End If

If Not pListIndex = FormMain.hlst.ListIndex Then
    pListIndex = FormMain.hlst.ListIndex
    Call addSpellIndex(pListIndex)
End If

If Not FormMain.PicInv.Visible = pInventoryVisible Then
    pInventoryVisible = FormMain.PicInv.Visible
    Call addInventoryVisible(pInventoryVisible)
End If
End Sub

Public Sub Play()
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
On Error GoTo ErrHandler
    Const MAX_PREFRAME As Long = 4
    Const KEY_FRAME As Long = 2
    
    Dim i As Long
    Dim CantFrames As Long
    Dim nextTime As Long
    
    CantFrames = UBound(Frames())
    
    EngineRun = True
    FormMain.Show
    DoEvents
    
    Me.startedTime = getHighFreqTime()
    nextTime = Frames(0).Timing
    
    Do
        With Frames(i)
            If (getCurrentTime >= nextTime) And (UserMoving = 0) Then
                nextTime = Frames(i + 1).Timing
                
                Select Case .Channel
                    Case eChannel.Incoming, eChannel.Outgoing
                        Dim arr() As Byte
                        
                        If .data = vbNullString Then Exit Sub
                        
                        arr = StrConv(.data, vbFromUnicode)
                        Call incomingData.WriteBlock(arr)
                        
                        If .Channel = eChannel.Incoming Then
                            Call HandleIncomingData
                            'Call incomingData.CleanBuffer
                        End If
                        
                    Case eChannel.Cursor
                        CursorPos.X = .X
                        CursorPos.Y = .Y
                        
                    Case eChannel.Spell
                        FormMain.hlst.ListIndex = .ListIndex
                        
                    Case eChannel.Inventory
                        Inventario.setSelectedItem (.ItemSelected)
                        
                    Case eChannel.FormEvents
                        Select Case .miniForm.Action
                            Case eFormEvents.CloseForm
                                Unload getForm(.miniForm.FormName)
                                
                            Case eFormEvents.OpenForm
                                getForm(.miniForm.FormName).Show
                        End Select
                        
                    Case eChannel.Inventariovisible
                        'Hacer esto con una función en el frmMain, así limpiamos el código
                        Call Audio.PlayWave(SND_CLICK)
                        
                        FormMain.PicInv.Visible = .Visible 'and true
                        FormMain.hlst.Visible = Not .Visible
                        FormMain.cmdInfo.Visible = Not .Visible
                        FormMain.CmdLanzar.Visible = Not .Visible
                        
                        If .Visible Then
                            FormMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario.jpg")
                        Else
                            FormMain.InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos.jpg")
                        End If
                        
                    Case eChannel.Move
                        Call MoveTo(.Direction)
                        frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                
                    Case eChannel.CotrolSM
                        Call frmMain.ControlSM(.ItemSelected, .Visible)
                End Select
                
                i = i + 1
                Call Conversor.resetFrame
            End If
            
            Call ShowNextFrame(FormMain.top, FormMain.left, CursorPos.X, CursorPos.Y)
            'Call RenderSounds
            
            If Conversor.Converting Then
                Call Conversor.addFrame
                Call Conversor.incrementFrame
            End If
        End With
        
        DoEvents
    Loop While (i < CantFrames)
    
    If Conversor.Converting Then
        Conversor.closeFile
        Conversor.Converting = False
        
        MsgBox "Conversión terminada."
    Else
        MsgBox "Fin del video."
    End If
    
    Exit Sub
    
ErrHandler:
    Conversor.closeFile
    Conversor.Converting = False
    MsgBox Err.number & " " & Err.Description
End Sub

Public Function getForm(ByVal Name As String) As Form
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
Dim frm As Form

Load frmOpciones

For Each frm In Forms
    If frm.Name = Name Then
        Set getForm = frm
        Exit Function
    End If
Next
End Function

Public Function getHighFreqTime() As Long
'***************************************************
'Author: Budi
'Last Modify Date: 29/01/10
'
'***************************************************
    Dim ret As Currency
    
    Call QueryPerformanceCounter(ret)
    getHighFreqTime = Fix(ret)
End Function

Private Sub UpdateBar(ByVal Shp As Shape, ByVal Lbl As Label, ByVal min As Integer, ByVal max As Integer)
Dim bWidth As Byte
    
Lbl = min & "/" & max

bWidth = (((min / 100) / (max / 100)) * 75)

Shp.Visible = (bWidth <> 75)
Shp.Width = 75 - bWidth
Shp.left = 584 + (75 - Shp.Width)
End Sub
