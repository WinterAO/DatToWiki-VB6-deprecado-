Attribute VB_Name = "ModConversor"
Option Explicit

'OBJTYPES
Public Const TypeCasco As Byte = 17
Public Const TypeArmadura As Byte = 3
Public Const TypeEscudos As Byte = 16
Public Const TypeArmas As Byte = 2

'Numero total de clases
Public Const NUMCLASES As Byte = 12

Public ObjDatDir As String
Public NPCDatDir As String
Public HechizosDatDir As String

Sub Main()
    'Rutas
    ObjDatDir = "F:\Proyectos\Proyecto Winter\Server-WinterAO\Dat\OBJ.dat"
    NPCDatDir = "F:\Proyectos\Proyecto Winter\Server-WinterAO\Dat\NPCs.dat"
    HechizosDatDir = "F:\Proyectos\Proyecto Winter\Server-WinterAO\Dat\Hechizos.dat"
    
    frmmain.Show
End Sub

Public Sub ConvertCascos()
'***********************************************
'Autor: Lorwik
'Fecha: 09/05/2020
'Descripcion: Lee los cascos de Obj.dat y los pasa a una tabla de Wiki
'***********************************************
    Dim TotalObj As Long
    Dim i As Long
    Dim j As Byte
    Dim TipoObjeto As Integer
    Dim Clase As String
    Dim Resultado As String
    
    Dim Leer As New clsIniManager
    
    'Limpiamos el txt
    frmmain.txtTxtResultado.Text = ""

    Call Leer.Initialize(ObjDatDir)
    
    TotalObj = Val(Leer.GetValue("INIT", "NumOBJs"))
    
    'Preparamos la cabecera de la tabla
    Resultado = "{| class='wikitable'" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    Resultado = Resultado + "!Imagen !! Nombre !! Min !! Max !! Nivel !! Precio !! Clases no Permitidas" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    
    'Reccoremos todos los objetos en busca de cascos
    For i = 1 To TotalObj
        
        TipoObjeto = Val(Leer.GetValue("OBJ" & i, "ObjType"))
        
        '¿Es un casco?
        If TipoObjeto = TypeCasco Then
            
            Resultado = Resultado + "|[[Archivo:" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png|thumb]] || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
            Resultado = Resultado + " || " & Val(Leer.GetValue("OBJ" & i, "Nivel")) & " || " & Val(Leer.GetValue("OBJ" & i, "Valor")) & "||"
            
             For j = 1 To NUMCLASES
                Clase = Leer.GetValue("OBJ" & i, "CP" & j)
                If Not Clase = "" Then
                    Resultado = Resultado & Clase & ", "
                End If
            Next j
            
            Resultado = Resultado & vbCrLf & "|-" & vbCrLf
            
        End If
        
    Next i
    
    Resultado = Resultado + "|}"
    
    frmmain.txtTxtResultado.Text = Resultado
    
    Set Leer = Nothing
    
End Sub

Public Sub ConvertArmaduras()
'***********************************************
'Autor: Lorwik
'Fecha: 09/05/2020
'Descripcion: Lee las armaduras de Obj.dat y los pasa a una tabla de Wiki
'***********************************************
    Dim TotalObj As Long
    Dim i As Long
    Dim j As Byte
    Dim TipoObjeto As Integer
    Dim Clase As String
    Dim Resultado As String
    
    Dim Leer As New clsIniManager
    
    'Limpiamos el txt
    frmmain.txtTxtResultado.Text = ""

    Call Leer.Initialize(ObjDatDir)
    
    TotalObj = Val(Leer.GetValue("INIT", "NumOBJs"))
    
    'Preparamos la cabecera de la tabla
    Resultado = "{| class='wikitable'" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    Resultado = Resultado + "!Imagen !! Nombre !! Min !! Max !! Nivel !! Precio !! Clases no Permitidas" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    
    'Reccoremos todos los objetos en busca de cascos
    For i = 1 To TotalObj
        
        TipoObjeto = Val(Leer.GetValue("OBJ" & i, "ObjType"))
        
        '¿Es una armadura?
        If TipoObjeto = TypeArmadura Then
            
            Resultado = Resultado + "|[[Archivo:" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png|thumb]] || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
            Resultado = Resultado + " || " & Val(Leer.GetValue("OBJ" & i, "Nivel")) & " || " & Val(Leer.GetValue("OBJ" & i, "Valor")) & "||"
            
             For j = 1 To NUMCLASES
                Clase = Leer.GetValue("OBJ" & i, "CP" & j)
                If Not Clase = "" Then
                    Resultado = Resultado & Clase & ", "
                End If
            Next j
            
            Resultado = Resultado & vbCrLf & "|-" & vbCrLf
            
        End If
        
    Next i
    
    Resultado = Resultado + "|}"
    
    frmmain.txtTxtResultado.Text = Resultado
    
    Set Leer = Nothing
    
End Sub

Public Sub ConvertEscudos()
'***********************************************
'Autor: Lorwik
'Fecha: 09/05/2020
'Descripcion: Lee las escudos de Obj.dat y los pasa a una tabla de Wiki
'***********************************************
    Dim TotalObj As Long
    Dim i As Long
    Dim j As Byte
    Dim TipoObjeto As Integer
    Dim Clase As String
    Dim Resultado As String
    
    Dim Leer As New clsIniManager
    
    'Limpiamos el txt
    frmmain.txtTxtResultado.Text = ""

    Call Leer.Initialize(ObjDatDir)
    
    TotalObj = Val(Leer.GetValue("INIT", "NumOBJs"))
    
    'Preparamos la cabecera de la tabla
    Resultado = "{| class='wikitable'" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    Resultado = Resultado + "!Imagen !! Nombre !! Min !! Max !! Nivel !! Precio !! Clases no Permitidas" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    
    'Reccoremos todos los objetos en busca de cascos
    For i = 1 To TotalObj
        
        TipoObjeto = Val(Leer.GetValue("OBJ" & i, "ObjType"))
        
        '¿Es un escudo?
        If TipoObjeto = TypeEscudos Then
            
            Resultado = Resultado + "|[[Archivo:" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png|thumb]] || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
            Resultado = Resultado + " || " & Val(Leer.GetValue("OBJ" & i, "Nivel")) & " || " & Val(Leer.GetValue("OBJ" & i, "Valor")) & "||"
            
             For j = 1 To NUMCLASES
                Clase = Leer.GetValue("OBJ" & i, "CP" & j)
                If Not Clase = "" Then
                    Resultado = Resultado & Clase & ", "
                End If
            Next j
            
            Resultado = Resultado & vbCrLf & "|-" & vbCrLf
            
        End If
        
    Next i
    
    Resultado = Resultado + "|}"
    
    frmmain.txtTxtResultado.Text = Resultado
    
    Set Leer = Nothing
    
End Sub

Public Sub ConvertArmas()
'***********************************************
'Autor: Lorwik
'Fecha: 09/05/2020
'Descripcion: Lee las armas de Obj.dat y los pasa a una tabla de Wiki
'***********************************************
    Dim TotalObj As Long
    Dim i As Long
    Dim j As Byte
    Dim TipoObjeto As Integer
    Dim Clase As String
    Dim Resultado As String
    
    Dim Leer As New clsIniManager
    
    'Limpiamos el txt
    frmmain.txtTxtResultado.Text = ""

    Call Leer.Initialize(ObjDatDir)
    
    TotalObj = Val(Leer.GetValue("INIT", "NumOBJs"))
    
    'Preparamos la cabecera de la tabla
    Resultado = "{| class='wikitable'" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    Resultado = Resultado + "!Imagen !! Nombre !! Min !! Max !! Nivel !! Precio !! Clases no Permitidas" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    
    'Reccoremos todos los objetos en busca de cascos
    For i = 1 To TotalObj
        
        TipoObjeto = Val(Leer.GetValue("OBJ" & i, "ObjType"))
        
        '¿Es un arma?
        If TipoObjeto = TypeArmas Then
            
            Resultado = Resultado + "|[[Archivo:" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png|thumb]] || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinHit")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxHit"))
            Resultado = Resultado + " || " & Val(Leer.GetValue("OBJ" & i, "Nivel")) & " || " & Val(Leer.GetValue("OBJ" & i, "Valor")) & "||"
            
             For j = 1 To NUMCLASES
                Clase = Leer.GetValue("OBJ" & i, "CP" & j)
                If Not Clase = "" Then
                    Resultado = Resultado & Clase & ", "
                End If
            Next j
            
            Resultado = Resultado & vbCrLf & "|-" & vbCrLf
            
        End If
        
    Next i
    
    Resultado = Resultado + "|}"
    
    frmmain.txtTxtResultado.Text = Resultado
    
    Set Leer = Nothing
    
End Sub

Public Sub ConvertHechizos()
'***********************************************
'Autor: Lorwik
'Fecha: 13/08/2022
'Descripcion: Lee los Hechizos de Obj.dat y los pasa a una tabla de Wiki
'***********************************************
End Sub

Public Sub ConvertNPCHostiles()
'***********************************************
'Autor: Lorwik
'Fecha: 13/08/2022
'Descripcion: Lee los Hechizos de Obj.dat y los pasa a una tabla de Wiki
'***********************************************

    Dim TotalNPC As Long
    Dim i As Long
    Dim j As Byte
    Dim TipoNPC As Integer
    Dim Clase As String
    Dim Resultado As String
    
    Dim Leer As New clsIniManager
    
    'Limpiamos el txt
    frmmain.txtTxtResultado.Text = ""

    Call Leer.Initialize(NPCDatDir)
    
    TotalNPC = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    'Preparamos la cabecera de la tabla
    Resultado = "{| class='wikitable'" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    Resultado = Resultado + "!Imagen !! Nombre !! Vida !! Exp !! Oro !! Drop !! Daño !! Defensa !! Evasión !! Nivel" & vbCrLf
    Resultado = Resultado + "|-" & vbCrLf
    
    'Reccoremos todos los objetos en busca de cascos
    For i = 1 To TotalNPC
        
        TipoNPC = Val(Leer.GetValue("NPC" & i, "Hostile"))
        
        '¿Es hostil?
        If TipoNPC = 1 Then
            
            Resultado = Resultado + "|[[Archivo:" & i & ".png|thumb]] || " & Leer.GetValue("NPC" & i, "Name") & " || " & Val(Leer.GetValue("NPC" & i, "MinHP")) & " / " & Val(Leer.GetValue("NPC" & i, "MaxHP")) & " || " & Val(Leer.GetValue("NPC" & i, "GiveEXP"))
            Resultado = Resultado + " || " & Val(Leer.GetValue("NPC" & i, "GiveGLD")) & " || " & Val(Leer.GetValue("NPC" & i, "MaxHIT")) & " / " & Val(Leer.GetValue("NPC" & i, "MinHIT")) & "||" & Val(Leer.GetValue("NPC" & i, "PoderEvasion")) & "||" & Val(Leer.GetValue("NPC" & i, "ELV")) & "||"
            
            Resultado = Resultado & vbCrLf & "|-" & vbCrLf
            
        End If
        
    Next i
    
    Resultado = Resultado + "|}"
    
    frmmain.txtTxtResultado.Text = Resultado
    
    Set Leer = Nothing
    
End Sub
