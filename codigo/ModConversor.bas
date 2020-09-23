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
    ObjDatDir = "D:\Proyectos\WinterAO\Server-WinterAO-Resurrection\Dat\OBJ.dat"
    NPCDatDir = "D:\Proyectos\WinterAO\Server-WinterAO-Resurrection\Dat\NPCs.dat"
    HechizosDatDir = "D:\Proyectos\WinterAO\Server-WinterAO-Resurrection\Dat\Hechizos.dat"
    
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
            
            Resultado = Resultado + "|" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
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
            
            Resultado = Resultado + "|" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
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
            
            Resultado = Resultado + "|" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinDef")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxDef"))
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
            
            Resultado = Resultado + "|" & Leer.GetValue("OBJ" & i, "GrhIndex") & ".png || " & Leer.GetValue("OBJ" & i, "Name") & " || " & Val(Leer.GetValue("OBJ" & i, "MinHit")) & " || " & Val(Leer.GetValue("OBJ" & i, "MaxHit"))
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
