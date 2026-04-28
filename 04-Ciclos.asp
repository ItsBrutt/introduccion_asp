<%@ Language="VBScript" %>
<% Option Explicit %>
<% 
 'Estructuras repetitivas
 Dim i, persona, personas(2)
 
 'Ciclo For
 Response.Write("Ciclo For:<br>")
 For i = 1 To 10
    Response.Write(i & "<br>")
 Next
 
 'Ciclo For Each
 Response.Write("Ciclo For Each:<br>")
 personas(0) = "Juan"
 personas(1) = "Maria"
 personas(2) = "Pedro"
 
 For Each persona In personas
    Response.Write(persona & "<br>")
 Next
 
 'Ciclo Do While
 Response.Write("Ciclo Do While:<br>")
 i = 1
 Do While i <= 5
    Response.Write(i & "<br>")
    i = i + 1
 Loop
 
 'Ciclo Do Until
 Response.Write("Ciclo Do Until:<br>")
 i = 1
 Do Until i > 5
    Response.Write(i & "<br>")
    i = i + 1
 Loop
 
 'Ciclo Do
 Response.Write("Ciclo Do:<br>")
 i = 1
 Do
    Response.Write(i & "<br>")
    i = i + 1
 Loop While i <= 5
 
 'Exit For
 Response.Write("Ciclo For con Exit For:<br>")
 For i = 1 To 5
    If i = 3 Then
        Exit For
    End If
    Response.Write(i & "<br>")
 Next
 
 'Exit Do
 Response.Write("Ciclo Do con Exit Do:<br>")
 i = 1
 Do
    If i = 3 Then
        Exit Do
    End If
    Response.Write(i & "<br>")
    i = i + 1
 Loop
 
 'Exit For Each
 Response.Write("Ciclo For Each con Exit For Each:<br>")
 For Each persona In personas
    If persona = "Maria" Then
        Exit For
    End If
    Response.Write(persona & "<br>")
 Next
 '
 'Continue For (no existe en VBScript, se simula con If)
 Response.Write("Ciclo For con Continue For (simulado):<br>")
 For i = 1 To 5
    If i <> 3 Then
        Response.Write(i & "<br>")
    End If
 Next
 
 'Continue Do (no existe en VBScript, se simula con If)
 Response.Write("Ciclo Do con Continue Do (simulado):<br>")
 i = 1
 Do While i <= 5
    If i <> 3 Then
        Response.Write(i & "<br>")
    End If
    i = i + 1
 Loop
 
 'Continue For Each (no existe en VBScript, se simula con If)
 Response.Write("Ciclo For Each con Continue For Each (simulado):<br>")
 For Each persona In personas
    If persona <> "Maria" Then
        Response.Write(persona & "<br>")
    End If
 Next
 
 %>