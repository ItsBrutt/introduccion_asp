'Operadores
<%@ Language="VBScript" %>
<% Option Explicit %>
<% 
 Dim a, b, resultado
 a = 10
 b = 5
 
 Response.Write("a = " & a & "<br>")
 Response.Write("b = " & b & "<br>")
 
 'Aritméticos
 Response.Write("a + b = " & (a + b) & "<br>")
 Response.Write("a - b = " & (a - b) & "<br>")
 Response.Write("a * b = " & (a * b) & "<br>")
 Response.Write("a / b = " & (a / b) & "<br>")
 Response.Write("a mod b = " & (a Mod b) & "<br>") ' Módulo
 Response.Write("a ^ b = " & (a ^ b) & "<br>") ' Potencia
 
 'Comparación
 Response.Write("a = b: " & (a = b) & "<br>")
 Response.Write("a <> b: " & (a <> b) & "<br>")
 Response.Write("a > b: " & (a > b) & "<br>")
 Response.Write("a < b: " & (a < b) & "<br>")
 Reponse.Write("a <= b: " & (a <= b) & "<br>")
 Reponse.Write("a >= b: " & (a >= b) & "<br>")
 
 'Lógicos
 resultado = (a > 5) And (b < 10)
 Response.Write("(a > 5) And (b < 10): " & resultado & "<br>")

 resultado = (a > 5) Or (b < 10)
 Response.Write("(a > 5) Or (b < 10): " & resultado & "<br>")

 resultado = Not (a = 10)
 Response.Write("Not (a = 10): " & resultado & "<br>")
 


 'Concatenación
 Dim nombre, apellido, nombreCompleto
 nombre = "Juan"
 apellido = "Perez"
 nombreCompleto = nombre & " " & apellido
 Response.Write("Nombre completo: " & nombreCompleto & "<br>")

 'Operador Is
 Dim obj1, obj2
 Set obj1 = Server
 Set obj2 = Server
 
 Response.Write("obj1 Is obj2: " & (obj1 Is obj2) & "<br>") ' True, ambos apuntan al mismo objeto

 Set obj2 = Nothing
 Response.Write("obj1 Is obj2 después de Nothing: " & (obj1 Is obj2) & "<br>") ' False

 'Operador Eqv y Imp (no tan comunes pero válidos)
 Response.Write("a Eqv b: " & (a Eqv b) & "<br>") ' Equivalencia lógica
 Response.Write("a Imp b: " & (a Imp b) & "<br>") ' Implicación lógica

 %>