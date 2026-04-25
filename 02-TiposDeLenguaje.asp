' ============================================'
' TÍTULO: 02 - Tipos de Lenguajes
'<%@ Language="JScript" %>
<%@ Language="VBScript" %>
<% option explicit %>
' ============================================'
' Definir variables
Dim nombre, edad, precio,fecha,activo
nombre = "Arnaldo" 'String
edad = 25 'Integer
precio = 10.99 'Double
fecha = #07/18/1987# 'Date
activo = true 'Boolean
%> 
' ============================================'
' Imprimir variables
<% Response.Write("<h1>Bienvenido " & nombre & "</h1>") %><br>
<% Response.Write("<h2>" & edad & "</h2>") %><br>
<% Response.Write("<h3>" & precio & "</h3>") %><br>
' ============================================'