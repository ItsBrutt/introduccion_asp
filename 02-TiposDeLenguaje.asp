<%@ Language="VBScript"%>
<% Option Explicit %>
<%
' ============================================
' TÍTULO: 02 - Tipos de Lenguajes
' ============================================

' Definir variables
Dim nombre, edad, precio, fecha, activo

' Asignación de tipos
nombre = "Arnaldo"     ' String
edad = 25              ' Integer
precio = 10.99         ' Double
fecha = #07/18/1987#   ' Date (se rodea con #)
activo = true          ' Boolean
%>

<%
' ============================================
' Imprimir variables
' ============================================
%>
<h1>Bienvenido <% =nombre %></h1>
<br>
<h2>Edad: <% =edad %></h2>
<br>
<h3>Precio: <% =precio %></h3>
<br>
<h4>Fecha: <% =fecha %></h4>
<br>
<h5>Activo: <% =activo %></h5>
