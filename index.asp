<%@ Language="VBScript" %>
<%
    ' Lógica de servidor básica
    Dim userName, serverTime, serverName
    userName = "Usuario" ' Podríamos obtenerlo de una sesión
    serverTime = Now()
    serverName = Request.ServerVariables("SERVER_NAME")
    
    ' Array de colores para demostrar lógica en el servidor
    Dim techStack(3)
    techStack(0) = "VBScript"
    techStack(1) = "IIS"
    techStack(2) = "Classic ASP"
    techStack(3) = "Modern CSS"
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mi Primera Aplicación ASP | Inicio</title>
    <meta name="description" content="Bienvenido a tu primera aplicación desarrollada con Classic ASP y diseño moderno.">
    
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&family=Outfit:wght@400;700&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --primary: #8b5cf6;
            --primary-dark: #6d28d9;
            --bg: #0f172a;
            --card-bg: rgba(30, 41, 59, 0.7);
            --text: #f8fafc;
            --text-muted: #94a3b8;
            --accent: #0ea5e9;
            --glass-border: rgba(255, 255, 255, 0.1);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg);
            background-image: 
                radial-gradient(circle at 0% 0%, rgba(139, 92, 246, 0.15) 0%, transparent 35%),
                radial-gradient(circle at 100% 100%, rgba(14, 165, 233, 0.15) 0%, transparent 35%);
            color: var(--text);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            overflow-x: hidden;
        }

        .container {
            max-width: 800px;
            width: 90%;
            padding: 2rem;
            position: relative;
        }

        /* Hero Section */
        .hero {
            text-align: center;
            margin-bottom: 3rem;
            animation: fadeInDown 0.8s ease-out;
        }

        h1 {
            font-family: 'Outfit', sans-serif;
            font-size: clamp(2.5rem, 8vw, 4rem);
            font-weight: 800;
            line-height: 1.1;
            margin-bottom: 1rem;
            background: linear-gradient(to right, #fff, var(--primary), var(--accent));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .subtitle {
            font-size: 1.25rem;
            color: var(--text-muted);
            margin-bottom: 2rem;
        }

        /* Glass Card */
        .card {
            background: var(--card-bg);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid var(--glass-border);
            border-radius: 24px;
            padding: 2.5rem;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            animation: fadeInUp 1s ease-out 0.2s both;
        }

        .server-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1.5rem;
            margin-top: 2rem;
        }

        .info-box {
            padding: 1.25rem;
            background: rgba(255, 255, 255, 0.03);
            border-radius: 16px;
            border: 1px solid var(--glass-border);
            transition: transform 0.3s ease, border-color 0.3s ease;
        }

        .info-box:hover {
            transform: translateY(-5px);
            border-color: var(--primary);
        }

        .info-label {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: var(--text-muted);
            margin-bottom: 0.5rem;
        }

        .info-value {
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--accent);
        }

        .stack-list {
            display: flex;
            flex-wrap: wrap;
            gap: 0.75rem;
            margin-top: 1.5rem;
            list-style: none;
        }

        .stack-item {
            padding: 0.5rem 1rem;
            background: rgba(139, 92, 246, 0.1);
            color: var(--primary);
            border: 1px solid rgba(139, 92, 246, 0.2);
            border-radius: 99px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        /* Animations */
        @keyframes fadeInDown {
            from { opacity: 0; transform: translateY(-30px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(30px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .footer {
            margin-top: 3rem;
            text-align: center;
            color: var(--text-muted);
            font-size: 0.9rem;
        }

        .status-dot {
            display: inline-block;
            width: 10px;
            height: 10px;
            background-color: #10b981;
            border-radius: 50%;
            margin-right: 8px;
            box-shadow: 0 0 10px #10b981;
            animation: pulse 2s infinite;
        }

        @keyframes pulse {
            0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
            70% { transform: scale(1); box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }
            100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
        }
    </style>
</head>
<body>

    <div class="container">
        <header class="hero">
            <h1>¡Hola, ASP!</h1>
            <p class="subtitle">Has iniciado con éxito tu primera aplicación en el servidor.</p>
        </header>

        <main class="card">
            <div style="display: flex; align-items: center; margin-bottom: 1.5rem;">
                <span class="status-dot"></span>
                <h2 style="font-size: 1.5rem; font-weight: 600;">Servidor Activo</h2>
            </div>
            
            <p>Este contenido está siendo procesado por el motor de <strong>Classic ASP</strong> en Windows IIS. A continuación verás datos generados dinámicamente en el servidor:</p>

            <div class="server-info">
                <div class="info-box">
                    <p class="info-label">Fecha y Hora del Servidor</p>
                    <p class="info-value"><%= serverTime %></p>
                </div>
                <div class="info-box">
                    <p class="info-label">Nombre del Host</p>
                    <p class="info-value"><%= serverName %></p>
                </div>
            </div>

            <h3 style="margin-top: 2.5rem; font-size: 1.1rem; color: var(--text-muted);">Tecnologías en uso:</h3>
            <ul class="stack-list">
                <% For i = 0 To UBound(techStack) %>
                    <li class="stack-item"><%= techStack(i) %></li>
                <% Next %>
            </ul>
        </main>

        <footer class="footer">
            <p>&copy; <%= Year(Now()) %> - Tu Primer Proyecto ASP</p>
        </footer>
    </div>

</body>
</html>
