import pandas as pd
import math

def generate_html():
    try:
        df = pd.read_excel('claude_final.xlsx')
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Clean column names just in case
    df.columns = [c.strip() for c in df.columns]

    # HTML Structure
    html_content = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Academic Course Feedback Report</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400&family=Lato:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary-color: #002147; /* Deep Navy Blue */
            --secondary-color: #555555; /* Dark Gray */
            --bg-color: #FDFBF7; /* Cream/Off-White */
            --accent-color: #8D0801; /* Deep Burgundy */
            --border-color: #E0E0E0;
            --highlight-bg: #F0F4F8;
        }}

        body {{
            font-family: 'Lato', sans-serif;
            background-color: var(--bg-color);
            color: var(--secondary-color);
            margin: 0;
            padding: 40px;
            line-height: 1.6;
            font-size: 16px;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 60px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
            border-top: 5px solid var(--primary-color);
        }}

        header {{
            text-align: center;
            margin-bottom: 50px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 30px;
        }}

        h1 {{
            font-family: 'Playfair Display', serif;
            font-size: 2.8rem;
            color: var(--primary-color);
            margin: 0 0 10px 0;
            font-weight: 700;
            letter-spacing: -0.5px;
        }}

        h2 {{
            font-family: 'Playfair Display', serif;
            font-size: 1.2rem;
            color: var(--secondary-color);
            margin: 0;
            font-weight: 400;
            font-style: italic;
            opacity: 0.8;
        }}

        .table-wrapper {{
            overflow-x: auto;
            border-radius: 4px;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
            white-space: nowrap;
        }}

        th, td {{
            padding: 16px 20px;
            text-align: left;
            vertical-align: middle;
        }}

        th {{
            font-family: 'Playfair Display', serif;
            font-weight: 700;
            color: var(--primary-color);
            border-bottom: 2px solid var(--primary-color);
            text-transform: uppercase;
            letter-spacing: 1px;
            font-size: 0.8rem;
            background-color: #ffffff;
            position: sticky;
            top: 0;
            z-index: 10;
        }}

        /* Specific column alignments */
        th:not(:first-child):not(:nth-child(2)),
        td:not(:first-child):not(:nth-child(2)) {{
            text-align: center;
        }}

        tr {{
            border-bottom: 1px solid var(--border-color);
            transition: background-color 0.2s ease;
        }}

        tr:hover {{
            background-color: var(--highlight-bg);
        }}

        /* Typography tweaks for specific columns */
        td:first-child {{ /* Department */
            font-weight: 700;
            color: var(--primary-color);
            font-size: 0.85rem;
            white-space: normal;
            min-width: 200px;
        }}

        td:nth-child(2) {{ /* Course */
            font-style: italic;
            font-family: 'Playfair Display', serif;
            font-size: 1rem;
            color: #222;
            white-space: normal;
            min-width: 250px;
        }}

        .score {{
            font-weight: 700;
            color: var(--primary-color);
            font-size: 1.1rem;
        }}

        .metric {{
            color: #666;
            font-size: 0.9rem;
        }}

        .pain-point {{
            color: var(--accent-color);
            font-weight: 600;
            opacity: 0.9;
        }}

        /* Empty/NaN cell styling */
        .empty {{
            color: #ccc;
            font-weight: 300;
        }}

        .footer {{
            margin-top: 60px;
            text-align: center;
            font-size: 0.8rem;
            color: #999;
            border-top: 1px solid #eee;
            padding-top: 30px;
            font-family: 'Playfair Display', serif;
            font-style: italic;
        }}
    </style>
</head>
<body>

    <div class="container">
        <header>
            <h1>Reporte de Evaluación Académica</h1>
            <h2>Análisis Integral de Cursos y Departamentos</h2>
        </header>

        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>Departamento</th>
                        <th>Materia</th>
                        <th title="Promedio CAE">Promedio</th>
                        <th title="Total Comentarios">Total</th>
                        <th title="Comentarios Positivos">Pos (+)</th>
                        <th title="Comentarios Negativos">Neg (-)</th>
                        <th title="Comentarios Neutros">Neu (~)</th>
                        <th title="Dolor Docente">Docente</th>
                        <th title="Dolor Contenido">Contenido</th>
                        <th title="Dolor Actividades">Actividades</th>
                        <th title="Dolor Plataforma">Plataforma</th>
                        <th title="Dolor Modalidad">Modalidad</th>
                        <th title="Dolor Equipo">Equipo</th>
                        <th title="Dolor Otros">Otros</th>
                    </tr>
                </thead>
                <tbody>
    """

    for index, row in df.iterrows():
        # Helper to format values
        def fmt(val, is_float=False):
            if pd.isna(val) or val == '':
                return '<span class="empty">-</span>'
            if is_float:
                return f"{val:.2f}"
            if isinstance(val, float) and val.is_integer():
                return str(int(val))
            return str(val)

        # Department & Course
        dept = fmt(row.get('departamento', ''))
        course = fmt(row.get('materia', ''))

        # Metrics
        avg_val = row.get('promedio_cae', 0)
        avg = fmt(avg_val, is_float=True)

        total = fmt(row.get('total_comentarios', 0))
        pos = fmt(row.get('comentarios_positivos', 0))
        neg = fmt(row.get('comentarios_negativos', 0))
        neu = fmt(row.get('comentarios_neutros', 0))

        # Pain points
        d_docente = fmt(row.get('dolor_docente', ''))
        d_contenido = fmt(row.get('dolor_contenido', ''))
        d_actividades = fmt(row.get('dolor_actividades', ''))
        d_plataforma = fmt(row.get('dolor_plataforma', ''))
        d_modalidad = fmt(row.get('dolor_modalidad', ''))
        d_equipo = fmt(row.get('dolor_equipo', ''))
        d_otros = fmt(row.get('dolor_otros', ''))

        html_content += f"""
                    <tr>
                        <td>{dept}</td>
                        <td>{course}</td>
                        <td class="score">{avg}</td>
                        <td class="metric">{total}</td>
                        <td class="metric">{pos}</td>
                        <td class="metric">{neg}</td>
                        <td class="metric">{neu}</td>
                        <td class="pain-point">{d_docente}</td>
                        <td class="pain-point">{d_contenido}</td>
                        <td class="pain-point">{d_actividades}</td>
                        <td class="pain-point">{d_plataforma}</td>
                        <td class="pain-point">{d_modalidad}</td>
                        <td class="pain-point">{d_equipo}</td>
                        <td class="pain-point">{d_otros}</td>
                    </tr>
        """

    html_content += """
                </tbody>
            </table>
        </div>

        <div class="footer">
            Generado automáticamente a partir de datos inmutables. | Confidential Report
        </div>
    </div>

</body>
</html>
    """

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

    print("HTML report generated successfully: index.html")

if __name__ == "__main__":
    generate_html()
