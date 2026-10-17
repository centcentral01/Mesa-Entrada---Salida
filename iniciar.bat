@echo off
REM Abre la app Mesa de Entradas en el navegador (puerto 8510).
cd /d "%~dp0"
"C:\Proyectos\PythonEnvs\viaticos_casa\Scripts\python.exe" -m streamlit run app.py --server.port 8510
