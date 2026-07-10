# Soporte de Técnicos

Aplicación Flask para gestionar asignaciones de casos, técnicos y reporte en PDF.

## Requisitos
- Python 3.10+
- Base de datos MySQL/MariaDB con las tablas:
  - soporte_departamentos
  - soporte_grupos
  - soporte_tecnicos
  - soporte_casos_asignados

## Instalación
1. Crear un entorno virtual:
   ```bash
   python -m venv venv
   .\\venv\\Scripts\\activate
   ```
2. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Configurar la conexión a la base de datos usando usuario y contraseña reales. La base de datos debe llamarse `soporte`.
   - Opción 1: exportar variables de entorno antes de iniciar la app:
     ```powershell
     $env:DB_USER = 'tu_usuario'
     $env:DB_PASSWORD = 'tu_contraseña'
     $env:DB_HOST = 'localhost'
     $env:DB_NAME = 'soporte'
     ```
   - Opción 2: cambiar directamente en `app.py` las variables `DB_USER`, `DB_PASSWORD`, `DB_HOST` y `DB_NAME`.
4. Ejecutar la app:
   ```bash
   python app.py
   ```

## Uso
- `Registro`: asigna un caso a un técnico y guarda el registro en `soporte_casos_asignados`.
- `Base de datos`: filtra y exporta los datos a PDF.
- `Moderación`: revisa y elimina asignaciones.
- `Técnicos`: agrega o elimina técnicos.
