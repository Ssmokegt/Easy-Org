# Easy-Org: Gestor de Organigramas Empresariales

Easy-Org es una aplicación de escritorio en Python para crear, editar y exportar organigramas empresariales de manera visual y sencilla. Este programa es gratuito y de código abierto, apto para profesionales en RRHH y administradores de empresas que deseen una herramienta rápida y sencilla para gestionar sus organigramas.

## Características principales
- **Gestión de empleados:** Añade, edita, elimina y reordena empleados con facilidad.
- **Arrastrar y soltar:** Cambia la jerarquía de supervisores con drag & drop en la lista.
- **Personalización visual:** Cambia colores, fuentes y estilos de los cuadros.
- **Exportación:** Organigrama a PNG, PDF y lista de empleados a CSV.
- **Importación/Exportación masiva:** Importa empleados desde Excel usando una plantilla.
- **Persistencia automática:** Guarda y carga los datos automáticamente en JSON.
- **Notificaciones visuales:** Mensajes de éxito/error integrados en la interfaz.

## Instalación
1. **Clona el repositorio:**
   ```sh
   git clone https://github.com/tu-usuario/Easy-Org.git
   cd Easy-Org
   ```
2. **Instala las dependencias:**
   ```sh
   pip install -r requirements.txt
   ```

## Uso
1. Ejecuta la aplicación:
   ```sh
   python main.py
   ```
2. Agrega empleados manualmente o usa el botón "Importar empleados desde Excel".
3. Personaliza el organigrama y expórtalo a PNG o PDF.
4. Guarda y carga tus datos automáticamente.

## Formato de plantilla Excel
El archivo debe tener estos encabezados:
```
Nombre | Puesto | Departamento | Supervisor | Color
```
Puedes generar la plantilla desde la app con el botón "Exportar plantilla Excel".

## Captura de pantalla
![alt text](image.png)
## Licencia
Este proyecto está bajo la licencia MIT. Consulta el archivo LICENSE para más información.

---

**¡Contribuciones y sugerencias son bienvenidas!**
