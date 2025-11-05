# CTP Consulting Engineers - Document Submittal Response (DSR) Template

## Descripción

Este es un template profesional de Excel para **DOCUMENT SUBMITTAL RESPONSE / DESIGN SUBMITTAL REVIEW (DSR)** creado para CTP Consulting Engineers para revisar diseños estructurales de otras empresas.

## Archivos Incluidos

1. **CTP_Document_Submittal_Response_DSR_Template.xlsx** - Template principal de Excel
2. **CTP_Logo_Placeholder.png** - Logo placeholder (temporal)
3. **create_dsr_template.py** - Script para generar el template
4. **create_logo_placeholder.py** - Script para crear el logo placeholder

## Características del Template

El template incluye las siguientes secciones:

### 1. Información del Proyecto
- Nombre del proyecto
- Número de proyecto
- Cliente
- Contratista/Diseñador
- Fecha

### 2. Información del Submittal
- Número de submittal
- Revisión
- Título del documento
- Número de documento/plano
- Fechas de recepción y revisión

### 3. Estado de Revisión
- ☐ APPROVED (Aprobado)
- ☐ APPROVED AS NOTED (Aprobado con notas)
- ☐ REVISE AND RESUBMIT (Revisar y reenviar)
- ☐ REJECTED (Rechazado)

### 4. Checklist de Revisión Estructural
Lista detallada de verificación que incluye:
- Cálculos de diseño completos
- Combinaciones de carga conformes a códigos
- Especificaciones de materiales
- Planos estructurales coordinados
- Cumplimiento de códigos (IBC, ACI, AISC, etc.)
- Sistema de resistencia a cargas laterales
- Diseño de fundaciones
- Detalles de conexiones
- Límites de deflexión
- Requisitos de inspección especial
- Criterios de diseño sísmico
- Cálculos de carga de viento
- Detalles de construcción
- Dimensionamiento de elementos estructurales
- Detalles de refuerzo

### 5. Comentarios y Observaciones
Sección amplia para notas detalladas del revisor

### 6. Información del Revisor
- Nombre del revisor
- Título
- Número de licencia
- Firma
- Fecha

## Cómo Usar el Template

1. **Abrir el archivo Excel**: `CTP_Document_Submittal_Response_DSR_Template.xlsx`

2. **Completar la información del proyecto**: Llenar todos los campos de información del proyecto y submittal

3. **Marcar el estado de revisión**: Seleccionar una de las opciones de estado

4. **Completar el checklist**: Marcar el estado de cada ítem en la columna "Status":
   - ✓ = Cumple
   - ✗ = No cumple
   - N/A = No aplica

5. **Agregar comentarios**: Escribir observaciones detalladas en la sección de comentarios

6. **Firmar y fechar**: Completar la información del revisor

7. **Guardar y enviar**: Guardar el documento con un nombre descriptivo (ej: `DSR-001_ProjectName_Rev0.xlsx`)

## Reemplazar el Logo Placeholder

El template actualmente incluye un logo placeholder. Para reemplazarlo con el logo oficial de CTP:

### Método 1: Manualmente en Excel
1. Abrir el archivo Excel
2. Click derecho en el logo placeholder
3. Seleccionar "Delete" o "Eliminar"
4. Ir a Insert > Pictures > Picture from File
5. Seleccionar el logo oficial de CTP
6. Ajustar el tamaño según sea necesario
7. Guardar el archivo

### Método 2: Obtener el logo desde la web
1. Visitar https://ctp-llp.com/ (UK) o https://www.c-t-p.co.za/ (Sudáfrica)
2. Click derecho en el logo > "Save image as..."
3. Guardar como "CTP_Logo.png"
4. Usar el Método 1 para insertarlo en el Excel

### Método 3: Con Python (si tienes el archivo de logo)
```python
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

wb = load_workbook("CTP_Document_Submittal_Response_DSR_Template.xlsx")
ws = wb.active

# Agregar el logo real
img = Image("CTP_Logo_Real.png")
img.width = 300
img.height = 90
ws.add_image(img, 'B1')

wb.save("CTP_Document_Submittal_Response_DSR_Template.xlsx")
```

## Información de CTP Consulting Engineers

**Website**: https://ctp-llp.com/

**Servicios**:
- Structural Engineering
- Civil Engineering
- Heritage Engineering
- Marine Engineering

**Contacto**:
- Address: Suffolk House, 154 High Street, Sevenoaks, Kent TN13 1XE, United Kingdom
- Phone: +44 1732 740195
- Email: info@ctp-llp.com

## Notas Técnicas

- El template está diseñado siguiendo estándares internacionales de ingeniería estructural
- Compatible con Microsoft Excel 2010 o superior
- También funciona con Google Sheets y LibreOffice Calc
- El checklist está basado en códigos IBC, ACI, y AISC
- Puede ser personalizado según necesidades específicas del proyecto

## Scripts de Python

### create_dsr_template.py
Genera el template de Excel desde cero con:
- Formato profesional
- Estilos y colores corporativos
- Estructura completa de DSR

### create_logo_placeholder.py
Crea un logo placeholder temporal usando PIL/Pillow

## Licencia y Uso

Este template es para uso de CTP Consulting Engineers y sus proyectos.

---

**Creado**: Noviembre 2025
**Versión**: 1.0
