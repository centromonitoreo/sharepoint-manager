# SharePoint Manager

**SharePoint Manager** es una librería en Python desarrollada por el Centro de Monitoreo de la ANLA para facilitar la interacción y gestión de listas y documentos en entornos de SharePoint.

## 🔧 Características

- Estructura modular basada en clases y funciones.
- Interacción sencilla con listas, ítems y archivos adjuntos de SharePoint.
- Pensado para automatizar flujos de trabajo comunes en SharePoint.

## 📦 Instalación

Puedes instalar la librería directamente desde GitHub:

```bash
pip install git+hhttps://github.com/centromonitoreo/sharepoint-manager.git 
```

O instalarla localmente desde el repositorio:

```bash
pip install .
```

## 📁 Estructura del proyecto

```
src/
└── sharepoint_manager/
    ├── sharepoint_class_management/
    └── sharepoint_functions/
```

## ✅ Requisitos

- Python >= 3.7
- requests
- office365-rest-python-client

Instala las dependencias con:

```bash
pip install -r requirements.txt
```

## 🚀 Ejemplo de uso

```python
from sharepoint_manager.sharepoint_class_management import sharepoint_management

sp = sharepoint_management.SharePointManager(site_url="...", credentials={...})
sp.get_list_items("NombreDeLista")
```

## 📄 Licencia

Este proyecto está licenciado bajo la licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.
