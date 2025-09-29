# SharePoint Manager

**SharePoint Manager** is a Python library developed by ANLA's Monitoring Center to simplify the management of SharePoint lists, items, and attachments.

## 🔧 Features

- Modular structure with both class-based and functional components.
- Easy interaction with SharePoint lists, list items, and file attachments.
- Designed to support automation of common SharePoint workflows.

## 📦 Installation

You can install the library directly from GitHub:

```bash
pip install git+https://github.com/centromonitoreo/sharepoint-manager.git
```

Or install it locally from the project folder:

```bash
pip install .
```

## 📁 Project Structure

```
src/
└── sharepoint_manager/
    ├── sharepoint_class_management/
    └── sharepoint_functions/
```

## ✅ Requirements

- Python >= 3.7
- requests
- office365-rest-python-client

Install dependencies with:

```bash
pip install -r requirements.txt
```

## 🚀 Example Usage

```python
from sharepoint_manager.sharepoint_class_management import sharepoint_management

sp = sharepoint_management.SharePointManager(site_url="...", credentials={...})
sp.get_list_items("MyListName")
```

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
