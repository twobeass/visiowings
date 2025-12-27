"""Helper functions for Visio COM connection"""

from pathlib import Path

import win32com.client


def get_visio_app():
    """Holt die laufende Visio-Instanz oder startet eine neue"""
    try:
        return win32com.client.Dispatch("Visio.Application")
    except Exception as e:
        print(f"❌ Fehler: Visio konnte nicht gestartet werden: {e}")
        return None


def find_open_document(visio_app, file_path):
    """Sucht ein geöffnetes Dokument anhand des Pfads"""
    file_path = Path(file_path).resolve()

    for doc in visio_app.Documents:
        if Path(doc.FullName).resolve() == file_path:
            return doc

    return None


def list_open_documents(visio_app):
    """Listet alle geöffneten Visio-Dokumente auf"""
    documents = []
    for doc in visio_app.Documents:
        documents.append({
            'name': doc.Name,
            'path': doc.FullName,
            'has_vba': doc.VBProject.VBComponents.Count > 0
        })
    return documents
