import pytest
from visiowings.vba_export import VisioVBAExporter
from visiowings.vba_import import VisioVBAImporter

class TestHeaderProcessing:
    def test_strip_vba_header_export(self):
        exporter = VisioVBAExporter("dummy.vsdm")

        # Test basic header stripping
        code = """VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub Test()
    Debug.Print "Hello"
End Sub"""

        # Test with keep_vb_name=True (default for .bas export)
        cleaned = exporter._strip_vba_header_export(code, keep_vb_name=True)
        assert 'Attribute VB_Name = "TestClass"' in cleaned
        assert 'VERSION 1.0 CLASS' not in cleaned
        assert 'BEGIN' not in cleaned
        assert 'Attribute VB_GlobalNameSpace' not in cleaned
        assert 'Option Explicit' in cleaned
        assert 'Sub Test()' in cleaned

        # Test with keep_vb_name=False (for comparison)
        cleaned_no_name = exporter._strip_vba_header_export(code, keep_vb_name=False)
        assert 'Attribute VB_Name = "TestClass"' not in cleaned_no_name

    def test_strip_vba_header_import(self):
        importer = VisioVBAImporter("dummy.vsdm")

        code = """Attribute VB_Name = "Module1"
Option Explicit

Sub Test()
    MsgBox "Hi"
End Sub"""

        # Import logic typically strips everything except code
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert 'Attribute VB_Name' not in cleaned
        assert 'Option Explicit' in cleaned
        assert 'Sub Test()' in cleaned

    def test_nested_begin_blocks(self):
        exporter = VisioVBAExporter("dummy.vsdm")

        # Form with nested controls
        code = """VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommandButton1
      Caption         =   "CommandButton1"
   End
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    MsgBox "Clicked"
End Sub"""

        cleaned = exporter._strip_vba_header_export(code, keep_vb_name=True)
        assert 'Begin' not in cleaned
        assert 'End' not in cleaned.splitlines()[0] # The block 'End' should be gone
        assert 'Option Explicit' in cleaned
        assert 'Private Sub CommandButton1_Click()' in cleaned
        assert 'End Sub' in cleaned # This code 'End' should remain
