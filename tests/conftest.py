import sys
from unittest.mock import MagicMock

# Mock win32com.client and pythoncom before any visiowings modules are imported
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["pythoncom"] = MagicMock()
