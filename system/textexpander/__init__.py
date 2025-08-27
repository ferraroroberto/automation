"""
Text Expander Package

A Windows-native text expansion utility that monitors global keyboard input
and replaces user-defined abbreviations with full text snippets.
"""

__version__ = "1.0.0"
__author__ = "Automation Project"

from .text_expander_core import TextExpanderCore, TextExpanderConfig, TextExpanderSettings
from .text_expander_gui import TextExpanderGUI
from .text_expander_monitor import KeyboardMonitor
from .text_expander_tray import TextExpanderTray
from .main import TextExpanderApp

__all__ = [
    'TextExpanderCore',
    'TextExpanderConfig',
    'TextExpanderSettings',
    'TextExpanderGUI',
    'KeyboardMonitor',
    'TextExpanderTray',
    'TextExpanderApp'
]
