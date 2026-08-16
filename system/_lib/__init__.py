"""Shared helpers for the standalone scripts in this repo.

Started as helpers for the tray/utility scripts under ``system/`` and stayed
here when ``console_output`` was promoted in ``#108``: native Windows console
tools are a system-level concern, and this is the repo's only established
shared-helper package. Scripts in other domain folders reach it by putting
``system/`` on ``sys.path`` - see ``console_output``'s module docstring.
"""
