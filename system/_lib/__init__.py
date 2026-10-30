"""Shared helpers for the standalone scripts in this repo.

Native Windows console-tool decoding (``console_output``) lives here too:
``system/`` is a system-level concern, and this is the repo's only established
shared-helper package. Scripts in other domain folders reach it by putting
``system/`` on ``sys.path`` - see ``console_output``'s module docstring.
"""
