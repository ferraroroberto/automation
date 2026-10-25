"""The one logger every articles-sync module writes to.

`NotionArticlesSync.setup_logging` attaches the console and file handlers to
this logger, so the modules split out of `notion_articles_sync.py` must log
through it, not through a logger named after their own module.
"""

import logging

logger = logging.getLogger("notion_articles_sync")
