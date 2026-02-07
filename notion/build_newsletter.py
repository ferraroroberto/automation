#!/usr/bin/env python3
"""
Newsletter Builder for Notion

This script fetches articles from a Notion database based on a newsletter number,
groups them by topic, and outputs HTML lists ready for Substack.

Usage:
    python build_newsletter.py --newsletter "057"
    python build_newsletter.py --newsletter "001" --debug
"""

import argparse
import re
import json
import logging
import os
import sys
import webbrowser
from typing import Dict, List, Optional, Tuple
import requests
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


class NotionNewsletterBuilder:
    """Main class for building newsletters from Notion data."""
    
    def __init__(self, config_path: str):
        """Initialize the builder with configuration."""
        self.config = self._load_config(config_path)
        
        # Get config values
        self.notion_api_key = os.getenv('NOTION_API_TOKEN') or self.config.get('notion_api_key')
        self.articles_db_id = self.config.get('articles_database_id')
        self.newsletter_db_id = self.config.get('newsletter_database_id')
        
        # Validate required config values
        if not all([self.notion_api_key, self.articles_db_id, self.newsletter_db_id]):
            raise ValueError("Missing required configuration values")
        
        # Set up Notion API headers
        self.headers = {
            'Authorization': f'Bearer {self.notion_api_key}',
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        }
        
        # Define the three topic categories
        self.topics = [
            'personal development',
            'innovation',
            'leadership and management'
        ]
        
        logging.info("✅ Newsletter builder initialized")
        logging.info(f"📊 Articles database: {self.articles_db_id}")
        logging.info(f"📊 Newsletter database: {self.newsletter_db_id}")
    
    def _load_config(self, config_path: str) -> Dict:
        """Load and parse the JSON configuration file."""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            # Try looking in the same directory as this script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            fallback_path = os.path.join(script_dir, os.path.basename(config_path))
            try:
                with open(fallback_path, 'r') as f:
                    logging.info(f"📁 Loaded config from fallback path: {fallback_path}")
                    return json.load(f)
            except FileNotFoundError:
                raise FileNotFoundError(f"Configuration file not found at {config_path} or {fallback_path}")
            except json.JSONDecodeError:
                raise ValueError(f"Invalid JSON in fallback configuration file: {fallback_path}")
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON in configuration file: {config_path}")
    
    def _query_notion_database(self, database_id: str, filter_data: Optional[Dict] = None) -> List[Dict]:
        """Query a Notion database with optional filtering."""
        url = f"https://api.notion.com/v1/databases/{database_id}/query"
        
        payload = {}
        if filter_data:
            payload['filter'] = filter_data
        
        all_results = []
        has_more = True
        start_cursor = None
        
        while has_more:
            if start_cursor:
                payload['start_cursor'] = start_cursor
            
            try:
                response = requests.post(url, headers=self.headers, json=payload)
                response.raise_for_status()
                
                data = response.json()
                all_results.extend(data.get('results', []))
                
                has_more = data.get('has_more', False)
                start_cursor = data.get('next_cursor')
                
            except requests.exceptions.RequestException as e:
                logging.error(f"❌ Failed to query Notion database: {e}")
                raise
        
        return all_results
    
    def find_newsletter_by_title(self, newsletter_title: str) -> Optional[Dict]:
        """Find a newsletter record by its number."""
        logging.info(f"🔍 Searching for newsletter: {newsletter_title}")
        
        filter_data = {
            "property": "number",
            "title": {
                "equals": newsletter_title
            }
        }
        
        results = self._query_notion_database(self.newsletter_db_id, filter_data)
        
        if not results:
            logging.error(f"❌ No newsletter found with number: {newsletter_title}")
            return None
        
        logging.info(f"✅ Found newsletter: {newsletter_title}")
        return results[0]
    
    def get_related_articles(self, newsletter_id: str) -> List[Dict]:
        """Get all articles related to a specific newsletter."""
        logging.info(f"📥 Fetching articles for newsletter: {newsletter_id}")
        
        filter_data = {
            "property": "news",
            "relation": {
                "contains": newsletter_id
            }
        }
        
        results = self._query_notion_database(self.articles_db_id, filter_data)
        logging.info(f"📊 Found {len(results)} articles")
        return results
    
    def extract_article_data(self, article: Dict) -> Optional[Tuple[str, str, str, bool, List[str]]]:
        """Extract name, URL, topic, star, and niche from an article record."""
        try:
            properties = article.get('properties', {})
            
            # Extract title
            title_prop = properties.get('article', {})
            if title_prop.get('type') == 'title':
                title_content = title_prop.get('title', [])
                if not title_content:
                    return None
                name = title_content[0].get('plain_text', '').strip()
            else:
                return None
            
            # Extract URL
            url_prop = properties.get('link', {})
            if url_prop.get('type') == 'url':
                url = url_prop.get('url', '').strip()
                if not url:
                    return None
            else:
                return None
            
            # Extract topic
            topic_prop = properties.get('topic', {})
            if topic_prop.get('type') == 'select':
                topic_obj = topic_prop.get('select')
                if not topic_obj:
                    return None
                topic = topic_obj.get('name', '').strip()
            else:
                return None
            
            # Extract star (checkbox)
            star_prop = properties.get('star', {})
            star = False
            if star_prop.get('type') == 'checkbox':
                star = star_prop.get('checkbox', False)
            
            # Extract niche (multi_select)
            niche_prop = properties.get('niche', {})
            niche = []
            if niche_prop.get('type') == 'multi_select':
                niche_objs = niche_prop.get('multi_select', [])
                niche = [obj.get('name', '').strip() for obj in niche_objs if obj.get('name')]
            
            return name, url, topic, star, niche
            
        except Exception as e:
            logging.error(f"❌ Error extracting data from article: {e}")
            return None
    
    def group_articles_by_topic(self, articles: List[Dict]) -> Dict[str, List[Tuple[str, str]]]:
        """Group articles by topic and sort them according to specified criteria."""
        grouped = {topic: [] for topic in self.topics}
        
        for article in articles:
            article_data = self.extract_article_data(article)
            if article_data:
                name, url, topic, star, niche = article_data
                if topic in self.topics:
                    grouped[topic].append((name, url, star, niche))
        
        # Sort articles within each topic group
        for topic in self.topics:
            # Sort by: 1) star (descending), 2) niche (ascending), 3) article title (ascending)
            grouped[topic].sort(key=lambda x: (
                -x[2],  # star (descending, so negative for reverse sort)
                sorted(x[3])[0] if x[3] else '',  # niche (ascending, first niche value)
                x[0].lower()  # article title (ascending, case-insensitive)
            ))
            
            # Log the sorted order for this topic
            if grouped[topic]:
                logging.info(f"📋 Topic '{topic}' sorted order:")
                for i, (name, url, star, niche) in enumerate(grouped[topic], 1):
                    niche_str = ', '.join(sorted(niche)) if niche else 'none'
                    star_str = '⭐' if star else '⚪'
                    logging.info(f"  {i}. {star_str} {name} (niche: {niche_str})")
            
            # Remove star and niche from the final output, keeping only name and url
            grouped[topic] = [(name, url) for name, url, star, niche in grouped[topic]]
        
        # Log grouping results
        for topic in self.topics:
            count = len(grouped[topic])
            logging.info(f"📋 Topic '{topic}': {count} articles")
        
        return grouped
    
    def generate_html_lists(self, grouped_articles: Dict[str, List[Tuple[str, str]]]) -> str:
        """Generate HTML output with three topic sections."""
        output = []
        
        for topic in self.topics:
            articles = grouped_articles[topic]
            
            # Add topic header with capitalized first letter
            capitalized_topic = topic[0].upper() + topic[1:]
            output.append(f'<h2>{capitalized_topic}</h2>')
            
            # Generate HTML list
            if articles:
                output.append('<ul>')
                for name, url in articles:
                    output.append(f'  <li><a href="{url}">{name}</a></li>')
                output.append('</ul>')
            else:
                output.append('<ul></ul>')
            
            # Add spacing between sections
            output.append('')
        
        return '\n'.join(output)
    
    def generate_complete_html(self, grouped_articles: Dict[str, List[Tuple[str, str]]]) -> str:
        """Generate complete HTML document with proper structure."""
        html_content = self.generate_html_lists(grouped_articles)

        html_document = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Newsletter Content</title>
    <style>
        body {{
            background-color: black;
            color: white;
            font-family: Arial, sans-serif;
            margin: 20px;
        }}
        h2 {{
            color: #ffffff;
            border-bottom: 1px solid #333;
            padding-bottom: 10px;
        }}
        a {{
            color: #4a9eff;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}
        ul {{
            margin: 0;
            padding-left: 20px;
        }}
        li {{
            margin: 8px 0;
        }}
    </style>
</head>
<body>
{html_content}
</body>
</html>"""

        return html_document
    
    def build_newsletter(self, newsletter_title: str) -> Tuple[str, Dict[str, List[Tuple[str, str]]]]:
        """Main method to build a complete newsletter."""
        logging.info(f"🚀 Building newsletter: {newsletter_title}")
        
        # Find the newsletter record
        newsletter = self.find_newsletter_by_title(newsletter_title)
        if not newsletter:
            raise ValueError(f"Newsletter '{newsletter_title}' not found")
        
        # Get related articles
        articles = self.get_related_articles(newsletter['id'])
        if not articles:
            raise ValueError(f"No articles found for newsletter '{newsletter_title}'")
        
        # Group articles by topic and generate HTML
        grouped_articles = self.group_articles_by_topic(articles)
        html_output = self.generate_html_lists(grouped_articles)
        
        logging.info("✅ Newsletter build completed successfully")
        return html_output, grouped_articles


def setup_logging(debug: bool = False):
    """Set up logging configuration."""
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stdout)]
    )


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Build newsletters from Notion data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python build_newsletter.py --newsletter "057"
  python build_newsletter.py --newsletter "001" --debug
        """
    )
    
    parser.add_argument(
        '--newsletter',
        type=str,
        help='Newsletter number (Nxxx). If omitted, you will be prompted.'
    )
    
    parser.add_argument(
        '--config',
        type=str,
        default="build_newsletter.json",
        help='Path to JSON config file (default: "build_newsletter.json")'
    )
    
    parser.add_argument(
        '--debug',
        action='store_true',
        help='Enable debug logging'
    )
    
    args = parser.parse_args()
    
    # Set up logging
    setup_logging(args.debug)
    
    try:
        # Determine newsletter number: CLI arg or prompt via input()
        newsletter_number = args.newsletter
        if not newsletter_number:
            try:
                newsletter_number = input("Enter newsletter number (Nxxx): ")
            except (EOFError, KeyboardInterrupt):
                logging.error("❌ Newsletter number input cancelled")
                sys.exit(2)

        if not newsletter_number:
            logging.error("❌ Newsletter number is required and must be in the format Nxxx (e.g., N057)")
            sys.exit(2)

        newsletter_number = newsletter_number.strip().upper()

        if not re.fullmatch(r'N\d{3}', newsletter_number):
            logging.error("❌ Newsletter number must be in the format Nxxx (e.g., N057)")
            sys.exit(2)

        # Initialize the newsletter builder
        builder = NotionNewsletterBuilder(args.config)
        
        # Build the newsletter
        html_output, grouped_articles = builder.build_newsletter(newsletter_number)
        
        # Generate complete HTML document
        complete_html = builder.generate_complete_html(grouped_articles)
        
        # Save to HTML file in the same directory as this script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        html_filename = os.path.join(script_dir, "build_newsletter.html")
        with open(html_filename, 'w', encoding='utf-8') as f:
            f.write(complete_html)
        
        logging.info(f"💾 HTML file saved to: {html_filename}")
        
        # Open the HTML file with default browser
        webbrowser.open(f'file://{html_filename}')
        logging.info(f"🌐 Opened HTML file with default browser: {html_filename}")
        
    except (ValueError, FileNotFoundError, requests.exceptions.RequestException) as e:
        logging.error(f"❌ Failed to build newsletter: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"❌ Unexpected error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == "__main__":
    main()