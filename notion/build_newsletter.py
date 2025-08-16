#!/usr/bin/env python3
"""
Newsletter Builder for Notion

This script fetches articles from a Notion database based on a newsletter number,
groups them by topic, and outputs HTML lists ready for Substack.

Usage:
    python build_newsletter.py --newsletter "057" --config "build_newsletter.json"
    python build_newsletter.py --newsletter "001" --debug
"""

import argparse
import json
import logging
import os
import sys
from typing import Dict, List, Optional, Tuple, Any
import requests
from dotenv import load_dotenv

# Load environment variables from root .env file
load_dotenv()


class NotionNewsletterBuilder:
    """Main class for building newsletters from Notion data."""
    
    def __init__(self, config_path: str):
        """Initialize the builder with configuration."""
        self.config = self._load_config(config_path)
        self.notion_api_key = self.config.get('notion_api_key')
        self.articles_db_id = self.config.get('articles_database_id')
        self.newsletter_db_id = self.config.get('newsletter_database_id')
        
        # Validate required config values
        if not all([self.notion_api_key, self.articles_db_id, self.newsletter_db_id]):
            raise ValueError("Missing required configuration values. Check your JSON config file.")
        
        # Set up Notion API headers
        self.headers = {
            'Authorization': f'Bearer {self.notion_api_key}',
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        }
        
        # Define the three topic categories
        self.topics = [
            'Management & Leadership',
            'Personal Development', 
            'Innovation'
        ]
    
    def _load_config(self, config_path: str) -> Dict:
        """Load and parse the JSON configuration file."""
        # First try the provided path
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
                return config
        except FileNotFoundError:
            # If not found, try looking in the same directory as this script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            fallback_path = os.path.join(script_dir, os.path.basename(config_path))
            try:
                with open(fallback_path, 'r') as f:
                    config = json.load(f)
                    logging.info(f"Loaded config from fallback path: {fallback_path}")
                    return config
            except FileNotFoundError:
                raise FileNotFoundError(f"Configuration file not found at {config_path} or {fallback_path}")
            except json.JSONDecodeError:
                raise ValueError(f"Invalid JSON in fallback configuration file: {fallback_path}")
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON in configuration file: {config_path}")
        
        # Process environment variables in config
        config = self._process_environment_variables(config)
        return config
    
    def _process_environment_variables(self, config: Dict) -> Dict:
        """Process environment variables in configuration values.
        
        Supports ${ENV_VAR_NAME} syntax for environment variable substitution.
        """
        def replace_env_vars(obj: Any) -> Any:
            """Recursively replace environment variables in configuration values."""
            if isinstance(obj, dict):
                return {k: replace_env_vars(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [replace_env_vars(item) for item in obj]
            elif isinstance(obj, str) and obj.startswith('${') and obj.endswith('}'):
                env_var = obj[2:-1]
                value = os.getenv(env_var)
                if value is None:
                    logging.error(f"Environment variable not found: {env_var}")
                    sys.exit(1)
                return value
            return obj
        
        return replace_env_vars(config)
    
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
                logging.error(f"Failed to query Notion database: {e}")
                raise
        
        return all_results
    
    def find_newsletter_by_title(self, newsletter_title: str) -> Optional[Dict]:
        """Find a newsletter record by its title."""
        logging.info(f"Searching for newsletter with title: {newsletter_title}")
        
        # Filter to find newsletter with matching title
        filter_data = {
            "property": "Title",
            "title": {
                "equals": newsletter_title
            }
        }
        
        results = self._query_notion_database(self.newsletter_db_id, filter_data)
        
        if not results:
            logging.error(f"No newsletter found with title: {newsletter_title}")
            return None
        
        newsletter = results[0]
        logging.info(f"Found newsletter: {newsletter.get('id')}")
        return newsletter
    
    def get_related_articles(self, newsletter_id: str) -> List[Dict]:
        """Get all articles related to a specific newsletter."""
        logging.info(f"Fetching articles related to newsletter: {newsletter_id}")
        
        # Filter articles by relation to the newsletter
        filter_data = {
            "property": "Newsletter",
            "relation": {
                "contains": newsletter_id
            }
        }
        
        results = self._query_notion_database(self.articles_db_id, filter_data)
        logging.info(f"Found {len(results)} related articles")
        return results
    
    def extract_article_data(self, article: Dict) -> Optional[Tuple[str, str, str]]:
        """Extract name, URL, and topic from an article record."""
        try:
            # Extract article name from title property
            title_prop = article.get('properties', {}).get('Name', {})
            if title_prop.get('type') == 'title':
                title_content = title_prop.get('title', [])
                if title_content:
                    name = title_content[0].get('plain_text', '').strip()
                else:
                    logging.warning(f"Article {article.get('id')} has empty title")
                    return None
            else:
                logging.warning(f"Article {article.get('id')} has invalid title property")
                return None
            
            # Extract URL from url property
            url_prop = article.get('properties', {}).get('URL', {})
            if url_prop.get('type') == 'url':
                url = url_prop.get('url', '').strip()
                if not url:
                    logging.warning(f"Article '{name}' has empty URL")
                    return None
            else:
                logging.warning(f"Article '{name}' has invalid URL property")
                return None
            
            # Extract topic from select property
            topic_prop = article.get('properties', {}).get('Topic', {})
            if topic_prop.get('type') == 'select':
                topic_obj = topic_prop.get('select')
                if topic_obj:
                    topic = topic_obj.get('name', '').strip()
                else:
                    logging.warning(f"Article '{name}' has no topic selected")
                    return None
            else:
                logging.warning(f"Article '{name}' has invalid topic property")
                return None
            
            return name, url, topic
            
        except Exception as e:
            logging.error(f"Error extracting data from article {article.get('id')}: {e}")
            return None
    
    def group_articles_by_topic(self, articles: List[Dict]) -> Dict[str, List[Tuple[str, str]]]:
        """Group articles by topic, returning only the three specified topics."""
        grouped = {topic: [] for topic in self.topics}
        
        for article in articles:
            article_data = self.extract_article_data(article)
            if article_data:
                name, url, topic = article_data
                
                if topic in self.topics:
                    grouped[topic].append((name, url))
                    logging.debug(f"Added article '{name}' to topic '{topic}'")
                else:
                    logging.warning(f"Article '{name}' has unknown topic: {topic} (skipping)")
        
        # Log grouping results
        for topic in self.topics:
            count = len(grouped[topic])
            logging.info(f"Topic '{topic}': {count} articles")
        
        return grouped
    
    def generate_html_lists(self, grouped_articles: Dict[str, List[Tuple[str, str]]]) -> str:
        """Generate HTML output with three topic sections."""
        output = []
        
        for topic in self.topics:
            articles = grouped_articles[topic]
            
            # Add topic header
            output.append(topic)
            
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
    
    def build_newsletter(self, newsletter_title: str) -> str:
        """Main method to build a complete newsletter."""
        logging.info(f"Building newsletter: {newsletter_title}")
        
        # Step 1: Find the newsletter record
        newsletter = self.find_newsletter_by_title(newsletter_title)
        if not newsletter:
            raise ValueError(f"Newsletter '{newsletter_title}' not found")
        
        # Step 2: Get related articles
        articles = self.get_related_articles(newsletter['id'])
        if not articles:
            raise ValueError(f"No articles found for newsletter '{newsletter_title}'")
        
        # Step 3: Group articles by topic
        grouped_articles = self.group_articles_by_topic(articles)
        
        # Step 4: Generate HTML output
        html_output = self.generate_html_lists(grouped_articles)
        
        logging.info("Newsletter build completed successfully")
        return html_output


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
  python build_newsletter.py --config "custom_config.json"
        """
    )
    
    parser.add_argument(
        '--newsletter',
        type=str,
        default="001",
        help='Newsletter number/title to fetch (default: "001")'
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
        # Initialize the newsletter builder
        builder = NotionNewsletterBuilder(args.config)
        
        # Build the newsletter
        html_output = builder.build_newsletter(args.newsletter)
        
        # Output the result
        print(html_output)
        
    except (ValueError, FileNotFoundError, requests.exceptions.RequestException) as e:
        logging.error(f"Failed to build newsletter: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == "__main__":
    main()