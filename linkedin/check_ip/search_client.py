#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Search Client Module

This module handles Google Lens searches via SerpAPI with API history tracking
for the LinkedIn image search application.

Author: Roberto (Refactored by Claude)
Date: March 2025
"""

import datetime
import json
import logging
import pandas as pd
import requests
import urllib.parse
from typing import Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class SearchClient:
    """Client for performing Google Lens searches via SerpAPI."""
    
    def __init__(self, api_key: str, search_endpoint: str):
        """
        Initialize search client.
        
        Args:
            api_key (str): SerpAPI key
            search_endpoint (str): SerpAPI Google Lens endpoint
        """
        self.api_key = api_key
        self.search_endpoint = search_endpoint
    
    def search_google_lens(self, image_url: str, local_image: str, 
                          api_history_df: pd.DataFrame, search_type: str = "all") -> Tuple[Optional[dict], pd.DataFrame]:
        """
        Search Google Lens via SerpAPI with the specified search type.
        
        Args:
            image_url (str): URL of the image to search
            local_image (str): Filename of the local image
            api_history_df (pd.DataFrame): API history database
            search_type (str): Type of search - "all", "exact_matches", "visual_matches", or "products"
            
        Returns:
            Tuple[Optional[dict], pd.DataFrame]: Search data and updated API history
        """
        params = {
            "engine": "google_lens",
            "url": image_url,
            "type": search_type,
            "api_key": self.api_key,
            "no_cache": True  # Forces fresh results
        }
        
        try:
            response = requests.get(self.search_endpoint, params=params)
            
            logger.info(f"🔍 Response Status for {search_type}: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                
                # Add to API history
                api_history_df = self._add_to_api_history(
                    api_history_df,
                    f'{search_type.replace("_", " ").title()} Search',
                    local_image,
                    image_url,
                    data,
                    search_type_param=search_type
                )
                
                return data, api_history_df
            else:
                logger.error(f"❌ SerpAPI Request Failed: {response.status_code}")
                return None, api_history_df
        except Exception as e:
            logger.error(f"❌ Exception during Google Lens search: {str(e)}")
            return None, api_history_df
    
    def _add_to_api_history(self, history_df: pd.DataFrame, search_type: str, 
                           local_image: str, imgur_url: str, response_data: dict,
                           page_token: Optional[str] = None, 
                           search_type_param: Optional[str] = None) -> pd.DataFrame:
        """
        Add an API call record to the API history database.
        
        Args:
            history_df (pd.DataFrame): API history database
            search_type (str): Type of search (human-readable description)
            local_image (str): Filename of the local image
            imgur_url (str): Imgur URL of the uploaded image
            response_data (dict): API response data
            page_token (Optional[str]): Page token for legacy searches (deprecated)
            search_type_param (Optional[str]): The actual 'type' parameter used in the API call
            
        Returns:
            pd.DataFrame: Updated API history database
        """
        try:
            # Extract metadata from response
            search_metadata = response_data.get('search_metadata', {})
            search_parameters = response_data.get('search_parameters', {})
            
            # Create a new record
            new_record = {
                'search_type': search_type,
                'local_image': local_image,
                'imgur_url': imgur_url,
                'search_date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                
                # API metadata
                'api_id': search_metadata.get('id'),
                'api_status': search_metadata.get('status'),
                'json_endpoint': search_metadata.get('json_endpoint'),
                'created_at': search_metadata.get('created_at'),
                'processed_at': search_metadata.get('processed_at'),
                'total_time_taken': search_metadata.get('total_time_taken'),
                'search_engine_query': search_metadata.get('google_lens_url'),
                'raw_html_file': search_metadata.get('raw_html_file'),
                
                # Search parameters
                'engine': search_parameters.get('engine'),
                'url': search_parameters.get('url'),
                'page_token': page_token,  # Keep for backward compatibility, will be None for new searches
                
                # Additional SerpAPI-specific fields
                'user': 'roberto.ferraro@gmail.com',  # User email is not in the API response, but we know it
                'requester_ip': None,  # Not available in API response
                'rtt': 1,  # Not available in API response, default to 1
                'device': 'desktop',  # Not available in API response, default to desktop
                'source_user_agent': 'python-requests/2.28.2',  # Not available in API response
                'used_user_agent': None,  # Not available in API response
                'proxy_provider': None,  # Not available in API response
                'playground_link': None,  # We'll construct this below
                
                # Store the complete response for future reference
                'raw_json': json.dumps(response_data)
            }
            
            # Add the search type parameter to the record if provided
            if search_type_param:
                new_record['search_type_param'] = search_type_param
            
            # Construct playground link based on search type
            url_param = urllib.parse.quote(imgur_url)
            if search_type_param:
                new_record['playground_link'] = f"https://serpapi.com/playground?engine=google_lens&url={url_param}&type={search_type_param}"
            else:
                new_record['playground_link'] = f"https://serpapi.com/playground?engine=google_lens&url={url_param}"
            
            # Add to history database
            return pd.concat([history_df, pd.DataFrame([new_record])], ignore_index=True)
        except Exception as e:
            logger.error(f"❌ Error adding to API history: {str(e)}")
            return history_df
    
    def validate_api_key(self) -> bool:
        """
        Validate SerpAPI key by making a test request.
        
        Returns:
            bool: True if API key is valid, False otherwise
        """
        try:
            # Make a simple test request to validate the API key
            test_params = {
                "engine": "google",
                "q": "test",
                "api_key": self.api_key
            }
            
            response = requests.get("https://serpapi.com/search", params=test_params)
            
            if response.status_code == 200:
                logger.info("✅ SerpAPI key validated successfully")
                return True
            else:
                logger.error(f"❌ SerpAPI key validation failed: {response.status_code}")
                return False
        
        except Exception as e:
            logger.error(f"❌ Error validating SerpAPI key: {str(e)}")
            return False 