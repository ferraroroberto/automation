#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Imgur Client Module

This module handles image upload operations to Imgur with retry logic and exponential backoff
for the LinkedIn image search application.

Author: Roberto
Date: March 2025
"""

import logging
from pathlib import Path
import requests
import time
from typing import Optional

logger = logging.getLogger(__name__)


class ImgurClient:
    """Client for uploading images to Imgur with retry logic."""
    
    def __init__(self, client_id: str, access_token: str):
        """
        Initialize Imgur client.
        
        Args:
            client_id (str): Imgur Client ID
            access_token (str): Imgur API access token
        """
        self.client_id = client_id
        self.access_token = access_token
        self.api_url = "https://api.imgur.com/3/image"
        self.headers = {"Authorization": f"Bearer {access_token}"}
    
    def upload_image(self, image_path: Path, retries: int = 20) -> Optional[str]:
        """
        Upload an image to Imgur with retry logic and exponential backoff.
        
        Args:
            image_path (Path): Path to the image file
            retries (int): Number of retry attempts
            
        Returns:
            Optional[str]: Imgur URL if successful, None otherwise
        """
        delay = 5  # Initial retry delay in seconds
        
        for attempt in range(1, retries + 1):
            try:
                logger.info(f"🚀 Attempt {attempt}/{retries}: Uploading {image_path.name}")
                
                with open(image_path, "rb") as file:
                    response = requests.post(
                        self.api_url,
                        headers=self.headers,
                        files={"image": file}
                    )
                
                if response.status_code == 200:
                    img_url = response.json()["data"]["link"]
                    logger.info(f"✅ Image uploaded successfully: {img_url}")
                    return img_url
                
                elif response.status_code == 503:
                    logger.warning(f"⚠️ Server busy (503 Error). Attempt {attempt}/{retries}. Retrying in {delay} seconds...")
                
                elif response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", delay))
                    logger.warning(f"⚠️ Rate limit exceeded (429 Error). Attempt {attempt}/{retries}. Waiting {retry_after} seconds before retrying...")
                    delay = retry_after  # Use API-provided delay instead of exponential
                
                else:
                    logger.error(f"❌ Error {response.status_code}: {response.text}. Upload failed.")
                    return None
            
            except requests.exceptions.RequestException as e:
                logger.error(f"❌ Network error: {str(e)}. Attempt {attempt}/{retries}. Retrying in {delay} seconds...")
            
            # Apply exponential backoff, doubling the delay each attempt
            time.sleep(delay)
            delay = min(delay * 2, 1800)  # Max delay capped at 30 minutes
        
        logger.error(f"❌ Upload failed after {retries} attempts. Skipping {image_path.name}.")
        return None
    
    def validate_credentials(self) -> bool:
        """
        Validate Imgur credentials by attempting a test request.
        
        Returns:
            bool: True if credentials are valid, False otherwise
        """
        try:
            # Make a simple request to validate credentials
            response = requests.get(
                "https://api.imgur.com/3/account/me",
                headers=self.headers
            )
            
            if response.status_code == 200:
                logger.info("✅ Imgur credentials validated successfully")
                return True
            else:
                logger.error(f"❌ Imgur credential validation failed: {response.status_code}")
                return False
        
        except Exception as e:
            logger.error(f"❌ Error validating Imgur credentials: {str(e)}")
            return False 