#!/usr/bin/env python3
"""
Script to save images from FastAPI PowerPoint conversion response
Supports both JSON response and direct API calls
"""

import json
import base64
import os
import requests
from pathlib import Path
import argparse
from typing import Dict, Any, Optional
from datetime import datetime

class ImageSaver:
    def __init__(self, output_dir: str = "slides_output"):
        """
        Initialize ImageSaver
        
        Args:
            output_dir: Directory to save images (default: slides_output)
        """
        self.output_dir = Path(output_dir)
        self.ensure_output_dir()
    
    def ensure_output_dir(self):
        """Create output directory if it doesn't exist"""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Output directory: {self.output_dir.absolute()}")
    
    def save_from_json_response(self, response_data: Dict[str, Any], filename_prefix: str = "slide") -> None:
        """
        Save images from JSON API response
        
        Args:
            response_data: JSON response from the API
            filename_prefix: Prefix for saved image files
        """
        if response_data.get("status") != "success":
            raise ValueError(f"API response indicates failure: {response_data}")
        
        slides = response_data.get("slides", [])
        total_slides = response_data.get("total_slides", len(slides))
        
        print(f"Saving {total_slides} slides...")
        
        for slide_data in slides:
            slide_number = slide_data.get("slide_number", 1)
            image_data = slide_data.get("image_data")
            content_type = slide_data.get("content_type", "image/png")
            
            if not image_data:
                print(f"Warning: No image data for slide {slide_number}")
                continue
            
            # Determine file extension from content type
            extension = self._get_extension_from_content_type(content_type)
            
            # Create filename
            filename = f"{filename_prefix}_{slide_number:03d}.{extension}"
            filepath = self.output_dir / filename
            
            # Decode and save image
            try:
                image_bytes = base64.b64decode(image_data)
                with open(filepath, 'wb') as f:
                    f.write(image_bytes)
                print(f"✓ Saved: {filename} ({len(image_bytes)} bytes)")
            except Exception as e:
                print(f"✗ Failed to save slide {slide_number}: {e}")
        
        print(f"\nAll slides saved to: {self.output_dir.absolute()}")
    
    def save_from_json_file(self, json_file_path: str, filename_prefix: str = "slide") -> None:
        """
        Load JSON response from file and save images
        
        Args:
            json_file_path: Path to JSON file containing API response
            filename_prefix: Prefix for saved image files
        """
        try:
            with open(json_file_path, 'r') as f:
                response_data = json.load(f)
            self.save_from_json_response(response_data, filename_prefix)
        except FileNotFoundError:
            print(f"Error: JSON file not found: {json_file_path}")
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON in file {json_file_path}: {e}")
    
    def save_from_api_call(self, api_url: str, ppt_file_path: str, filename_prefix: str = "slide") -> None:
        """
        Make API call and save images directly
        
        Args:
            api_url: URL of the FastAPI endpoint
            ppt_file_path: Path to PowerPoint file to convert
            filename_prefix: Prefix for saved image files
        """
        if not os.path.exists(ppt_file_path):
            raise FileNotFoundError(f"PowerPoint file not found: {ppt_file_path}")
        
        print(f"Uploading {ppt_file_path} to {api_url}...")
        
        try:
            with open(ppt_file_path, 'rb') as f:
                files = {'file': (os.path.basename(ppt_file_path), f, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')}
                response = requests.post(api_url, files=files, timeout=120)
            
            response.raise_for_status()
            response_data = response.json()
            
            print(f"✓ API call successful")
            self.save_from_json_response(response_data, filename_prefix)
            
        except requests.exceptions.RequestException as e:
            print(f"✗ API call failed: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response status: {e.response.status_code}")
                print(f"Response text: {e.response.text}")
    
    def _get_extension_from_content_type(self, content_type: str) -> str:
        """Get file extension from content type"""
        content_type_map = {
            "image/png": "png",
            "image/jpeg": "jpg",
            "image/jpg": "jpg",
            "image/gif": "gif",
            "image/bmp": "bmp",
            "image/webp": "webp"
        }
        return content_type_map.get(content_type.lower(), "png")

def main():
    parser = argparse.ArgumentParser(description="Save images from PowerPoint conversion API response")
    parser.add_argument(
        '--mode', 
        choices=['json-file', 'json-string', 'api-call'], 
        required=True,
        help="Mode of operation"
    )
    parser.add_argument(
        '--input', 
        required=True,
        help="Input: JSON file path, JSON string, or PowerPoint file path (depending on mode)"
    )
    parser.add_argument(
        '--output-dir', 
        default=f"slides_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        help="Output directory for images"
    )
    parser.add_argument(
        '--prefix', 
        default="slide",
        help="Filename prefix for saved images"
    )
    parser.add_argument(
        '--api-url', 
        default="http://localhost:8000/convert-ppt/",
        help="API URL for api-call mode"
    )
    
    args = parser.parse_args()
    
    # Create ImageSaver instance
    saver = ImageSaver(args.output_dir)
    
    try:
        if args.mode == 'json-file':
            saver.save_from_json_file(args.input, args.prefix)
        
        elif args.mode == 'json-string':
            response_data = json.loads(args.input)
            saver.save_from_json_response(response_data, args.prefix)
        
        elif args.mode == 'api-call':
            saver.save_from_api_call(args.api_url, args.input, args.prefix)
        
        print(f"\n🎉 Success! Images saved to: {saver.output_dir.absolute()}")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        return 1
    
    return 0

# Example usage functions for interactive use
def save_from_response(response_json: dict, output_dir: str = None) -> None:
    """
    Convenience function to save images from API response
    
    Args:
        response_json: JSON response from the API
        output_dir: Output directory (optional)
    """
    if output_dir is None:
        output_dir = f"slides_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    saver = ImageSaver(output_dir)
    saver.save_from_json_response(response_json)

def save_from_file(json_file_path: str, output_dir: str = None) -> None:
    """
    Convenience function to save images from JSON file
    
    Args:
        json_file_path: Path to JSON file
        output_dir: Output directory (optional)
    """
    if output_dir is None:
        output_dir = f"slides_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    saver = ImageSaver(output_dir)
    saver.save_from_json_file(json_file_path)

def convert_and_save(ppt_file_path: str, api_url: str = "http://localhost:8000/convert-ppt/", output_dir: str = None) -> None:
    """
    Convenience function to convert PowerPoint and save images in one step
    
    Args:
        ppt_file_path: Path to PowerPoint file
        api_url: API endpoint URL
        output_dir: Output directory (optional)
    """
    if output_dir is None:
        output_dir = f"slides_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    saver = ImageSaver(output_dir)
    saver.save_from_api_call(api_url, ppt_file_path)

if __name__ == "__main__":
    exit(main())