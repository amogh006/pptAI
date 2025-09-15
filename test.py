#!/usr/bin/env python3
"""
Presentation-only WebSocket test with audio playback
"""

import asyncio
import websockets
import json
import time
import argparse
import base64
import tempfile
import os
from pathlib import Path
from collections import defaultdict

# Audio playback
try:
    import pygame
    PYGAME_AVAILABLE = True
except ImportError:
    PYGAME_AVAILABLE = False
    print("Warning: pygame not available. Install with: pip install pygame")

def play_audio_chunk(audio_data_b64):
    """Play a base64 encoded audio chunk"""
    if not PYGAME_AVAILABLE:
        print("Audio playback not available (pygame not installed)")
        return False
    
    try:
        # Initialize pygame mixer if needed
        if not pygame.mixer.get_init():
            pygame.mixer.init(frequency=22050, size=-16, channels=2, buffer=1024)
        
        # Decode base64 audio
        audio_bytes = base64.b64decode(audio_data_b64)
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(suffix='.mp3', delete=False) as temp_file:
            temp_file.write(audio_bytes)
            temp_filename = temp_file.name
        
        # Play audio
        pygame.mixer.music.load(temp_filename)
        pygame.mixer.music.play()
        
        # Wait for playback to complete (blocking)
        while pygame.mixer.music.get_busy():
            time.sleep(0.1)
        
        # Clean up
        os.unlink(temp_filename)
        return True
        
    except Exception as e:
        print(f"Audio playback error: {e}")
        return False

class AudioChunkCollector:
    """Collects and reconstructs split audio chunks"""
    
    def __init__(self):
        self.chunks = defaultdict(dict)  # chunk_id -> {part_num: data}
        self.completed_chunks = {}
    
    def add_chunk(self, chunk_id, audio_data, part=None, total_parts=None):
        """Add an audio chunk or chunk part"""
        if part is not None and total_parts is not None:
            # This is a split chunk
            base_id = chunk_id.split('-')[0]  # Remove part suffix
            self.chunks[base_id][part] = audio_data
            
            # Check if we have all parts
            if len(self.chunks[base_id]) == total_parts:
                # Reconstruct complete chunk
                complete_data = ""
                for i in range(1, total_parts + 1):
                    complete_data += self.chunks[base_id][i]
                
                self.completed_chunks[base_id] = complete_data
                del self.chunks[base_id]  # Clean up
                return base_id, complete_data
        else:
            # Single chunk
            self.completed_chunks[chunk_id] = audio_data
            return chunk_id, audio_data
        
        return None, None

async def test_presentation(presentation_file, uri="ws://localhost:8000/ws/presentation"):
    """Test presentation loading and slide presentation with audio playback"""
    
    if not Path(presentation_file).exists():
        print(f"Error: Presentation file '{presentation_file}' not found")
        return
    
    try:
        print(f"Loading presentation: {presentation_file}")
        with open(presentation_file, 'r') as f:
            presentation_data = json.load(f)
        print(f"Presentation has {len(presentation_data.get('slides', []))} slides")
    except Exception as e:
        print(f"Error reading presentation file: {e}")
        return
    
    audio_collector = AudioChunkCollector()
    
    try:
        print(f"Connecting to {uri}...")
        websocket = await websockets.connect(uri)
        
        # Wait for connection confirmation
        response = await asyncio.wait_for(websocket.recv(), timeout=5.0)
        data = json.loads(response)
        print(f"Connected: {data.get('message', '')}")
        
        # Configure TTS
        print("Configuring TTS...")
        await websocket.send(json.dumps({
            "type": "configure_tts",
            "voice": "alloy",
            "model": "tts-1",
            "speed": 1.0
        }))
        response = await asyncio.wait_for(websocket.recv(), timeout=5.0)
        print("TTS configured")
        
        # Load presentation
        print("Loading presentation...")
        await websocket.send(json.dumps({
            "type": "load_presentation",
            "data": presentation_data
        }))
        response = await asyncio.wait_for(websocket.recv(), timeout=10.0)
        data = json.loads(response)
        if data.get("type") == "presentation_loaded":
            print(f"Presentation loaded: {data.get('total_slides')} slides")
        else:
            print(f"Failed to load presentation: {data}")
            return
        
        # Test slide 1
        print("\nStarting slide 1...")
        await websocket.send(json.dumps({
            "type": "slide_start",
            "slide_number": 1
        }))
        
        # Listen for responses
        slide_started = False
        audio_chunks_received = 0
        completed_audio_chunks = 0
        start_time = time.time()
        
        while time.time() - start_time < 60:  # Wait up to 60 seconds
            try:
                response = await asyncio.wait_for(websocket.recv(), timeout=5.0)
                data = json.loads(response)
                msg_type = data.get("type")
                
                if msg_type == "slide_started":
                    slide_started = True
                    print("Slide started")
                
                elif msg_type == "audio_chunk":
                    audio_chunks_received += 1
                    chunk_id = data.get("chunk_id", "?")
                    is_final = data.get("is_final", False)
                    audio_data = data.get("audio_data", "")
                    part = data.get("part")
                    total_parts = data.get("total_parts")
                    
                    if part and total_parts:
                        print(f"Audio chunk {chunk_id} part {part}/{total_parts}")
                    else:
                        print(f"Audio chunk {chunk_id}")
                    
                    # Collect chunk
                    complete_id, complete_audio = audio_collector.add_chunk(
                        chunk_id, audio_data, part, total_parts
                    )
                    
                    # If we have a complete chunk, play it
                    if complete_audio:
                        completed_audio_chunks += 1
                        print(f"Playing complete audio chunk {complete_id}")
                        
                        if PYGAME_AVAILABLE:
                            try:
                                # Decode audio
                                audio_bytes = base64.b64decode(complete_audio)
                                
                                # Create temporary file
                                with tempfile.NamedTemporaryFile(suffix='.mp3', delete=False) as temp_file:
                                    temp_file.write(audio_bytes)
                                    temp_filename = temp_file.name
                                
                                # Initialize pygame if needed
                                if not pygame.mixer.get_init():
                                    pygame.mixer.init()
                                
                                # Stop any currently playing audio
                                pygame.mixer.music.stop()
                                
                                # Load and play new audio
                                pygame.mixer.music.load(temp_filename)
                                pygame.mixer.music.play()
                                
                                print(f"Started playing chunk {complete_id}")
                                
                                # Wait for this chunk to finish playing before continuing
                                playback_start = time.time()
                                while pygame.mixer.music.get_busy():
                                    await asyncio.sleep(0.1)
                                    # Safety timeout to prevent infinite loop
                                    if time.time() - playback_start > 30:
                                        print(f"Playback timeout for chunk {complete_id}")
                                        pygame.mixer.music.stop()
                                        break
                                
                                print(f"Finished playing chunk {complete_id}")
                                
                                # Clean up
                                os.unlink(temp_filename)
                                
                            except Exception as e:
                                print(f"Error playing audio: {e}")
                        else:
                            print(f"Would play audio chunk {complete_id} ({len(complete_audio)} chars)")
                    
                    if is_final:
                        print("Received final audio chunk")
                        # Wait a bit more for any remaining playback
                        if PYGAME_AVAILABLE and pygame.mixer.music.get_busy():
                            print("Waiting for final audio to complete...")
                            while pygame.mixer.music.get_busy():
                                await asyncio.sleep(0.1)
                        break
                
                elif msg_type == "slide_done":
                    print("Slide completed")
                    break
                
                elif msg_type == "error":
                    error_msg = data.get("message", "Unknown error")
                    print(f"Server error: {error_msg}")
                    break
                
                else:
                    print(f"{msg_type}: {data.get('message', '')}")
            
            except asyncio.TimeoutError:
                print("Waiting for more messages...")
                continue
            except Exception as e:
                print(f"Error receiving message: {e}")
                break
        
        print(f"\nSummary:")
        print(f"- Slide started: {'Yes' if slide_started else 'No'}")
        print(f"- Audio chunk parts received: {audio_chunks_received}")
        print(f"- Complete audio chunks played: {completed_audio_chunks}")
        print(f"- Audio playback available: {'Yes' if PYGAME_AVAILABLE else 'No'}")
        
        await websocket.close()
        print("\nTest completed")
        
    except Exception as e:
        print(f"Connection error: {e}")

def main():
    parser = argparse.ArgumentParser(description="Presentation WebSocket Test with Audio Playback")
    parser.add_argument("--presentation", "-p", required=True,
                       help="Path to presentation JSON file")
    parser.add_argument("--uri", default="ws://localhost:8000/ws/presentation",
                       help="WebSocket URI")
    
    args = parser.parse_args()
    
    print("Presentation WebSocket Test with Audio Playback")
    print(f"File: {args.presentation}")
    print(f"Server: {args.uri}")
    print(f"Audio support: {'Yes' if PYGAME_AVAILABLE else 'No (install pygame)'}")
    print("-" * 50)
    
    asyncio.run(test_presentation(args.presentation, args.uri))

if __name__ == "__main__":
    main()