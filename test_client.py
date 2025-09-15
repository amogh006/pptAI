import streamlit as st
import asyncio
import websockets
import json
import base64
import threading
import time
from typing import Optional, Dict, Any
import io
import tempfile
import os
from datetime import datetime
import queue
from collections import defaultdict

# Audio playback imports - define availability flags
PYGAME_AVAILABLE = False
PYAUDIO_AVAILABLE = False
PYDUB_AVAILABLE = False

try:
    import pygame
    PYGAME_AVAILABLE = True
except ImportError:
    pass

try:
    import pyaudio
    import wave
    PYAUDIO_AVAILABLE = True
except ImportError:
    pass

try:
    from pydub import AudioSegment
    from pydub.playback import play
    import simpleaudio as sa
    PYDUB_AVAILABLE = True
except ImportError:
    pass

class AudioChunkCollector:
    """Collects and reconstructs split audio chunks"""
    
    def __init__(self):
        self.chunks = defaultdict(dict)  # chunk_id -> {part_num: data}
        self.completed_chunks = {}
        self.playback_queue = queue.Queue()
    
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
                self.playback_queue.put((base_id, complete_data))
                del self.chunks[base_id]  # Clean up
                return base_id, complete_data
        else:
            # Single chunk
            self.completed_chunks[chunk_id] = audio_data
            self.playback_queue.put((chunk_id, audio_data))
            return chunk_id, audio_data
        
        return None, None

class WebSocketClient:
    def __init__(self):
        self.websocket: Optional[websockets.WebSocketServerProtocol] = None
        self.is_connected = False
        self.message_queue = queue.Queue()
        self.audio_collector = AudioChunkCollector()
        self.connection_task = None
        self.audio_playback_active = False
        self.listen_thread = None
        
    def connect_sync(self, uri: str):
        """Synchronous connection wrapper"""
        try:
            # Run in new event loop to avoid conflicts
            def _connect():
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    result = loop.run_until_complete(self._async_connect(uri))
                    return result
                finally:
                    loop.close()
            
            return _connect()
        except Exception as e:
            st.error(f"Connection failed: {e}")
            return False
    
    async def _async_connect(self, uri: str):
        """Actual async connection logic"""
        try:
            self.websocket = await websockets.connect(uri)
            # Wait for connection confirmation
            response = await asyncio.wait_for(self.websocket.recv(), timeout=5.0)
            data = json.loads(response)
            if data.get("type") == "connected":
                self.is_connected = True
                # Start listening thread
                self.listen_thread = threading.Thread(target=self._listen_worker, daemon=True)
                self.listen_thread.start()
                return True
            return False
        except Exception as e:
            st.error(f"Connection error: {e}")
            return False
    
    def _listen_worker(self):
        """Background thread for listening to WebSocket messages"""
        def _listen():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(self._listen_messages())
            except Exception as e:
                st.error(f"Listen worker error: {e}")
            finally:
                loop.close()
        
        _listen()
    
    async def _listen_messages(self):
        """Listen for messages in background thread"""
        try:
            while self.is_connected and self.websocket:
                try:
                    message = await asyncio.wait_for(self.websocket.recv(), timeout=1.0)
                    data = json.loads(message)
                    self.message_queue.put(data)
                    
                    # Handle audio chunks
                    if data.get("type") == "audio_chunk":
                        chunk_id = data.get("chunk_id", "")
                        audio_data = data.get("audio_data", "")
                        part = data.get("part")
                        total_parts = data.get("total_parts")
                        
                        self.audio_collector.add_chunk(chunk_id, audio_data, part, total_parts)
                        
                except asyncio.TimeoutError:
                    continue
                except websockets.exceptions.ConnectionClosed:
                    self.is_connected = False
                    break
        except Exception as e:
            self.is_connected = False
    
    def disconnect_sync(self):
        """Synchronous disconnect wrapper"""
        def _disconnect():
            self.is_connected = False
            if self.websocket:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(self.websocket.close())
                finally:
                    loop.close()
            self.websocket = None
        
        _disconnect()
    
    def send_message_sync(self, message: Dict[str, Any]):
        """Synchronous message sending wrapper"""
        if not self.is_connected or not self.websocket:
            return False
        
        def _send():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                return loop.run_until_complete(self._async_send(message))
            finally:
                loop.close()
        
        try:
            return _send()
        except Exception as e:
            st.error(f"Send failed: {e}")
            return False
    
    async def _async_send(self, message: Dict[str, Any]):
        """Actual async send logic"""
        try:
            await self.websocket.send(json.dumps(message))
            return True
        except Exception as e:
            self.is_connected = False
            return False

def play_audio_pygame(audio_data_b64):
    """Play audio using pygame with proper sequencing"""
    if not PYGAME_AVAILABLE:
        return False
        
    try:
        # Initialize pygame mixer if not already done
        if not pygame.mixer.get_init():
            pygame.mixer.init(frequency=22050, size=-16, channels=2, buffer=1024)
        
        # Decode audio
        audio_bytes = base64.b64decode(audio_data_b64)
        
        # Create temporary file for audio
        with tempfile.NamedTemporaryFile(suffix='.mp3', delete=False) as temp_file:
            temp_file.write(audio_bytes)
            temp_filename = temp_file.name
        
        # Stop any currently playing audio
        pygame.mixer.music.stop()
        
        # Play audio
        pygame.mixer.music.load(temp_filename)
        pygame.mixer.music.play()
        
        # Wait for playback to finish
        start_time = time.time()
        while pygame.mixer.music.get_busy():
            time.sleep(0.1)
            # Safety timeout
            if time.time() - start_time > 30:
                pygame.mixer.music.stop()
                break
        
        # Clean up
        os.unlink(temp_filename)
        return True
        
    except Exception as e:
        st.error(f"Pygame audio playback error: {e}")
        return False

def play_audio_browser(audio_data_b64):
    """Play audio using Streamlit's audio widget"""
    try:
        audio_bytes = base64.b64decode(audio_data_b64)
        audio_io = io.BytesIO(audio_bytes)
        st.audio(audio_io, format='audio/mp3', start_time=0)
        return True
    except Exception as e:
        st.error(f"Browser audio error: {e}")
        return False

def audio_playback_worker(client, method="pygame"):
    """Background worker for sequential audio playback"""
    while True:
        try:
            if not client.audio_collector.playback_queue.empty():
                chunk_id, audio_data = client.audio_collector.playback_queue.get(timeout=1.0)
                
                if method == "pygame" and PYGAME_AVAILABLE:
                    success = play_audio_pygame(audio_data)
                    if success:
                        st.session_state.audio_status = f"Played chunk {chunk_id}"
                else:
                    # For browser playback, we'll handle it in the main thread
                    st.session_state.pending_audio = (chunk_id, audio_data)
            else:
                time.sleep(0.1)
                
        except queue.Empty:
            continue
        except Exception as e:
            st.error(f"Audio playback worker error: {e}")
            break

def main():
    st.set_page_config(
        page_title="WebSocket Presentation Client",
        page_icon="🎤",
        layout="wide"
    )
    
    st.title("🎤 WebSocket Presentation Client")
    st.markdown("Real-time AI presentation with TTS audio streaming")
    
    # Show audio library warnings in main function
    if not PYGAME_AVAILABLE and not PYDUB_AVAILABLE:
        st.warning("⚠️ No advanced audio libraries detected. Install pygame for better audio: `pip install pygame`")
    
    # Initialize session state
    if 'client' not in st.session_state:
        st.session_state.client = WebSocketClient()
    if 'presentation_data' not in st.session_state:
        st.session_state.presentation_data = None
    if 'current_slide' not in st.session_state:
        st.session_state.current_slide = 1
    if 'messages' not in st.session_state:
        st.session_state.messages = []
    if 'is_presenting' not in st.session_state:
        st.session_state.is_presenting = False
    if 'audio_status' not in st.session_state:
        st.session_state.audio_status = "Ready"
    if 'pending_audio' not in st.session_state:
        st.session_state.pending_audio = None
    if 'audio_worker_started' not in st.session_state:
        st.session_state.audio_worker_started = False
    
    client = st.session_state.client
    
    # Sidebar for connection and configuration
    with st.sidebar:
        st.header("🔗 Connection")
        
        ws_url = st.text_input(
            "WebSocket URL", 
            value="ws://localhost:8000/ws/presentation",
            help="WebSocket endpoint URL"
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Connect", disabled=client.is_connected):
                with st.spinner("Connecting..."):
                    success = client.connect_sync(ws_url)
                    if success:
                        st.success("Connected!")
                        st.rerun()  # Refresh to update UI
        
        with col2:
            if st.button("Disconnect", disabled=not client.is_connected):
                client.disconnect_sync()
                st.success("Disconnected!")
                st.rerun()  # Refresh to update UI
        
        # Connection status
        status_color = "🟢" if client.is_connected else "🔴"
        st.write(f"Status: {status_color} {'Connected' if client.is_connected else 'Disconnected'}")
        
        st.header("🎛️ TTS Configuration")
        
        voice = st.selectbox(
            "Voice",
            ["alloy", "echo", "fable", "onyx", "nova", "shimmer"],
            index=4  # Default to nova
        )
        
        model = st.selectbox(
            "Model",
            ["tts-1", "tts-1-hd"],
            index=0
        )
        
        speed = st.slider(
            "Speed",
            min_value=0.25,
            max_value=4.0,
            value=1.0,
            step=0.1
        )
        
        if st.button("Configure TTS", disabled=not client.is_connected):
            message = {
                "type": "configure_tts",
                "voice": voice,
                "model": model,
                "speed": speed
            }
            success = client.send_message_sync(message)
            if success:
                st.success("TTS configured!")
                st.rerun()
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📊 Presentation Control")
        
        # Upload presentation script
        uploaded_file = st.file_uploader(
            "Upload Presentation Script (JSON)",
            type=['json'],
            help="Upload the JSON file generated by /generate-script/ endpoint"
        )
        
        if uploaded_file is not None:
            try:
                presentation_data = json.loads(uploaded_file.read().decode('utf-8'))
                st.session_state.presentation_data = presentation_data
                
                # Load presentation
                if st.button("Load Presentation", disabled=not client.is_connected):
                    message = {
                        "type": "load_presentation",
                        "data": presentation_data
                    }
                    success = client.send_message_sync(message)
                    if success:
                        st.success(f"Loaded presentation with {len(presentation_data.get('slides', []))} slides!")
                        # Start audio worker
                        if not st.session_state.audio_worker_started:
                            audio_method = st.session_state.get('audio_method', 'pygame')
                            worker_thread = threading.Thread(
                                target=audio_playback_worker, 
                                args=(client, audio_method), 
                                daemon=True
                            )
                            worker_thread.start()
                            st.session_state.audio_worker_started = True
                        st.rerun()
                
            except json.JSONDecodeError:
                st.error("Invalid JSON file")
        
        # Presentation info
        if st.session_state.presentation_data:
            presentation_info = st.session_state.presentation_data.get('presentation_info', {})
            st.info(f"**{presentation_info.get('title', 'Untitled')}** - {presentation_info.get('total_slides', 0)} slides")
        
        # Slide controls
        st.subheader("🎬 Slide Controls")
        
        col_prev, col_slide, col_next = st.columns([1, 2, 1])
        
        with col_prev:
            if st.button("⬅️ Previous", disabled=not client.is_connected or st.session_state.current_slide <= 1):
                st.session_state.current_slide = max(1, st.session_state.current_slide - 1)
        
        with col_slide:
            if st.session_state.presentation_data:
                total_slides = len(st.session_state.presentation_data.get('slides', []))
                slide_num = st.selectbox(
                    "Current Slide",
                    range(1, total_slides + 1),
                    index=st.session_state.current_slide - 1
                )
                st.session_state.current_slide = slide_num
        
        with col_next:
            max_slides = len(st.session_state.presentation_data.get('slides', [])) if st.session_state.presentation_data else 1
            if st.button("➡️ Next", disabled=not client.is_connected or st.session_state.current_slide >= max_slides):
                st.session_state.current_slide = min(max_slides, st.session_state.current_slide + 1)
        
        # Start/Stop presentation
        col_start, col_stop = st.columns(2)
        
        with col_start:
            if st.button("▶️ Start Slide", disabled=not client.is_connected or st.session_state.is_presenting):
                message = {
                    "type": "slide_start",
                    "slide_number": st.session_state.current_slide
                }
                success = client.send_message_sync(message)
                if success:
                    st.session_state.is_presenting = True
                    st.success(f"Started slide {st.session_state.current_slide}")
                    st.rerun()
        
        with col_stop:
            if st.button("⏹️ Stop", disabled=not client.is_connected):
                message = {"type": "stop"}
                success = client.send_message_sync(message)
                if success:
                    st.session_state.is_presenting = False
                    st.success("Presentation stopped")
                    st.rerun()
        
        # Interrupt controls
        st.subheader("❓ Q&A / Interrupt")
        
        question = st.text_input(
            "Question (optional)",
            placeholder="Enter your question here...",
            help="Optional question to ask during interrupt"
        )
        
        if st.button("🛑 Interrupt", disabled=not client.is_connected or not st.session_state.is_presenting):
            message = {
                "type": "interrupt",
                "question": question
            }
            success = client.send_message_sync(message)
            if success:
                st.success("Presentation interrupted for Q&A")
                st.rerun()
    
    with col2:
        st.header("📨 Messages")
        
        # Display recent messages
        message_container = st.container()
        
        # Process new messages
        while not client.message_queue.empty():
            try:
                message = client.message_queue.get_nowait()
                timestamp = datetime.now().strftime("%H:%M:%S")
                st.session_state.messages.append({
                    "timestamp": timestamp,
                    "data": message
                })
            except queue.Empty:
                break
        
        # Display messages
        with message_container:
            for msg in st.session_state.messages[-10:]:  # Show last 10 messages
                msg_type = msg["data"].get("type", "unknown")
                timestamp = msg["timestamp"]
                
                if msg_type == "audio_chunk":
                    chunk_id = msg["data"].get("chunk_id", "?")
                    part = msg["data"].get("part")
                    total_parts = msg["data"].get("total_parts")
                    is_final = msg["data"].get("is_final", False)
                    
                    if part and total_parts:
                        st.text(f"🎵 [{timestamp}] Chunk {chunk_id} part {part}/{total_parts}")
                    else:
                        st.text(f"🎵 [{timestamp}] Audio chunk {chunk_id}" + (" (Final)" if is_final else ""))
                        
                elif msg_type == "slide_done":
                    slide_num = msg["data"].get("slide_number", "?")
                    st.text(f"✅ [{timestamp}] Slide {slide_num} completed")
                    st.session_state.is_presenting = False
                elif msg_type == "qa_response":
                    st.text(f"💬 [{timestamp}] Q&A Response")
                elif msg_type == "error":
                    error_msg = msg["data"].get("message", "Unknown error")
                    st.text(f"❌ [{timestamp}] Error: {error_msg}")
                else:
                    st.text(f"📋 [{timestamp}] {msg_type}")
        
        # Audio playback section
        st.header("🔊 Audio Status")
        
        # Available audio methods
        audio_methods = ["Browser (Streamlit)"]
        
        if PYGAME_AVAILABLE:
            audio_methods.append("pygame (Sequential)")
        if PYDUB_AVAILABLE:
            audio_methods.append("pydub + simpleaudio")
        
        audio_method = st.selectbox(
            "Playback Method",
            audio_methods,
            help="Choose how to play audio chunks"
        )
        
        st.session_state.audio_method = audio_method
        
        # Show audio status
        st.write(f"**Status:** {st.session_state.audio_status}")
        
        # Show available/missing audio libraries
        st.write("**Audio Library Status:**")
        st.write(f"🎮 pygame: {'✅ Available' if PYGAME_AVAILABLE else '❌ Not installed'}")
        st.write(f"🎵 pydub: {'✅ Available' if PYDUB_AVAILABLE else '❌ Not installed'}")
        st.write(f"🎤 pyaudio: {'✅ Available' if PYAUDIO_AVAILABLE else '❌ Not installed'}")
        
        # Handle pending browser audio
        if st.session_state.pending_audio and "Browser" in audio_method:
            chunk_id, audio_data = st.session_state.pending_audio
            st.write(f"Playing chunk {chunk_id}")
            play_audio_browser(audio_data)
            st.session_state.pending_audio = None
        
        # Audio queue status
        queue_size = client.audio_collector.playback_queue.qsize()
        if queue_size > 0:
            st.write(f"**Queue:** {queue_size} chunks waiting")
    
    # Status section
    st.header("📊 Status")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Connection", "Connected" if client.is_connected else "Disconnected")
    
    with col2:
        st.metric("Current Slide", st.session_state.current_slide)
    
    with col3:
        st.metric("Presenting", "Yes" if st.session_state.is_presenting else "No")
    
    with col4:
        total_slides = len(st.session_state.presentation_data.get('slides', [])) if st.session_state.presentation_data else 0
        st.metric("Total Slides", total_slides)
    
    # Auto-refresh every 2 seconds
    time.sleep(2)
    st.rerun()

if __name__ == "__main__":
    main()