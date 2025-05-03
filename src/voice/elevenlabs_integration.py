"""
ElevenLabs Voice Integration Module for Agent WordDoc

This module integrates ElevenLabs capabilities for speech-to-text and text-to-speech
to enable voice commands and responses in the Agent WordDoc application.
"""

import os
import sys
import asyncio
import tempfile
import logging
from typing import Optional, Callable, List, Dict, Any
import threading
import queue
import time
from pathlib import Path

# Add the elevenlabs library path to PYTHONPATH
ELEVENLABS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                               'voice', 'elevenlabs-python-main', 'src')
sys.path.append(ELEVENLABS_PATH)

# Import elevenlabs components
try:
    from elevenlabs import generate, stream, save, Voice, VoiceSettings
    from elevenlabs import voices as list_voices
    from elevenlabs import set_api_key
    from elevenlabs.api import Speech, User
    from elevenlabs.client import ElevenLabs
    from elevenlabs.speech_to_text.client import SpeechToTextClient
except ImportError as e:
    raise ImportError(f"Failed to import ElevenLabs library: {e}. Make sure the library is installed.")

from src.core.logging import get_logger
from src.utils.environment import load_env_variables

# Initialize logger
logger = get_logger(__name__)

class ElevenLabsIntegration:
    """ElevenLabs voice integration for Agent WordDoc"""
    
    def __init__(self, api_key: Optional[str] = None, voice_id: Optional[str] = None):
        # Load API key from environment if not provided
        if api_key is None:
            env_vars = load_env_variables()
            api_key = env_vars.get('ELEVENLABS_API_KEY')
            if not api_key:
                raise ValueError("ElevenLabs API key not found. Please set ELEVENLABS_API_KEY in .env file.")
            
        # Set voice ID from environment if not provided
        if voice_id is None:
            env_vars = load_env_variables()
            voice_id = env_vars.get('ELEVENLABS_VOICE_ID')
        
        # Initialize ElevenLabs API
        set_api_key(api_key)
        self.api_key = api_key
        self.voice_id = voice_id
        self.client = ElevenLabs(api_key=api_key)
        
        # Initialize STT components
        self.stt_client = self.client.speech_to_text
        self.listen_thread = None
        self.command_queue = queue.Queue()
        self.is_listening = False
        self.command_callback = None
        
        # Record available voices
        try:
            self.available_voices = list_voices()
            logger.info(f"Found {len(self.available_voices)} voices in ElevenLabs account")
        except Exception as e:
            logger.error(f"Failed to retrieve voices: {e}")
            self.available_voices = []
        
        logger.info("ElevenLabs integration initialized successfully")
    
    def speak(self, text: str, voice_id: Optional[str] = None, stream_audio: bool = True) -> bool:
        """Convert text to speech using ElevenLabs API"""
        try:
            # Use provided voice_id or default
            current_voice_id = voice_id or self.voice_id
            
            if stream_audio:
                audio = generate(
                    text=text,
                    voice=current_voice_id,
                    model="eleven_monolingual_v1"
                )
                stream(audio)
            else:
                audio = generate(
                    text=text,
                    voice=current_voice_id,
                    model="eleven_monolingual_v1"
                )
                # Save to a temporary file and play
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
                    save(audio, tmp.name)
                    # Here we could use a proper audio player if needed
                    logger.info(f"Audio saved to {tmp.name}")
                    
            logger.info(f"Text spoken successfully: {text[:50]}...")
            return True   # Return success status
                
        except Exception as e:
            logger.error(f"Failed to convert text to speech: {e}")
            return False  # Return failure status
    
    def get_available_voices(self) -> List[Voice]:
        """Get list of available voices"""
        return self.available_voices
    
    def start_voice_command_listener(self, callback: Callable[[str], None], model_id: str = "scribe_v1") -> None:
        """Start listening for voice commands in the background"""
        if self.is_listening:
            logger.warning("Voice command listener is already running")
            return
        
        self.command_callback = callback
        self.is_listening = True
        self.listen_thread = threading.Thread(target=self._voice_command_worker, args=(model_id,))
        self.listen_thread.daemon = True
        self.listen_thread.start()
        logger.info("Voice command listener started")
    
    def stop_voice_command_listener(self) -> None:
        """Stop the voice command listener thread"""
        if not self.is_listening:
            return
        
        self.is_listening = False
        if self.listen_thread and self.listen_thread.is_alive():
            self.listen_thread.join(timeout=2.0)
        logger.info("Voice command listener stopped")
    
    def _voice_command_worker(self, model_id: str) -> None:
        """Worker thread for voice command processing"""
        import pyaudio
        from array import array
        
        # Audio recording parameters
        FORMAT = pyaudio.paInt16
        CHANNELS = 1
        RATE = 16000
        CHUNK_SIZE = 1024
        SILENCE_THRESHOLD = 800  # Adjust based on microphone sensitivity
        SILENCE_DURATION = 2  # Seconds of silence to consider the command complete
        
        audio = pyaudio.PyAudio()
        logger.info("Initializing audio recording for voice commands")
        
        try:
            while self.is_listening:
                # Open stream for recording
                stream = audio.open(
                    format=FORMAT,
                    channels=CHANNELS,
                    rate=RATE,
                    input=True,
                    frames_per_buffer=CHUNK_SIZE
                )
                
                logger.info("Listening for voice commands... (speak now)")
                
                # Variables for recording
                frames = []
                silent_chunks = 0
                is_speaking = False
                max_silent_chunks = int(SILENCE_DURATION * RATE / CHUNK_SIZE)
                
                # Listen for voice commands
                while self.is_listening:
                    data = stream.read(CHUNK_SIZE)
                    frames.append(data)
                    
                    # Check audio levels
                    audio_data = array('h', data)
                    volume = max(audio_data) if audio_data else 0
                    
                    # Detect if speaking
                    if volume > SILENCE_THRESHOLD:
                        is_speaking = True
                        silent_chunks = 0
                    elif is_speaking:
                        silent_chunks += 1
                        
                    # End of command detection
                    if is_speaking and silent_chunks > max_silent_chunks:
                        logger.info("Command detected, processing...")
                        break
                
                # If we actually recorded something
                if is_speaking:
                    stream.stop_stream()
                    stream.close()
                    
                    # Save audio to temporary file
                    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_file:
                        import wave
                        wf = wave.open(temp_file.name, 'wb')
                        wf.setnchannels(CHANNELS)
                        wf.setsampwidth(audio.get_sample_size(FORMAT))
                        wf.setframerate(RATE)
                        wf.writeframes(b''.join(frames))
                        wf.close()
                        
                        # Process with ElevenLabs STT
                        try:
                            with open(temp_file.name, 'rb') as audio_file:
                                logger.info(f"Sending audio to ElevenLabs for transcription")
                                from elevenlabs.core import File
                                result = self.stt_client.convert(
                                    model_id=model_id,
                                    file=File(content=audio_file.read()),
                                )
                                
                                if result and hasattr(result, 'text'):
                                    command_text = result.text
                                    logger.info(f"Voice command recognized: {command_text}")
                                    
                                    # Process command through callback
                                    if self.command_callback:
                                        # Use event loop to ensure callback is thread-safe
                                        asyncio.run(self._call_command_callback(command_text))
                                else:
                                    logger.warning("No text was transcribed from the audio")
                        except Exception as e:
                            logger.error(f"Error transcribing voice command: {e}")
                    
                    # Clean up temp file
                    try:
                        os.unlink(temp_file.name)
                    except:
                        pass
                else:
                    stream.stop_stream()
                    stream.close()
        
        except Exception as e:
            logger.error(f"Error in voice command worker: {e}")
        finally:
            audio.terminate()
            logger.info("Voice command listener terminated")
    
    async def _call_command_callback(self, command_text: str) -> None:
        """Call the command callback in an async safe way"""
        if self.command_callback:
            try:
                self.command_callback(command_text)
            except Exception as e:
                logger.error(f"Error in command callback: {e}")

# Helper functions for direct use
def speak_text(text: str, api_key: Optional[str] = None, voice_id: Optional[str] = None) -> None:
    """Quick helper to speak text without creating a full integration instance"""
    integration = ElevenLabsIntegration(api_key, voice_id)
    integration.speak(text)
