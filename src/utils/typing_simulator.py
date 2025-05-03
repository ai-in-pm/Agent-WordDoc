"""
Realistic typing simulation utility
"""

import random
import asyncio
from typing import Dict, Any
from abc import ABC, abstractmethod

from src.core.logging import get_logger

logger = get_logger(__name__)

class BaseTypingSimulator(ABC):
    """Base class for typing simulators"""
    
    def __init__(self, typing_mode: str):
        self.typing_mode = typing_mode
        self._setup_typing_patterns()
        
    @abstractmethod
    def _setup_typing_patterns(self) -> None:
        """Setup typing patterns based on mode"""
        pass
    
    @abstractmethod
    async def type_char(self, char: str) -> None:
        """Type a single character"""
        pass

class RealisticTypingSimulator(BaseTypingSimulator):
    """Simulates realistic human-like typing behavior"""
    
    def __init__(self, typing_mode: str):
        super().__init__(typing_mode)
        self._setup_typing_patterns()
        
    def _setup_typing_patterns(self) -> None:
        """Setup typing patterns with realistic variations"""
        self.typing_patterns = {
            'fast': {
                'base_delay': 0.05,
                'variation': 0.02,
                'error_rate': 0.01,
                'correction_delay': 0.1
            },
            'realistic': {
                'base_delay': 0.1,
                'variation': 0.05,
                'error_rate': 0.02,
                'correction_delay': 0.2
            },
            'slow': {
                'base_delay': 0.3,
                'variation': 0.1,
                'error_rate': 0.03,
                'correction_delay': 0.3
            }
        }
        
    async def type_char(self, char: str) -> None:
        """Type a single character with realistic behavior"""
        try:
            pattern = self.typing_patterns[self.typing_mode]
            base_delay = pattern['base_delay']
            variation = pattern['variation']
            
            # Simulate typing delay with variation
            delay = random.uniform(base_delay - variation, base_delay + variation)
            await asyncio.sleep(delay)
            
            # Simulate occasional typos
            if random.random() < pattern['error_rate']:
                logger.debug(f"Simulating typo for character: {char}")
                await self._simulate_typo(char, pattern)
            
            # Simulate actual typing
            logger.debug(f"Typing character: {char}")
            # Add actual typing implementation here
            
        except Exception as e:
            logger.error(f"Error typing character {char}: {str(e)}")
            raise
    
    async def _simulate_typo(self, correct_char: str, pattern: Dict[str, Any]) -> None:
        """Simulate a typo and correction"""
        # Add typo simulation logic here
        await asyncio.sleep(pattern['correction_delay'])

class TypingSimulator(RealisticTypingSimulator):
    """Main typing simulator implementation - extends RealisticTypingSimulator"""
    
    def __init__(self, typing_mode: str = "realistic"):
        super().__init__(typing_mode)
        logger.info(f"Created TypingSimulator with mode: {typing_mode}")
        self._stop_requested = False
    
    def stop_typing(self) -> None:
        """Request an immediate stop to typing"""
        self._stop_requested = True
        logger.info("Emergency typing stop requested")
    
    async def type_char(self, char: str) -> None:
        """Override to add stop check"""
        if self._stop_requested:
            logger.info("Typing stopped due to emergency stop request")
            return
        
        # Call parent implementation
        await super().type_char(char)

class TypingSimulatorFactory:
    """Factory for creating typing simulators"""
    
    @staticmethod
    def create_simulator(typing_mode: str) -> RealisticTypingSimulator:
        """Create a new typing simulator instance"""
        return RealisticTypingSimulator(typing_mode)
