"""
Test suite for the Word AI Agent
"""

import asyncio
import pytest
from pathlib import Path
from typing import Dict, Any

from src.config.config import Config
from src.agents.word_agent import WordAIAgent
from src.utils.typing_simulator import RealisticTypingSimulator
from src.utils.error_handler import ErrorHandler

@pytest.fixture
def config() -> Config:
    """Create a test configuration"""
    return Config(
        api_key="test_key",
        typing_mode="realistic",
        verbose=True,
        iterative=True,
        self_improve=True,
        self_evolve=True,
        track_position=True,
        robot_cursor=True,
        use_autoit=True,
        delay=1.0,
        max_retries=3,
        retry_delay=1.0,
        log_level="DEBUG"
    )

@pytest.fixture
def agent(config: Config) -> WordAIAgent:
    """Create a test agent"""
    return WordAIAgent(config)

@pytest.fixture
def typing_simulator() -> RealisticTypingSimulator:
    """Create a typing simulator"""
    return RealisticTypingSimulator("realistic")

@pytest.fixture
def error_handler() -> ErrorHandler:
    """Create an error handler"""
    return ErrorHandler()

@pytest.mark.asyncio
async def test_agent_initialization(agent: WordAIAgent):
    """Test agent initialization"""
    await agent.initialize()
    assert agent.state.is_typing == False
    assert agent.state.error_count == 0

@pytest.mark.asyncio
async def test_text_processing(agent: WordAIAgent):
    """Test text processing"""
    test_text = "Hello, World!"
    await agent.process_text(test_text)
    assert agent.state.is_typing == False
    assert agent.state.current_position == len(test_text)

@pytest.mark.asyncio
async def test_typing_simulation(typing_simulator: RealisticTypingSimulator):
    """Test typing simulation"""
    test_char = "A"
    await typing_simulator.type_char(test_char)
    # Add assertions for typing behavior

@pytest.mark.asyncio
async def test_error_handling(error_handler: ErrorHandler):
    """Test error handling"""
    try:
        raise ValueError("Test error")
    except Exception as e:
        error_handler.handle_error(e)
    assert error_handler.get_error_statistics()["ValueError"] == 1

@pytest.mark.asyncio
async def test_agent_finalization(agent: WordAIAgent):
    """Test agent finalization"""
    await agent.finalize()
    assert agent.state.is_typing == False

@pytest.mark.asyncio
async def test_config_validation():
    """Test configuration validation"""
    # Test invalid API key
    with pytest.raises(ValueError):
        Config(api_key="")
    
    # Test invalid typing mode
    with pytest.raises(ValueError):
        Config(typing_mode="invalid")

@pytest.mark.asyncio
async def test_performance(agent: WordAIAgent):
    """Test performance with large text"""
    large_text = "A" * 10000
    start_time = time.time()
    await agent.process_text(large_text)
    end_time = time.time()
    processing_time = end_time - start_time
    assert processing_time < 10  # Should process within 10 seconds

@pytest.mark.asyncio
async def test_error_recovery(agent: WordAIAgent):
    """Test error recovery mechanism"""
    # Simulate error during typing
    test_text = "Error Test"
    await agent.process_text(test_text)
    assert agent.state.retry_count <= agent.config.max_retries
