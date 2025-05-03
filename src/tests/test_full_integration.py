"""
Comprehensive integration tests for the Word AI Agent

Tests the full system integration with all components.
"""

import unittest
import asyncio
import os
import time
from unittest.mock import patch, MagicMock

from src.config.config import Config
from src.agents.word_agent import AgentFactory
from src.plugins.plugin_manager import PluginManager
from src.utils.typing_simulator import RealisticTypingSimulator
from src.utils.error_handler import ErrorHandler
from src.utils.performance_monitor import PerformanceMonitor

class TestFullIntegration(unittest.TestCase):
    """Test full system integration"""
    
    def setUp(self):
        """Set up the test case"""
        self.config = Config(
            api_key="test_key",
            typing_mode="fast",  # Use fast for tests
            verbose=True,
            iterative=True,
            self_improve=True,
            self_evolve=True,
            track_position=True,
            robot_cursor=False,  # Disable for tests
            use_autoit=False,  # Disable for tests
            delay=0.1,  # Short delay for tests
            max_retries=2,
            retry_delay=0.1,
            log_level="DEBUG"
        )
        
        self.plugin_manager = PluginManager()
        self.performance_monitor = PerformanceMonitor()
        self.error_handler = ErrorHandler()
    
    def tearDown(self):
        """Clean up after tests"""
        pass
    
    @patch('src.agents.word_agent.TypingSimulator')
    def test_agent_initialization(self, mock_typing_simulator):
        """Test agent initialization"""
        # Setup mock
        mock_typing_simulator.return_value = MagicMock()
        
        # Run test
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        async def run_test():
            agent = AgentFactory.create_agent(self.config)
            await agent.initialize()
            return agent
        
        agent = loop.run_until_complete(run_test())
        
        # Assert initialization worked
        self.assertIsNotNone(agent)
        self.assertEqual(agent.config.typing_mode, "fast")
        self.assertFalse(agent.state.is_typing)
    
    @patch('src.agents.word_agent.TypingSimulator')
    def test_text_processing(self, mock_typing_simulator):
        """Test text processing"""
        # Setup mock
        mock_typing_simulator.return_value = MagicMock()
        
        # Run test
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        async def run_test():
            agent = AgentFactory.create_agent(self.config)
            await agent.initialize()
            await agent.process_text("Test text")
            return agent
        
        agent = loop.run_until_complete(run_test())
        
        # Assert text processing worked
        self.assertFalse(agent.state.is_typing)
        self.assertEqual(agent.state.retry_count, 0)
    
    @patch('src.agents.word_agent.TypingSimulator')
    def test_full_agent_lifecycle(self, mock_typing_simulator):
        """Test full agent lifecycle"""
        # Setup mock
        mock_typing_simulator.return_value = MagicMock()
        
        # Start performance monitoring
        self.performance_monitor.start_timer("full_lifecycle")
        
        # Run test
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        async def run_test():
            # Create agent
            agent = AgentFactory.create_agent(self.config)
            
            # Initialize
            await agent.initialize()
            
            # Process text
            await agent.process_text("This is a comprehensive test of the Word AI Agent.")
            
            # Finalize
            await agent.finalize()
            
            return agent
        
        agent = loop.run_until_complete(run_test())
        
        # Stop performance monitoring
        duration = self.performance_monitor.stop_timer("full_lifecycle")
        
        # Assert full lifecycle worked
        self.assertIsNotNone(agent)
        self.assertFalse(agent.state.is_typing)
        self.assertLess(duration, 5.0, "Full lifecycle should complete within 5 seconds")
    
    def test_error_handling(self):
        """Test error handling"""
        # Create a test error
        try:
            raise ValueError("Test error")
        except Exception as e:
            self.error_handler.handle_error(e)
        
        # Check error statistics
        error_stats = self.error_handler.get_error_statistics()
        self.assertIn("ValueError", error_stats)
        self.assertEqual(error_stats["ValueError"], 1)
    
    def test_plugin_system(self):
        """Test plugin system"""
        # Register a test hook
        test_hook_called = False
        
        def test_hook_callback(**kwargs):
            nonlocal test_hook_called
            test_hook_called = True
            return kwargs.get("value", None)
        
        self.plugin_manager.register_hook("test_hook", test_hook_callback)
        
        # Execute the hook
        results = self.plugin_manager.execute_hook("test_hook", value="test")
        
        # Check that the hook was called
        self.assertTrue(test_hook_called)
        self.assertEqual(results, ["test"])
    
    def test_performance_monitoring(self):
        """Test performance monitoring"""
        # Record some test metrics
        self.performance_monitor.start_timer("test_operation")
        time.sleep(0.1)  # Simulate operation
        duration = self.performance_monitor.stop_timer("test_operation")
        
        # Record a direct metric
        self.performance_monitor.record_metric("test_metric", 42.0, "units")
        
        # Get summary
        summary = self.performance_monitor.get_metrics_summary()
        
        # Check metrics
        self.assertIn("test_operation", summary)
        self.assertIn("test_metric", summary)
        self.assertGreaterEqual(duration, 0.1)
        self.assertEqual(summary["test_metric"]["average"], 42.0)

if __name__ == '__main__':
    unittest.main()
