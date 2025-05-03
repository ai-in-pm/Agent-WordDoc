"""
Spell Check Plugin for the Word AI Agent

This plugin adds real-time spell checking capability to the agent.
"""

import re
from typing import Dict, Any, List, Optional

__version__ = "1.0.0"
__description__ = "Real-time spell checking for the Word AI Agent"
__author__ = "AI Assistant"

class SpellChecker:
    """Simple spell checker implementation"""
    
    def __init__(self):
        # Load a simple dictionary for demonstration
        self.dictionary = self._load_dictionary()
        self.corrections = {
            "teh": "the",
            "adn": "and",
            "taht": "that",
            "wiht": "with",
            "fo": "of",
            "toi": "to",
            "ot": "to",
            "int eh": "in the",
            "fo teh": "of the"
        }
    
    def _load_dictionary(self) -> set:
        """Load a simple English dictionary"""
        # This is a simplified version - in a real implementation, 
        # we would load a comprehensive dictionary
        words = {
            "the", "and", "that", "with", "for", "this", "on", "is", "are", "was",
            "be", "to", "have", "it", "in", "of", "not", "as", "at", "by", "but",
            "from", "or", "an", "they", "which", "you", "can", "more", "will", "some",
            "about", "when", "up", "out", "if", "so", "time", "no", "just", "than",
            "other", "into", "its", "then", "two", "only", "may", "first", "also", "now",
            "any", "like", "new", "over", "such", "our", "after", "even", "because", "these",
            "most", "through", "before", "between", "him", "we", "she", "who", "been", "both",
            "all", "each", "people", "year", "how", "well", "one", "where", "their", "what"
        }
        return words
    
    def check_text(self, text: str) -> List[Dict[str, Any]]:
        """Check text for spelling errors"""
        errors = []
        words = re.findall(r'\b\w+\b', text.lower())
        
        for word in words:
            if word not in self.dictionary and len(word) > 1:
                correction = self.corrections.get(word, None)
                errors.append({
                    "word": word,
                    "position": text.lower().find(word),
                    "suggestion": correction
                })
        
        return errors
    
    def correct_text(self, text: str) -> str:
        """Correct spelling errors in text"""
        corrected_text = text
        for error, correction in self.corrections.items():
            corrected_text = re.sub(r'\b' + error + r'\b', correction, corrected_text, flags=re.IGNORECASE)
        
        return corrected_text

def before_typing_callback(text: str, **kwargs) -> str:
    """Check spelling before typing"""
    spell_checker = SpellChecker()
    return spell_checker.correct_text(text)

def register_plugin(plugin_manager):
    """Register the plugin with the plugin manager"""
    plugin_manager.register_hook("before_typing", before_typing_callback)
