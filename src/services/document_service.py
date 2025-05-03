"""
Document Service for the Word AI Agent

Provides document creation, editing, and management services.
"""

import os
import time
import asyncio
from typing import Dict, Any, List, Optional, Union
from pathlib import Path

from src.core.logging import get_logger
from src.utils.error_handler import ErrorHandler
from src.utils.performance_monitor import PerformanceMonitor

logger = get_logger(__name__)

class DocumentService:
    """Document creation and management service"""
    
    def __init__(self):
        self.error_handler = ErrorHandler()
        self.performance_monitor = PerformanceMonitor()
        self.current_document = None
        self.documents = {}
        self.output_directory = Path("output")
        self._ensure_directory(self.output_directory)
    
    def _ensure_directory(self, directory: Path) -> None:
        """Ensure a directory exists"""
        directory.mkdir(parents=True, exist_ok=True)
    
    async def create_document(self, document_name: str, template: Optional[str] = None) -> Dict[str, Any]:
        """Create a new document"""
        try:
            self.performance_monitor.start_timer("create_document")
            
            if document_name in self.documents:
                logger.warning(f"Document {document_name} already exists")
                return self.documents[document_name]
            
            logger.info(f"Creating document: {document_name}")
            
            # In a real implementation, we would interact with Word
            # For now, we'll just create a mock document
            document = {
                "name": document_name,
                "created_at": time.time(),
                "modified_at": time.time(),
                "template": template,
                "content": {},
                "sections": [],
                "cursor_position": 0,
                "word_count": 0,
                "dirty": False
            }
            
            self.documents[document_name] = document
            self.current_document = document
            
            duration = self.performance_monitor.stop_timer("create_document")
            logger.info(f"Document created in {duration:.2f} seconds")
            
            return document
        except Exception as e:
            self.error_handler.handle_error(e)
            raise
    
    async def open_document(self, document_path: Union[str, Path]) -> Dict[str, Any]:
        """Open an existing document"""
        try:
            self.performance_monitor.start_timer("open_document")
            
            document_path = Path(document_path)
            if not document_path.exists():
                raise FileNotFoundError(f"Document not found: {document_path}")
            
            logger.info(f"Opening document: {document_path}")
            
            # In a real implementation, we would interact with Word
            # For now, we'll just create a mock document
            document_name = document_path.name
            document = {
                "name": document_name,
                "path": str(document_path),
                "created_at": os.path.getctime(document_path),
                "modified_at": os.path.getmtime(document_path),
                "content": {},
                "sections": [],
                "cursor_position": 0,
                "word_count": 0,
                "dirty": False
            }
            
            self.documents[document_name] = document
            self.current_document = document
            
            duration = self.performance_monitor.stop_timer("open_document")
            logger.info(f"Document opened in {duration:.2f} seconds")
            
            return document
        except Exception as e:
            self.error_handler.handle_error(e)
            raise
    
    async def save_document(self, document_name: Optional[str] = None, path: Optional[Union[str, Path]] = None) -> bool:
        """Save a document"""
        try:
            self.performance_monitor.start_timer("save_document")
            
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to save")
                return False
            
            if path:
                save_path = Path(path)
            else:
                save_path = self.output_directory / f"{document['name']}.docx"
            
            logger.info(f"Saving document to {save_path}")
            
            # In a real implementation, we would interact with Word
            # For now, we'll just simulate saving
            await asyncio.sleep(0.5)  # Simulate save operation
            
            document['modified_at'] = time.time()
            document['dirty'] = False
            document['path'] = str(save_path)
            
            duration = self.performance_monitor.stop_timer("save_document")
            logger.info(f"Document saved in {duration:.2f} seconds")
            
            return True
        except Exception as e:
            self.error_handler.handle_error(e)
            return False
    
    async def close_document(self, document_name: Optional[str] = None, save: bool = True) -> bool:
        """Close a document"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to close")
                return False
            
            if save and document['dirty']:
                await self.save_document(document_name)
            
            logger.info(f"Closing document: {document['name']}")
            
            # In a real implementation, we would interact with Word
            # For now, we'll just remove from our tracking
            if self.current_document == document:
                self.current_document = None
            
            del self.documents[document['name']]
            
            return True
        except Exception as e:
            self.error_handler.handle_error(e)
            return False
    
    async def add_section(self, section_name: str, content: str, document_name: Optional[str] = None) -> bool:
        """Add a section to a document"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to add section to")
                return False
            
            logger.info(f"Adding section '{section_name}' to document: {document['name']}")
            
            # Add section to document
            if section_name not in document['content']:
                document['sections'].append(section_name)
            
            document['content'][section_name] = content
            document['word_count'] = self._calculate_word_count(document)
            document['dirty'] = True
            
            return True
        except Exception as e:
            self.error_handler.handle_error(e)
            return False
    
    async def get_section(self, section_name: str, document_name: Optional[str] = None) -> Optional[str]:
        """Get a section from a document"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to get section from")
                return None
            
            if section_name not in document['content']:
                logger.warning(f"Section '{section_name}' not found in document: {document['name']}")
                return None
            
            return document['content'][section_name]
        except Exception as e:
            self.error_handler.handle_error(e)
            return None
    
    async def set_cursor_position(self, position: int, document_name: Optional[str] = None) -> bool:
        """Set cursor position in a document"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to set cursor position in")
                return False
            
            # In a real implementation, we would interact with Word
            # For now, we'll just update our tracking
            document['cursor_position'] = position
            
            return True
        except Exception as e:
            self.error_handler.handle_error(e)
            return False
    
    async def get_cursor_position(self, document_name: Optional[str] = None) -> Optional[int]:
        """Get cursor position in a document"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to get cursor position from")
                return None
            
            return document['cursor_position']
        except Exception as e:
            self.error_handler.handle_error(e)
            return None
    
    async def insert_text(self, text: str, document_name: Optional[str] = None) -> bool:
        """Insert text at current cursor position"""
        try:
            document = self._get_document(document_name)
            if not document:
                logger.error("No document to insert text into")
                return False
            
            logger.info(f"Inserting text at position {document['cursor_position']}")
            
            # In a real implementation, we would interact with Word
            # For now, we'll just simulate inserting
            current_section = self._get_current_section(document)
            if not current_section:
                logger.warning("No active section for text insertion")
                return False
            
            # Get current content and update
            current_content = document['content'].get(current_section, "")
            position = min(document['cursor_position'], len(current_content))
            
            new_content = current_content[:position] + text + current_content[position:]
            document['content'][current_section] = new_content
            
            # Update cursor position
            document['cursor_position'] = position + len(text)
            
            # Update word count and dirty flag
            document['word_count'] = self._calculate_word_count(document)
            document['dirty'] = True
            
            return True
        except Exception as e:
            self.error_handler.handle_error(e)
            return False
    
    def _get_document(self, document_name: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Get a document by name or the current document"""
        if document_name:
            return self.documents.get(document_name)
        return self.current_document
    
    def _get_current_section(self, document: Dict[str, Any]) -> Optional[str]:
        """Get the current section based on cursor position"""
        if not document['sections']:
            return None
        return document['sections'][0]  # Simplified for now
    
    def _calculate_word_count(self, document: Dict[str, Any]) -> int:
        """Calculate word count for a document"""
        word_count = 0
        for content in document['content'].values():
            word_count += len(content.split())
        return word_count

def create_document_service() -> DocumentService:
    """Create and return a document service instance"""
    return DocumentService()
