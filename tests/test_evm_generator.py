import unittest
from src.evm_paper_generator import EVMPaperGenerator
import os

class TestEVMPaperGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = EVMPaperGenerator()

    def test_document_creation(self):
        """Test if document is created successfully"""
        self.assertIsNotNone(self.generator.document)

    def test_content_generation(self):
        """Test if content is generated for a section"""
        content = self.generator.generate_section_content("Introduction")
        self.assertTrue(len(content) > 0)

    def test_section_addition(self):
        """Test if sections are added correctly"""
        initial_paragraphs = len(self.generator.document.paragraphs)
        self.generator.add_section("Test Section", "This is a test section content")
        # Check if paragraphs were added (heading + content = 2 more paragraphs)
        self.assertTrue(len(self.generator.document.paragraphs) > initial_paragraphs)

    def test_save_document(self):
        """Test if document is saved successfully"""
        test_filename = "test_paper.docx"
        self.generator.save_document(test_filename)
        self.assertTrue(os.path.exists(test_filename))
        # Clean up
        os.remove(test_filename)

if __name__ == '__main__':
    unittest.main()
