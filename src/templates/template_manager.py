"""Template manager for the Word AI Agent"""

from dataclasses import dataclass
from typing import List, Dict, Optional
import json
import os
from pathlib import Path

@dataclass
class PaperTemplate:
    """Paper template class"""
    id: str
    name: str
    description: str
    sections: List[str]
    formatting: Dict[str, str]
    citation_style: str
    example_content: Dict[str, str] = None

class TemplateManager:
    """Manager for paper templates"""
    def __init__(self, templates_dir: Optional[str] = None):
        self.templates: List[PaperTemplate] = []
        self.templates_dir = templates_dir or str(Path(__file__).parent / "presets")
        self._load_default_templates()
        self._load_custom_templates()
    
    def _load_default_templates(self):
        """Load default built-in templates"""
        self.templates = [
            PaperTemplate(
                id="research-paper",
                name="Research Paper",
                description="Standard academic research paper format with abstract, methodology, results, and discussion",
                sections=[
                    "Title",
                    "Abstract",
                    "Introduction",
                    "Literature Review",
                    "Methodology",
                    "Results",
                    "Discussion",
                    "Conclusion",
                    "References"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="APA"
            ),
            PaperTemplate(
                id="literature-review",
                name="Literature Review",
                description="Comprehensive literature review paper focusing on existing research",
                sections=[
                    "Title",
                    "Abstract",
                    "Introduction",
                    "Methodology of Review",
                    "Theoretical Framework",
                    "Literature Analysis",
                    "Research Gaps",
                    "Future Directions",
                    "Conclusion",
                    "References"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="APA"
            ),
            PaperTemplate(
                id="case-study",
                name="Case Study",
                description="In-depth analysis of a specific case or situation",
                sections=[
                    "Title",
                    "Abstract",
                    "Introduction",
                    "Background",
                    "Case Presentation",
                    "Analysis",
                    "Discussion",
                    "Recommendations",
                    "Conclusion",
                    "References"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="APA"
            ),
            PaperTemplate(
                id="thesis",
                name="Thesis/Dissertation",
                description="Comprehensive thesis or dissertation format",
                sections=[
                    "Title Page",
                    "Abstract",
                    "Acknowledgments",
                    "Table of Contents",
                    "List of Figures",
                    "List of Tables",
                    "Introduction",
                    "Literature Review",
                    "Theoretical Framework",
                    "Methodology",
                    "Results",
                    "Discussion",
                    "Conclusion",
                    "References",
                    "Appendices"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="APA"
            ),
            PaperTemplate(
                id="technical-report",
                name="Technical Report",
                description="Structured technical report for engineering or scientific research",
                sections=[
                    "Title Page",
                    "Executive Summary",
                    "Table of Contents",
                    "Introduction",
                    "Technical Background",
                    "Methodology",
                    "Technical Analysis",
                    "Results",
                    "Discussion",
                    "Recommendations",
                    "Conclusion",
                    "References",
                    "Technical Appendices"
                ],
                formatting={
                    "font": "Arial",
                    "size": "11pt",
                    "spacing": "1.5",
                    "margins": "1 inch"
                },
                citation_style="IEEE"
            ),
            PaperTemplate(
                id="journal-article",
                name="Journal Article",
                description="Format for peer-reviewed journal publication",
                sections=[
                    "Title",
                    "Abstract",
                    "Keywords",
                    "Introduction",
                    "Materials and Methods",
                    "Results",
                    "Discussion",
                    "Conclusion",
                    "Acknowledgments",
                    "References"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="Vancouver"
            ),
            PaperTemplate(
                id="systematic-review",
                name="Systematic Review",
                description="Comprehensive systematic review of literature on a specific topic",
                sections=[
                    "Title",
                    "Abstract",
                    "Introduction",
                    "Review Protocol",
                    "Search Strategy",
                    "Inclusion and Exclusion Criteria",
                    "Data Extraction",
                    "Quality Assessment",
                    "Results",
                    "Discussion",
                    "Conclusion",
                    "References",
                    "PRISMA Diagram"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "12pt",
                    "spacing": "double",
                    "margins": "1 inch"
                },
                citation_style="APA"
            ),
            PaperTemplate(
                id="conference-paper",
                name="Conference Paper",
                description="Concise format for academic conference submission",
                sections=[
                    "Title",
                    "Abstract",
                    "Keywords",
                    "Introduction",
                    "Related Work",
                    "Methodology",
                    "Evaluation",
                    "Results",
                    "Discussion",
                    "Conclusion",
                    "References"
                ],
                formatting={
                    "font": "Times New Roman",
                    "size": "10pt",
                    "spacing": "single",
                    "margins": "0.75 inch"
                },
                citation_style="ACM"
            )
        ]
    
    def _load_custom_templates(self):
        """Load custom templates from templates directory"""
        try:
            os.makedirs(self.templates_dir, exist_ok=True)
            
            # Load all JSON files in templates directory
            for file_path in Path(self.templates_dir).glob("*.json"):
                try:
                    with open(file_path, "r") as f:
                        template_data = json.load(f)
                        
                    # Create template from JSON data
                    template = PaperTemplate(
                        id=template_data.get("id", file_path.stem),
                        name=template_data.get("name", file_path.stem),
                        description=template_data.get("description", ""),
                        sections=template_data.get("sections", []),
                        formatting=template_data.get("formatting", {}),
                        citation_style=template_data.get("citation_style", "APA"),
                        example_content=template_data.get("example_content", None)
                    )
                    
                    # Add to templates list if not duplicate ID
                    if not any(t.id == template.id for t in self.templates):
                        self.templates.append(template)
                except Exception as e:
                    print(f"Error loading template {file_path}: {str(e)}")
        except Exception as e:
            print(f"Error loading custom templates: {str(e)}")
    
    def get_template_by_id(self, template_id: str) -> Optional[PaperTemplate]:
        """Get template by ID"""
        for template in self.templates:
            if template.id == template_id:
                return template
        return None
    
    def get_template_names(self) -> List[str]:
        """Get list of template names for UI display"""
        return [template.name for template in self.templates]
    
    def save_custom_template(self, template: PaperTemplate) -> bool:
        """Save a custom template to file"""
        try:
            os.makedirs(self.templates_dir, exist_ok=True)
            
            # Create file path
            file_path = os.path.join(self.templates_dir, f"{template.id}.json")
            
            # Convert template to dict
            template_dict = {
                "id": template.id,
                "name": template.name,
                "description": template.description,
                "sections": template.sections,
                "formatting": template.formatting,
                "citation_style": template.citation_style,
                "example_content": template.example_content
            }
            
            # Save to file
            with open(file_path, "w") as f:
                json.dump(template_dict, f, indent=4)
            
            return True
        except Exception as e:
            print(f"Error saving template: {str(e)}")
            return False
