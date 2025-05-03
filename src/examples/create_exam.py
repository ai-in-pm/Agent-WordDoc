import os
import sys
import re
from src.word_ai_agent import WordAIAssistant
from src.evm_paper_generator import EVMPaperGenerator

def create_exam():
    # Initialize the EVMPaperGenerator
    evm_generator = EVMPaperGenerator(topic="Putnam Mathematical Competition Exam")
    
    # Initialize the WordAIAssistant
    assistant = WordAIAssistant(
        typing_mode='realistic',
        verbose=True,
        topic="Putnam Mathematical Competition Exam",
        template=None
    )
    
    print("\nGenerating exam document...")
    
    # Define exam sections
    exam_content = [
        {
            "title": "Putnam Mathematical Competition Exam",
            "content": "Date: December 7, 2024\n\nInstructions:\n1. This exam consists of two sessions, A and B, each containing 6 problems.\n2. You have 3 hours for each session.\n3. Each problem is worth 10 points.\n4. Show all your work clearly for full credit.\n5. No calculators or electronic devices are allowed."
        },
        {
            "title": "Section A",
            "content": ""
        },
        {
            "title": "Section B",
            "content": ""
        }
    ]
    
    # Read the LaTeX file with problems
    tex_file_path = r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\Math Problems.tex"
    with open(tex_file_path, 'r') as file:
        tex_content = file.read()
    
    # Extract problems
    pattern = r'\\item\[(.[1-6])\](.*?)(?=\\item|\\end{itemize})'
    problems = []
    for match in re.finditer(pattern, tex_content, re.DOTALL):
        problem_num = match.group(1)
        problem_text = match.group(2).strip()
        problems.append(f"Problem {problem_num}: {problem_text}")
    
    # Split problems into sections
    section_a = problems[:6]
    section_b = problems[6:]
    
    # Format problems for each section
    exam_content[1]["content"] = "\n\n".join(section_a)
    exam_content[2]["content"] = "\n\n".join(section_b)
    
    # Generate and type the exam
    for section in exam_content:
        # Type section heading
        assistant.type_heading(section["title"], level=1)
        assistant.press_enter()
        
        # Type section content
        assistant.type_text(section["content"])
        assistant.press_enter(2)
    
    # Save the document
    filename = "Putnam_Exam.docx"
    evm_generator.save_document(filename)
    print(f"\nExam generation complete! The document has been saved as '{filename}'")

if __name__ == "__main__":
    create_exam()
