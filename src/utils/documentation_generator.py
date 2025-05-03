"""
Documentation Generator for the Word AI Agent

Generates comprehensive documentation for the Word AI Agent system.
"""

import os
import inspect
import importlib
import pkgutil
from typing import Dict, Any, List, Optional, Type, Set
from pathlib import Path
import markdown

from src.core.logging import get_logger

logger = get_logger(__name__)

class DocumentationGenerator:
    """Generates documentation for the Word AI Agent"""
    
    def __init__(self, output_directory: str = "docs"):
        self.output_directory = Path(output_directory)
        self.output_directory.mkdir(parents=True, exist_ok=True)
        self.modules = {}
        self.classes = {}
        self.functions = {}
    
    def scan_package(self, package_name: str) -> None:
        """Scan a package for modules, classes, and functions"""
        logger.info(f"Scanning package: {package_name}")
        
        try:
            package = importlib.import_module(package_name)
            package_path = os.path.dirname(package.__file__)
            
            for _, module_name, is_pkg in pkgutil.iter_modules([package_path]):
                full_module_name = f"{package_name}.{module_name}"
                
                try:
                    module = importlib.import_module(full_module_name)
                    self._scan_module(module, full_module_name)
                    
                    if is_pkg:
                        self.scan_package(full_module_name)
                except Exception as e:
                    logger.error(f"Error scanning module {full_module_name}: {str(e)}")
        except Exception as e:
            logger.error(f"Error scanning package {package_name}: {str(e)}")
    
    def _scan_module(self, module, module_name: str) -> None:
        """Scan a module for classes and functions"""
        logger.debug(f"Scanning module: {module_name}")
        
        # Store module info
        self.modules[module_name] = {
            "name": module_name,
            "doc": inspect.getdoc(module) or "",
            "file": inspect.getfile(module),
            "classes": [],
            "functions": []
        }
        
        # Scan for classes and functions
        for name, obj in inspect.getmembers(module):
            # Skip internal/private members
            if name.startswith('_') and name != '__init__':
                continue
            
            # Skip imported classes or functions
            if inspect.isclass(obj) or inspect.isfunction(obj):
                obj_module = inspect.getmodule(obj)
                if obj_module and obj_module.__name__ != module_name:
                    continue
            
            if inspect.isclass(obj):
                self._scan_class(obj, module_name)
                self.modules[module_name]["classes"].append(name)
            elif inspect.isfunction(obj):
                self._scan_function(obj, module_name)
                self.modules[module_name]["functions"].append(name)
    
    def _scan_class(self, cls, module_name: str) -> None:
        """Scan a class for methods and attributes"""
        class_name = cls.__name__
        full_name = f"{module_name}.{class_name}"
        logger.debug(f"Scanning class: {full_name}")
        
        # Store class info
        self.classes[full_name] = {
            "name": class_name,
            "full_name": full_name,
            "module": module_name,
            "doc": inspect.getdoc(cls) or "",
            "methods": [],
            "attributes": [],
            "bases": [base.__name__ for base in cls.__bases__ if base.__name__ != 'object']
        }
        
        # Scan for methods and attributes
        for name, obj in inspect.getmembers(cls):
            # Skip internal/private members
            if name.startswith('_') and name != '__init__':
                continue
            
            if inspect.isfunction(obj) or inspect.ismethod(obj):
                method_name = f"{full_name}.{name}"
                self._scan_method(obj, method_name, full_name)
                self.classes[full_name]["methods"].append(name)
            elif not inspect.isclass(obj) and not inspect.isfunction(obj) and not inspect.isbuiltin(obj):
                # Likely an attribute
                self.classes[full_name]["attributes"].append({
                    "name": name,
                    "type": type(obj).__name__ if hasattr(obj, "__name__") else str(type(obj))
                })
    
    def _scan_method(self, method, method_name: str, class_name: str) -> None:
        """Scan a method for parameters and return type"""
        logger.debug(f"Scanning method: {method_name}")
        
        try:
            # Get method signature
            sig = inspect.signature(method)
            
            # Store method info
            self.functions[method_name] = {
                "name": method_name.split('.')[-1],
                "full_name": method_name,
                "class": class_name,
                "doc": inspect.getdoc(method) or "",
                "parameters": [],
                "return_type": str(sig.return_annotation) if sig.return_annotation != inspect.Signature.empty else "None"
            }
            
            # Store parameter info
            for param_name, param in sig.parameters.items():
                if param_name == 'self':
                    continue
                
                self.functions[method_name]["parameters"].append({
                    "name": param_name,
                    "type": str(param.annotation) if param.annotation != inspect.Signature.empty else "Any",
                    "default": str(param.default) if param.default != inspect.Signature.empty else None
                })
        except Exception as e:
            logger.error(f"Error scanning method {method_name}: {str(e)}")
    
    def _scan_function(self, func, module_name: str) -> None:
        """Scan a function for parameters and return type"""
        func_name = func.__name__
        full_name = f"{module_name}.{func_name}"
        logger.debug(f"Scanning function: {full_name}")
        
        try:
            # Get function signature
            sig = inspect.signature(func)
            
            # Store function info
            self.functions[full_name] = {
                "name": func_name,
                "full_name": full_name,
                "module": module_name,
                "doc": inspect.getdoc(func) or "",
                "parameters": [],
                "return_type": str(sig.return_annotation) if sig.return_annotation != inspect.Signature.empty else "None"
            }
            
            # Store parameter info
            for param_name, param in sig.parameters.items():
                self.functions[full_name]["parameters"].append({
                    "name": param_name,
                    "type": str(param.annotation) if param.annotation != inspect.Signature.empty else "Any",
                    "default": str(param.default) if param.default != inspect.Signature.empty else None
                })
        except Exception as e:
            logger.error(f"Error scanning function {full_name}: {str(e)}")
    
    def generate_documentation(self) -> None:
        """Generate documentation from scanned information"""
        logger.info("Generating documentation")
        
        # Generate index page
        self._generate_index()
        
        # Generate module pages
        for module_name, module_info in self.modules.items():
            self._generate_module_page(module_name, module_info)
        
        # Generate class pages
        for class_name, class_info in self.classes.items():
            self._generate_class_page(class_name, class_info)
        
        logger.info(f"Documentation generated in {self.output_directory}")
    
    def _generate_index(self) -> None:
        """Generate index page"""
        content = ["# Word AI Agent Documentation\n"]
        content.append("## Overview\n")
        content.append("This is the comprehensive documentation for the Word AI Agent system.\n")
        
        # Modules section
        content.append("## Modules\n")
        for module_name in sorted(self.modules.keys()):
            module_info = self.modules[module_name]
            content.append(f"* [{module_name}]({module_name.replace('.', '_')}.md): {module_info['doc'].split('.')[0]}\n")
        
        # Classes section
        content.append("\n## Classes\n")
        for class_name in sorted(self.classes.keys()):
            class_info = self.classes[class_name]
            content.append(f"* [{class_info['name']}]({class_name.replace('.', '_')}.md): {class_info['doc'].split('.')[0]}\n")
        
        # Write to file
        with open(self.output_directory / "index.md", "w") as f:
            f.write("".join(content))
        
        # Generate HTML version
        html_content = markdown.markdown("".join(content))
        with open(self.output_directory / "index.html", "w") as f:
            f.write(f"<!DOCTYPE html>\n<html>\n<head>\n")
            f.write(f"<title>Word AI Agent Documentation</title>\n")
            f.write(f"<style>body {{ font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; }}</style>\n")
            f.write(f"</head>\n<body>\n")
            f.write(html_content)
            f.write(f"</body>\n</html>")
    
    def _generate_module_page(self, module_name: str, module_info: Dict[str, Any]) -> None:
        """Generate documentation page for a module"""
        content = [f"# Module: {module_name}\n\n"]
        
        # Module description
        if module_info["doc"]:
            content.append(f"{module_info['doc']}\n\n")
        
        # Classes in module
        if module_info["classes"]:
            content.append("## Classes\n\n")
            for class_name in sorted(module_info["classes"]):
                full_class_name = f"{module_name}.{class_name}"
                class_info = self.classes.get(full_class_name, {})
                doc_summary = class_info.get("doc", "").split('.')[0] + "." if class_info.get("doc", "") else ""
                content.append(f"* [{class_name}]({full_class_name.replace('.', '_')}.md): {doc_summary}\n")
        
        # Functions in module
        if module_info["functions"]:
            content.append("\n## Functions\n\n")
            for func_name in sorted(module_info["functions"]):
                full_func_name = f"{module_name}.{func_name}"
                func_info = self.functions.get(full_func_name, {})
                doc_summary = func_info.get("doc", "").split('.')[0] + "." if func_info.get("doc", "") else ""
                content.append(f"### {func_name}\n\n")
                content.append(f"{doc_summary}\n\n")
                
                # Function signature
                params = []
                for param in func_info.get("parameters", []):
                    param_str = param["name"]
                    if param.get("type"):
                        param_str += f": {param['type']}"
                    if param.get("default"):
                        param_str += f" = {param['default']}"
                    params.append(param_str)
                
                signature = f"{func_name}({', '.join(params)})"
                if func_info.get("return_type"):
                    signature += f" -> {func_info['return_type']}"
                
                content.append(f"```python\n{signature}\n```\n\n")
        
        # Write to file
        output_path = self.output_directory / f"{module_name.replace('.', '_')}.md"
        with open(output_path, "w") as f:
            f.write("".join(content))
        
        # Generate HTML version
        html_content = markdown.markdown("".join(content))
        html_path = self.output_directory / f"{module_name.replace('.', '_')}.html"
        with open(html_path, "w") as f:
            f.write(f"<!DOCTYPE html>\n<html>\n<head>\n")
            f.write(f"<title>Module: {module_name}</title>\n")
            f.write(f"<style>body {{ font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; }}</style>\n")
            f.write(f"</head>\n<body>\n")
            f.write(html_content)
            f.write(f"</body>\n</html>")
    
    def _generate_class_page(self, class_name: str, class_info: Dict[str, Any]) -> None:
        """Generate documentation page for a class"""
        content = [f"# Class: {class_info['name']}\n\n"]
        
        # Class description
        if class_info["doc"]:
            content.append(f"{class_info['doc']}\n\n")
        
        # Base classes
        if class_info["bases"]:
            content.append(f"Inherits from: {', '.join(class_info['bases'])}\n\n")
        
        # Attributes
        if class_info["attributes"]:
            content.append("## Attributes\n\n")
            for attr in sorted(class_info["attributes"], key=lambda x: x["name"]):
                content.append(f"* **{attr['name']}**: {attr['type']}\n")
        
        # Methods
        if class_info["methods"]:
            content.append("\n## Methods\n\n")
            for method_name in sorted(class_info["methods"]):
                full_method_name = f"{class_name}.{method_name}"
                method_info = self.functions.get(full_method_name, {})
                doc_summary = method_info.get("doc", "").split('.')[0] + "." if method_info.get("doc", "") else ""
                content.append(f"### {method_name}\n\n")
                content.append(f"{doc_summary}\n\n")
                
                # Method signature
                params = []
                for param in method_info.get("parameters", []):
                    param_str = param["name"]
                    if param.get("type"):
                        param_str += f": {param['type']}"
                    if param.get("default"):
                        param_str += f" = {param['default']}"
                    params.append(param_str)
                
                signature = f"{method_name}({', '.join(params)})"
                if method_info.get("return_type"):
                    signature += f" -> {method_info['return_type']}"
                
                content.append(f"```python\n{signature}\n```\n\n")
        
        # Write to file
        output_path = self.output_directory / f"{class_name.replace('.', '_')}.md"
        with open(output_path, "w") as f:
            f.write("".join(content))
        
        # Generate HTML version
        html_content = markdown.markdown("".join(content))
        html_path = self.output_directory / f"{class_name.replace('.', '_')}.html"
        with open(html_path, "w") as f:
            f.write(f"<!DOCTYPE html>\n<html>\n<head>\n")
            f.write(f"<title>Class: {class_info['name']}</title>\n")
            f.write(f"<style>body {{ font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; }}</style>\n")
            f.write(f"</head>\n<body>\n")
            f.write(html_content)
            f.write(f"</body>\n</html>")

# Example usage:
# doc_gen = DocumentationGenerator()
# doc_gen.scan_package("src")
# doc_gen.generate_documentation()
