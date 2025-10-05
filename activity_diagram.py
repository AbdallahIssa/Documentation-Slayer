#!/usr/bin/env python3
"""
Activity Diagram Generator Module
Extracted from C2AD.py for integration with Documentation Slayer
"""

import re
import os
import subprocess
from pathlib import Path

try:
    from graphviz import Digraph
    GRAPHVIZ_AVAILABLE = True
except ImportError:
    GRAPHVIZ_AVAILABLE = False
    print("Warning: graphviz not installed. Activity diagrams will not be generated.")

class ActivityDiagramGenerator:
    def __init__(self):
        self.graph = None
        self.node_counter = 0
        self.current_function = None
        self.type_mappings = {}
        
    def generate_node_id(self):
        """Generate unique node ID"""
        self.node_counter += 1
        return f"node_{self.node_counter}"
    
    def create_node(self, label, shape="box", style="filled", fillcolor="#E1F5FE"):
        """Create a node in the graph"""
        node_id = self.generate_node_id()
        self.graph.node(node_id, label=label, shape=shape, style=style, fillcolor=fillcolor,
                        width="2.5", height="0.6", fixedsize="false", 
                        fontname="Arial", fontsize="10", penwidth="1.5",
                        color="#1976D2" if shape == "box" else "#4CAF50")
        return node_id
    
    def create_decision_node(self, condition):
        """Create a diamond-shaped decision node"""
        node_id = self.generate_node_id()
        self.graph.node(node_id, label=condition, shape="diamond", style="filled", 
                       fillcolor="#FFF9C4", width="3.0", height="1.5", fixedsize="false",
                       fontname="Arial", fontsize="10", penwidth="1.5", color="#F57F17")
        return node_id
    
    def create_merge_node(self):
        """Create a merge diamond node (EA style)"""
        node_id = self.generate_node_id()
        self.graph.node(node_id, label="", shape="diamond", style="filled", 
                       fillcolor="#C8E6C9", width="0.8", height="0.8", fixedsize="true",
                       color="#388E3C", penwidth="2")
        return node_id
    
    def create_start_end_node(self, label, is_start=True):
        """Create start/end nodes"""
        fillcolor = "#4CAF50" if is_start else "#F44336"
        node_id = self.generate_node_id()
        self.graph.node(node_id, label="", shape="circle", style="filled", 
                       fillcolor=fillcolor, width="0.5", height="0.5", fixedsize="true",
                       color=fillcolor, penwidth="2")
        return node_id
    
    def connect_nodes(self, from_node, to_node, label=""):
        """Connect two nodes with an edge"""
        if label:
            self.graph.edge(from_node, to_node, label=label, fontname="Arial", fontsize="9", color="#424242")
        else:
            self.graph.edge(from_node, to_node, color="#424242")
    
    def setup_graphviz_windows(self):
        """Setup Graphviz for Windows systems"""
        try:
            subprocess.run(['dot', '-V'], capture_output=True, check=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("Graphviz not found in PATH. Attempting to locate...")
            
            # Common Graphviz installation paths on Windows
            possible_paths = [
                r"C:\Program Files\Graphviz\bin",
                r"C:\Program Files (x86)\Graphviz\bin",
                r"C:\Graphviz\bin",
                r"C:\tools\graphviz\bin",
                os.path.expanduser("~/AppData/Local/Programs/Graphviz/bin")
            ]
            
            for path in possible_paths:
                if os.path.exists(os.path.join(path, 'dot.exe')):
                    print(f"Found Graphviz at: {path}")
                    # Add to PATH for this session
                    os.environ['PATH'] = path + os.pathsep + os.environ.get('PATH', '')
                    return True
            
            return False
    
    def create_graph_with_fallback(self):
        """Create Graphviz graph with Windows compatibility"""
        if not GRAPHVIZ_AVAILABLE:
            print("Graphviz not available. Please install: pip install graphviz")
            return None
            
        try:
            # For Windows, try to setup Graphviz first
            if os.name == 'nt':  # Windows
                if not self.setup_graphviz_windows():
                    print("\nGraphviz installation issues detected!")
                    print("Please install Graphviz:")
                    print("1. Download from: https://graphviz.org/download/")
                    print("2. Or use chocolatey: choco install graphviz")
                    print("3. Or use conda: conda install graphviz")
                    print("4. Make sure to add Graphviz\\bin to your PATH")
                    return None
            
            # Create the graph
            graph = Digraph(comment='Activity Diagram')
            graph.attr(rankdir='TB', bgcolor='white', dpi='150')
            graph.attr('node', fontname='Arial', fontsize='10', margin='0.3,0.1', color='#424242')
            graph.attr('edge', fontname='Arial', fontsize='9', color='#424242', penwidth='1.2')
            graph.attr('graph', splines='polyline', nodesep='0.8', ranksep='1.0', overlap='false')
            
            # Set engine explicitly for better compatibility
            graph.engine = 'dot'
            
            return graph
            
        except Exception as e:
            print(f"Error creating graph: {e}")
            return None
    
    def preprocess_c_code(self, c_code):
        """Basic preprocessing to handle common C constructs"""
        # Store original types for display
        self.type_mappings = {}

        # Handle common automotive/AUTOSAR patterns first
        c_code = self.handle_automotive_patterns(c_code)

        # Remove comments
        c_code = re.sub(r'//.*', '', c_code)
        c_code = re.sub(r'/\*.*?\*/', '', c_code, flags=re.DOTALL)

        # Extract and store #define macros before removing them
        define_pattern = r'#define\s+(\w+)(?:\([^)]*\))?\s+(.+)'
        defines = {}
        for match in re.finditer(define_pattern, c_code):
            macro_name = match.group(1)
            macro_value = match.group(2).strip() if match.group(2) else ""
            defines[macro_name] = macro_value

        # Replace custom types with standard C types to help parser
        custom_type_patterns = [
            (r'\b[A-Z][A-Za-z0-9_]*_[ES]_[A-Za-z0-9_]*Type\b', 'int'),  # Enum/Struct types
            (r'\buint(8|16|32|64)\b', 'unsigned int'),                    # AUTOSAR integer types
            (r'\bsint(8|16|32|64)\b', 'int'),                           # AUTOSAR signed types
            (r'\bboolean\b', 'int'),                                     # AUTOSAR boolean
            (r'\b[A-Z][A-Za-z0-9_]*_[A-Za-z0-9_]*Type\b', 'int'),      # Generic custom types
        ]
        
        for pattern, replacement in custom_type_patterns:
            # Find and store all matches before replacement
            matches = re.findall(pattern, c_code)
            for match in matches:
                if isinstance(match, tuple):
                    original_type = match[0] if match[0] else match
                else:
                    original_type = match
                self.type_mappings[replacement] = original_type
            
            # Replace for parser compatibility
            c_code = re.sub(pattern, replacement, c_code)

        # Remove common includes
        c_code = re.sub(r'#include\s*<[^>]+>', '', c_code)
        c_code = re.sub(r'#include\s*"[^"]+"', '', c_code)

        # Expand FUNC macro specifically
        if 'FUNC' in defines:
            func_pattern = r'FUNC\s*\(\s*([^,]+)\s*,\s*([^)]+)\s*\)\s+(\w+)\s*\('
            def expand_func(match):
                rettype = match.group(1).strip()
                memclass = match.group(2).strip()
                func_name = match.group(3).strip()
                return f'{rettype} {func_name}('
            
            c_code = re.sub(func_pattern, expand_func, c_code)
        
        # Handle other common macro patterns
        for macro_name, macro_value in defines.items():
            if macro_name != 'FUNC':
                # Simple text replacement for basic macros
                c_code = re.sub(rf'\b{macro_name}\b', macro_value, c_code)
        
        # Remove other preprocessor directives
        c_code = re.sub(r'#define.*', '', c_code)
        c_code = re.sub(r'#ifdef.*', '', c_code)
        c_code = re.sub(r'#ifndef.*', '', c_code)
        c_code = re.sub(r'#endif.*', '', c_code)
        c_code = re.sub(r'#if.*', '', c_code)
        c_code = re.sub(r'#else.*', '', c_code)
        c_code = re.sub(r'#elif.*', '', c_code)
        c_code = re.sub(r'#pragma.*', '', c_code)
        
        return c_code

    def handle_automotive_patterns(self, c_code):
        """Handle specific automotive/AUTOSAR code patterns"""
        
        # Remove complex function documentation blocks
        c_code = re.sub(r'/\*\*[^*]*\*+(?:[^/*][^*]*\*+)*/', '', c_code, flags=re.MULTILINE)
        
        # Handle AUTOSAR function declarations with complex return types
        # Replace complex return types with simple int
        c_code = re.sub(r'^[A-Z][A-Za-z0-9_]*_[ES]_[A-Za-z0-9_]*Type\s+', 'int ', c_code, flags=re.MULTILINE)
        
        # Handle NULL_PTR constant
        c_code = re.sub(r'\bNULL_PTR\b', 'NULL', c_code)
        
        # Handle STD_ON/STD_OFF constants
        c_code = re.sub(r'\bSTD_ON\b', '1', c_code)
        c_code = re.sub(r'\bSTD_OFF\b', '0', c_code)
        
        # Handle AUTOSAR memory classifications
        c_code = re.sub(r'\b[A-Z][A-Za-z0-9_]*_CODE\b', '', c_code)
        
        # Handle complex variable declarations
        c_code = re.sub(r'^extern\s+[^;]+;', '', c_code, flags=re.MULTILINE)
        
        # Simplify complex pointer types
        c_code = re.sub(r'\*\s*const\s+', '* ', c_code)
        c_code = re.sub(r'const\s+\*', '* ', c_code)
        
        # Handle continue statements that might confuse parser
        c_code = re.sub(r'\bcontinue\s*;', 'break;', c_code)
        
        return c_code
    
    def generate_simple_activity_diagram(self, c_file_path, output_prefix="activity"):
        """Generate simple activity diagrams for functions using regex parsing"""
        try:
            if not GRAPHVIZ_AVAILABLE:
                print("❌ Graphviz not available. Please install: pip install graphviz")
                return False
                
            # Read file content
            with open(c_file_path, 'r', encoding='utf-8') as f:
                c_code = f.read()
            
            print("Preprocessing C code for activity diagram generation...")
            # Preprocess the code
            c_code = self.preprocess_c_code(c_code)
            
            # Find functions using regex (more robust than full parsing)
            func_patterns = [
                r'(?:FUNC\s*\([^)]*\)\s*)?(?:static\s+)?(?:inline\s+)?[a-zA-Z_][a-zA-Z0-9_\s\*]*\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\([^)]*\)\s*\{',
                r'^(?:static\s+)?(?:inline\s+)?[a-zA-Z_]\w*\s+([a-zA-Z_]\w*)\s*\([^)]*\)\s*\{',
            ]
            
            functions = set()
            for pattern in func_patterns:
                matches = re.findall(pattern, c_code, re.MULTILINE)
                functions.update(matches)
            
            if not functions:
                print("No functions found for activity diagram generation")
                return False
            
            print(f"Found functions: {', '.join(functions)}")
            
            function_count = 0
            for func_name in functions:
                if len(func_name) < 2:  # Skip single character matches
                    continue
                    
                print(f"Generating activity diagram for: {func_name}")
                
                # Reset for each function
                self.node_counter = 0
                self.graph = self.create_graph_with_fallback()
                
                if self.graph:
                    # Create simple activity diagram structure
                    start_node = self.create_start_end_node("", is_start=True)
                    func_node = self.create_node(f"{func_name}()", fillcolor="#BBDEFB")
                    
                    # Try to find some basic control flow
                    func_body = self.extract_function_body(c_code, func_name)
                    flow_nodes = self.create_basic_flow(func_body)
                    
                    # Connect nodes
                    self.connect_nodes(start_node, func_node)
                    
                    prev_node = func_node
                    for node in flow_nodes:
                        self.connect_nodes(prev_node, node)
                        prev_node = node
                    
                    end_node = self.create_start_end_node("", is_start=False)
                    self.connect_nodes(prev_node, end_node)
                    
                    # Render the diagram
                    output_file = f"{output_prefix}_{func_name}"
                    try:
                        self.graph.render(output_file, format='png', cleanup=True)
                        print(f"✅ Generated: {output_file}.png")
                        function_count += 1
                    except Exception as e:
                        print(f"Error rendering {func_name}: {e}")
                        # Try SVG as fallback
                        try:
                            self.graph.render(output_file, format='svg', cleanup=True)
                            print(f"✅ Generated SVG: {output_file}.svg")
                            function_count += 1
                        except:
                            print(f"❌ Failed to generate diagram for {func_name}")
            
            return function_count > 0
            
        except Exception as e:
            print(f"Error generating activity diagrams: {e}")
            return False
    
    def extract_function_body(self, c_code, func_name):
        """Extract function body using regex"""
        # Find function definition
        pattern = rf'\b{re.escape(func_name)}\s*\([^)]*\)\s*\{{([^{{}}]*(?:\{{[^{{}}]*\}}[^{{}}]*)*)\}}'
        match = re.search(pattern, c_code, re.DOTALL)
        
        if match:
            return match.group(1)
        return ""
    
    def create_basic_flow(self, func_body):
        """Create basic flow nodes from function body"""
        nodes = []
        
        if not func_body.strip():
            nodes.append(self.create_node("Function body", fillcolor="#E1F5FE"))
            return nodes
        
        # Look for basic patterns
        if re.search(r'\bif\s*\(', func_body):
            nodes.append(self.create_decision_node("Conditional logic"))
        
        if re.search(r'\b(for|while)\s*\(', func_body):
            nodes.append(self.create_decision_node("Loop logic"))
        
        if re.search(r'\breturn\b', func_body):
            nodes.append(self.create_node("Return statement", fillcolor="#FFCDD2"))
        
        # If no specific patterns found, create generic processing node
        if not nodes:
            nodes.append(self.create_node("Process function logic", fillcolor="#E1F5FE"))
        
        return nodes


def generate_activity_diagrams(c_file_path, output_dir, output_prefix="activity"):
    """
    Main function to generate activity diagrams from C file
    This is the interface function called from parser.py
    """
    try:
        if not GRAPHVIZ_AVAILABLE:
            print("❌ Graphviz module not available.")
            print("Please install it with: pip install graphviz")
            print("Also ensure Graphviz software is installed on your system.")
            return False
        
        print(f"Generating activity diagrams from: {c_file_path}")
        
        # Create output directory if it doesn't exist
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        # Generate diagrams
        generator = ActivityDiagramGenerator()
        output_path = Path(output_dir) / output_prefix
        
        success = generator.generate_simple_activity_diagram(c_file_path, str(output_path))
        
        if success:
            print("✅ Activity diagrams generated successfully!")
            return True
        else:
            print("❌ Failed to generate activity diagrams")
            return False
            
    except Exception as e:
        print(f"Error in activity diagram generation: {e}")
        return False


# Test function for standalone usage
def main():
    """Test function for standalone usage"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python activity_diagram.py <c_file_path> [output_prefix]")
        return
    
    c_file = sys.argv[1]
    output_prefix = sys.argv[2] if len(sys.argv) > 2 else "activity"
    output_dir = "."
    
    success = generate_activity_diagrams(c_file, output_dir, output_prefix)
    if success:
        print("Activity diagrams generated successfully!")
    else:
        print("Failed to generate activity diagrams")


if __name__ == "__main__":
    main()