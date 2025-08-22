# 🖥️ CLI Development Reference

*Comprehensive guide for building robust command line interfaces with multiple dependencies and commands.*

## 🏗️ **Architecture Patterns**

### **Modular Structure**
```
cli/
├── main.py              # Main entry point
├── commands/            # Command modules
│   ├── __init__.py     # Command registration
│   ├── process.py      # Processing commands
│   └── api.py          # API commands
├── config.py            # Configuration
└── utils/               # Utilities
    ├── formatters.py    # Output formatting
    └── validators.py    # Input validation
```

### **NAIB Project Structure Reorganization**
```
project_root/
├── launch.py                    # 🚀 Main application launcher
├── cli/                        # 🖥️ Command Line Interface
│   ├── main.py                 # Main CLI interface
│   ├── config.py               # CLI configuration
│   └── launchers/              # Various launcher scripts
├── process/                     # 📁 Processing domain
│   ├── social_media_pipeline.py # Moved from init.py
│   ├── data_processor.py       # Domain-specific modules
│   └── README.md               # Domain documentation
├── api_clients/                # 📁 API integration domain
│   ├── client1.py              # API client modules
│   └── README.md               # Domain documentation
├── config/                     # ⚙️ Configuration files
├── docs/                       # 📚 Documentation
└── requirements.txt            # 📦 Dependencies
```

### **Command Registration**
```python
# commands/__init__.py
COMMANDS = {}

def register_command(name: str, command_class):
    COMMANDS[name] = command_class

def get_command(name: str):
    return COMMANDS.get(name)

# Register commands
from .process import ProcessCommand
from .api import ApiCommand
register_command('process', ProcessCommand)
register_command('api', ApiCommand)
```

## 🚀 **Implementation**

### **Main Entry Point**
```python
# cli/main.py
#!/usr/bin/env python3
import sys
import argparse
import logging
from .commands import get_command, COMMANDS
from .config import CLIConfig

logger = logging.getLogger(__name__)

class CLI:
    def __init__(self):
        self.config = CLIConfig()
        self.parser = self._build_parser()
    
    def _build_parser(self):
        parser = argparse.ArgumentParser(description="Application CLI")
        parser.add_argument('--verbose', '-v', action='store_true')
        parser.add_argument('--debug', action='store_true', help='Enable debug mode')
        
        subparsers = parser.add_subparsers(dest='command')
        for cmd_name, cmd_class in COMMANDS.items():
            cmd_class.add_parser(subparsers)
        
        return parser
    
    def run(self, args=None):
        try:
            parsed_args = self.parser.parse_args(args)
            if not parsed_args.command:
                self.parser.print_help()
                return 1
            
            command_class = get_command(parsed_args.command)
            command = command_class(self.config, logger)
            return command.execute(parsed_args)
            
        except KeyboardInterrupt:
            logger.info("⏹️ Operation cancelled by user")
            return 130
        except Exception as e:
            logger.error(f"❌ Unexpected error: {e}")
            if parsed_args.debug:
                raise
            return 1

def main():
    cli = CLI()
    sys.exit(cli.run())
```

### **Base Command Class**
```python
# cli/commands/base.py
from abc import ABC, abstractmethod
import argparse
import logging

class BaseCommand(ABC):
    def __init__(self, config, logger):
        self.config = config
        self.logger = logger
    
    @classmethod
    @abstractmethod
    def add_parser(cls, subparsers):
        pass
    
    @abstractmethod
    def execute(self, args):
        pass
    
    def validate_args(self, args):
        return True
```

### **Command Implementation**
```python
# cli/commands/process.py
import argparse
import logging
from pathlib import Path
from .base import BaseCommand

logger = logging.getLogger(__name__)

class ProcessCommand(BaseCommand):
    @classmethod
    def add_parser(cls, subparsers):
        parser = subparsers.add_parser('process', help='Process data files')
        parser.add_argument('--input', '-i', required=True, help='Input file')
        parser.add_argument('--output', '-o', required=True, help='Output file')
        parser.add_argument('--format', choices=['json', 'csv'], default='json')
        parser.add_argument('--batch-size', type=int, default=1000)
    
    def validate_args(self, args):
        if not Path(args.input).exists():
            self.logger.error(f"❌ Input file not found: {args.input}")
            return False
        return True
    
    def execute(self, args):
        if not self.validate_args(args):
            return 1
        
        try:
            self.logger.info(f"🔍 Processing {args.input} -> {args.output}")
            # Your processing logic here
            self.logger.info("✅ Processing completed successfully")
            return 0
        except Exception as e:
            self.logger.error(f"❌ Processing error: {e}")
            return 1
```

## ⚙️ **Configuration & Utilities**

### **Configuration Management**
Create a `CLIConfig` class that:
- Loads JSON configuration files from common locations
- Provides default values for all settings
- Supports dot-notation access (e.g., `config.get('logging.level')`)
- Handles missing config files gracefully

### **Output Formatting**
```python
# cli/utils/formatters.py
import json
from rich.console import Console
from rich.table import Table

class OutputFormatter:
    def __init__(self, colorize=True):
        self.console = Console(color_system="auto" if colorize else None)
    
    def format_json(self, data, indent=2):
        return json.dumps(data, indent=indent, default=str)
    
    def format_table(self, data, headers):
        table = Table()
        for header in headers:
            table.add_column(header)
        for row in data:
            table.add_row(*[str(row.get(header, '')) for header in headers])
        return table
    
    def print_success(self, message):
        self.console.print(f"✅ {message}", style="green")
    
    def print_error(self, message):
        self.console.print(f"❌ {message}", style="red")
```

## 🔧 **Integration Patterns**

### **Module Discovery**
```python
# cli/utils/module_discovery.py
import importlib
import inspect
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class ModuleDiscoverer:
    def __init__(self, project_root):
        self.project_root = Path(project_root)
        self.discovered_modules = {}
    
    def discover_modules(self, module_paths):
        for module_path in module_paths:
            full_path = self.project_root / module_path
            if full_path.exists():
                if full_path.is_file():
                    self._discover_file_module(full_path)
                elif full_path.is_dir():
                    self._discover_directory_modules(full_path)
        return self.discovered_modules
    
    def _discover_file_module(self, file_path):
        try:
            module_name = file_path.stem
            spec = importlib.util.spec_from_file_location(module_name, file_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            self._analyze_module(module_name, module)
        except Exception as e:
            logger.warning(f"⚠️ Could not load {file_path}: {e}")
    
    def _analyze_module(self, module_name, module):
        for name, obj in inspect.getmembers(module):
            if (inspect.isfunction(obj) or inspect.isclass(obj)) and \
               (hasattr(obj, 'cli_help') or hasattr(obj, 'cli_args')):
                self.discovered_modules[f"{module_name}.{name}"] = obj
                logger.debug(f"🔍 Discovered CLI-compatible module: {module_name}.{name}")
```

### **Function Wrapper**
```python
# cli/commands/wrapper.py
import argparse
import logging
from .base import BaseCommand

logger = logging.getLogger(__name__)

class FunctionWrapperCommand(BaseCommand):
    def __init__(self, func, func_config):
        self.func = func
        self.func_config = func_config
    
    @classmethod
    def create_from_function(cls, func, config):
        class DynamicCommand(cls):
            @classmethod
            def add_parser(cls, subparsers):
                parser = subparsers.add_parser(
                    config.get('name', func.__name__),
                    help=config.get('help', func.__doc__)
                )
                
                import inspect
                sig = inspect.signature(func)
                for param_name, param in sig.parameters.items():
                    if param_name == 'self':
                        continue
                    
                    if param.default == param.empty:
                        parser.add_argument(f'--{param_name}', required=True)
                    else:
                        parser.add_argument(f'--{param_name}', default=param.default)
            
            def execute(self, args):
                try:
                    kwargs = {}
                    for param_name in inspect.signature(func).parameters:
                        if hasattr(args, param_name):
                            kwargs[param_name] = getattr(args, param_name)
                    
                    self.logger.info(f"🚀 Executing function: {func.__name__}")
                    result = self.func(**kwargs)
                    
                    if result:
                        self.logger.info("✅ Function executed successfully")
                        return 0
                    else:
                        self.logger.error("❌ Function execution failed")
                        return 1
                        
                except Exception as e:
                    self.logger.error(f"❌ Function execution error: {e}")
                    return 1
        
        return DynamicCommand
```

## 🏗️ **NAIB Project Structure Reorganization**

### **Mission & Purpose**
The NAIB (Project Structure Reorganization) approach transforms disorganized codebases into clean, logical, and maintainable structures following modern software engineering best practices.

### **Analysis Phase**

#### **1. Current Structure Assessment**
- Examine root directory for misplaced files (e.g., `init.py`, `main.py`)
- Identify logical groupings and dependencies between modules
- Note entry points and orchestrators
- Map the dependency hierarchy

#### **2. Architecture Pattern Recognition**
- Identify main application domains (data processing, API clients, integrations)
- Recognize orchestration patterns (pipelines, workflows, main functions)
- Understand dependency relationships
- Identify entry points and launchers

#### **3. Problem Identification**
- Find misnamed files that don't reflect their purpose
- Identify files that belong in specific domain folders
- Look for entry points that could be better organized
- Recognize separation of concerns opportunities

### **Reorganization Strategy**

#### **File Relocation Rules**
```python
# Example transformations
init.py → process/social_media_pipeline.py
main.py → cli/main.py (if CLI interface)
config.py → config/config.py (if configuration)
orchestrator.py → process/pipeline.py
```

#### **Naming Convention Updates**
- Use descriptive names indicating functionality
- Follow consistent patterns within each domain
- Update documentation to reflect new names
- Maintain backward compatibility where possible

#### **Entry Point Restructuring**
- Create main launcher (`launch.py`) in root directory
- Move old entry points to appropriate domain folders
- Ensure CLI can access all functionality
- Update import paths and references

### **Implementation Steps**

#### **Step 1: Create New Structure**
```bash
# Create domain folders
mkdir -p process api_clients config docs cli/launchers

# Move and rename files
mv init.py process/social_media_pipeline.py
mv main.py cli/main.py
mv config.py config/config.py
```

#### **Step 2: Update References**
```python
# Update import statements
# Before: from init import process_data
# After: from process.social_media_pipeline import process_data

# Update sys.path modifications if needed
import sys
sys.path.append(str(Path(__file__).parent / "process"))
```

#### **Step 3: Create New Entry Points**
```python
# launch.py - Main application launcher
#!/usr/bin/env python3
import sys
from pathlib import Path

# Add domain paths to Python path
project_root = Path(__file__).parent
sys.path.extend([
    str(project_root / "process"),
    str(project_root / "api_clients"),
    str(project_root / "config")
])

from cli.main import main

if __name__ == "__main__":
    sys.exit(main())
```

#### **Step 4: Update Documentation**
- Main README reflecting new structure
- Folder-specific READMEs explaining organization
- Usage instructions for new entry points
- Migration guide for existing users

### **CLI Integration with NAIB**

#### **Domain-Aware Command Discovery**
```python
# cli/utils/naib_discovery.py
import importlib
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class NAIBModuleDiscoverer:
    def __init__(self, project_root):
        self.project_root = Path(project_root)
        self.domain_modules = {}
    
    def discover_domain_modules(self):
        """Discover modules organized by NAIB domain structure"""
        domains = ['process', 'api_clients', 'config']
        
        for domain in domains:
            domain_path = self.project_root / domain
            if domain_path.exists():
                self.domain_modules[domain] = self._discover_domain(domain_path)
        
        return self.domain_modules
    
    def _discover_domain(self, domain_path):
        """Discover all modules within a specific domain"""
        modules = {}
        for py_file in domain_path.glob("*.py"):
            if py_file.name.startswith("__"):
                continue
            
            module_name = f"{domain_path.name}.{py_file.stem}"
            try:
                spec = importlib.util.spec_from_file_location(module_name, py_file)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                modules[py_file.stem] = module
                logger.debug(f"🔍 Discovered {module_name}")
            except Exception as e:
                logger.warning(f"⚠️ Could not load {py_file}: {e}")
        
        return modules
```

#### **NAIB-Aware CLI Commands**
```python
# cli/commands/naib.py
import argparse
import logging
from pathlib import Path
from .base import BaseCommand

logger = logging.getLogger(__name__)

class NAIBCommand(BaseCommand):
    @classmethod
    def add_parser(cls, subparsers):
        parser = subparsers.add_parser('naib', help='NAIB project structure management')
        subparsers_naib = parser.add_subparsers(dest='naib_action')
        
        # Analyze current structure
        analyze_parser = subparsers_naib.add_parser('analyze', help='Analyze current project structure')
        analyze_parser.add_argument('--output', '-o', help='Output analysis to file')
        
        # Reorganize structure
        reorganize_parser = subparsers_naib.add_parser('reorganize', help='Reorganize project structure')
        reorganize_parser.add_argument('--dry-run', action='store_true', help='Show changes without applying')
        reorganize_parser.add_argument('--backup', action='store_true', help='Create backup before changes')
        
        # Validate structure
        validate_parser = subparsers_naib.add_parser('validate', help='Validate current structure')
    
    def execute(self, args):
        if args.naib_action == 'analyze':
            return self._analyze_structure(args)
        elif args.naib_action == 'reorganize':
            return self._reorganize_structure(args)
        elif args.naib_action == 'validate':
            return self._validate_structure(args)
        else:
            self.logger.error("❌ No NAIB action specified")
            return 1
    
    def _analyze_structure(self, args):
        """Analyze current project structure and identify issues"""
        try:
            self.logger.info("🔍 Analyzing project structure...")
            
            # Implementation for structure analysis
            analysis_result = self._perform_structure_analysis()
            
            if args.output:
                self._save_analysis(analysis_result, args.output)
            else:
                self._display_analysis(analysis_result)
            
            return 0
        except Exception as e:
            self.logger.error(f"❌ Analysis failed: {e}")
            return 1
    
    def _reorganize_structure(self, args):
        """Reorganize project structure according to NAIB principles"""
        try:
            if args.backup:
                self.logger.info("💾 Creating backup...")
                self._create_backup()
            
            if args.dry_run:
                self.logger.info("🔍 Dry run - showing planned changes...")
                changes = self._plan_reorganization()
                self._display_planned_changes(changes)
            else:
                self.logger.info("🚀 Reorganizing project structure...")
                self._execute_reorganization()
                self.logger.info("✅ Reorganization completed successfully")
            
            return 0
        except Exception as e:
            self.logger.error(f"❌ Reorganization failed: {e}")
            return 1
    
    def _validate_structure(self, args):
        """Validate current structure against NAIB principles"""
        try:
            self.logger.info("✅ Validating project structure...")
            validation_result = self._validate_naib_structure()
            
            if validation_result['is_valid']:
                self.logger.info("✅ Structure is valid according to NAIB principles")
            else:
                self.logger.warning("⚠️ Structure has issues:")
                for issue in validation_result['issues']:
                    self.logger.warning(f"  - {issue}")
            
            return 0 if validation_result['is_valid'] else 1
        except Exception as e:
            self.logger.error(f"❌ Validation failed: {e}")
            return 1
```

### **NAIB Best Practices**

#### **✅ Do's**
- **Domain separation**: Group related functionality in domain-specific folders
- **Clear naming**: Use descriptive names that reflect actual purpose
- **Entry point consolidation**: Create single main launcher in root
- **Documentation**: Maintain README files in each domain folder
- **Backward compatibility**: Maintain import paths where possible
- **Incremental migration**: Reorganize in phases to avoid breaking changes

#### **❌ Don'ts**
- **Mixed concerns**: Don't mix different domains in single files
- **Unclear naming**: Avoid generic names like `init.py` or `main.py`
- **Root clutter**: Don't leave orchestration files in root directory
- **Broken imports**: Don't move files without updating all references
- **Big bang changes**: Avoid reorganizing everything at once

#### **🔧 NAIB Implementation Checklist**
- [ ] Analyze current project structure
- [ ] Identify logical domains and groupings
- [ ] Plan file relocations and renames
- [ ] Create new folder structure
- [ ] Move files to appropriate domains
- [ ] Update all import statements
- [ ] Create new entry points
- [ ] Update documentation
- [ ] Test all functionality
- [ ] Validate new structure

## 🧪 **Testing**

### **CLI Testing Framework**
```python
# tests/test_cli.py
import pytest
import logging
from unittest.mock import Mock, patch
from cli.main import CLI

# Configure logging for tests
logging.basicConfig(level=logging.DEBUG)

class TestCLI:
    def setup_method(self):
        self.cli = CLI()
    
    def test_help_output(self, capsys):
        with pytest.raises(SystemExit) as exc_info:
            self.cli.run(['--help'])
        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert 'Application CLI' in captured.out
    
    def test_unknown_command(self):
        result = self.cli.run(['unknown'])
        assert result == 1
    
    def test_command_execution(self):
        mock_command = Mock()
        mock_command.execute.return_value = 0
        
        with patch('cli.commands.get_command', return_value=Mock(return_value=mock_command)):
            result = self.cli.run(['test'])
            assert result == 0
            mock_command.execute.assert_called_once()
```

## 📚 **Best Practices**

### **✅ Do's**
- **Modular design**: Separate commands into individual modules
- **Consistent interface**: Use consistent argument patterns
- **Error handling**: Implement graceful error handling with logging
- **Configuration**: Use external configuration files
- **Documentation**: Provide comprehensive help and examples
- **Testing**: Include tests for CLI functionality

### **❌ Don'ts**
- **Monolithic commands**: Avoid putting all logic in one command
- **Hardcoded values**: Don't hardcode paths or settings
- **Poor error messages**: Avoid generic error messages
- **No validation**: Don't skip input validation
- **Inconsistent patterns**: Avoid different argument styles

### **🔧 Implementation Checklist**
- [ ] Create modular command structure
- [ ] Implement base command class with logging
- [ ] Add argument validation
- [ ] Include help text and examples
- [ ] Implement error handling with proper logging
- [ ] Create configuration management
- [ ] Add unit tests
- [ ] Document usage examples

## 🚀 **Quick Start Template**

```python
#!/usr/bin/env python3
import argparse
import sys
import logging

# Configure logging
logger = logging.getLogger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Your CLI Description")
    parser.add_argument('--input', '-i', required=True, help='Input file')
    parser.add_argument('--output', '-o', required=True, help='Output file')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    
    args = parser.parse_args()
    
    try:
        logger.info(f"🔍 Processing {args.input} -> {args.output}")
        # Your logic here
        logger.info("✅ Success!")
        return 0
    except Exception as e:
        logger.error(f"❌ Error: {e}")
        return 1

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    sys.exit(main())
```

---

**Remember**: Follow the project's logging standards with emojis and ensure all commands use proper logging instead of print statements.