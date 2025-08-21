# 🖥️ Command Line Interface (CLI) Development Reference

*This document provides a comprehensive guide for building robust, maintainable command line interfaces for projects with multiple dependencies and commands. Use this as a reference when implementing CLI functionality.*

## 🎯 **Overview**

A well-designed CLI provides users with an intuitive, powerful interface to interact with your application's functionality. This reference covers best practices for creating CLIs that are maintainable, extensible, and user-friendly.

## 🏗️ **CLI Architecture Patterns**

### **1. Modular Command Structure**
```
cli/
├── main.py              # Main CLI entry point
├── commands/            # Individual command modules
│   ├── __init__.py     # Command registration
│   ├── process.py      # Processing commands
│   ├── api.py          # API-related commands
│   └── utils.py        # Utility commands
├── config.py            # CLI configuration
├── launchers/           # Specialized launchers
│   ├── dev.py          # Development launcher
│   └── production.py   # Production launcher
└── utils/               # CLI utilities
    ├── formatters.py    # Output formatting
    ├── validators.py    # Input validation
    └── helpers.py       # Common helpers
```

### **2. Command Registration Pattern**
```python
# commands/__init__.py
from typing import Dict, Type
from .base import BaseCommand

COMMANDS: Dict[str, Type[BaseCommand]] = {}

def register_command(name: str, command_class: Type[BaseCommand]):
    """Register a command with the CLI system."""
    COMMANDS[name] = command_class

def get_command(name: str) -> Type[BaseCommand]:
    """Retrieve a registered command."""
    return COMMANDS.get(name)

# Register all commands
from .process import ProcessCommand
from .api import ApiCommand
from .utils import UtilsCommand

register_command('process', ProcessCommand)
register_command('api', ApiCommand)
register_command('utils', UtilsCommand)
```

## 🚀 **Implementation Guidelines**

### **1. Entry Point Structure**
```python
# cli/main.py
#!/usr/bin/env python3
"""
Main CLI entry point for the application.
Handles command parsing, routing, and execution.
"""

import sys
import argparse
from typing import List, Optional
from .commands import get_command, COMMANDS
from .config import CLIConfig
from .utils.logging import setup_logging

class CLI:
    def __init__(self):
        self.config = CLIConfig()
        self.parser = self._build_parser()
        self.logger = setup_logging()
    
    def _build_parser(self) -> argparse.ArgumentParser:
        """Build the main argument parser."""
        parser = argparse.ArgumentParser(
            description="Application CLI - Manage and execute various operations",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
  %(prog)s process --input data.csv --output results.json
  %(prog)s api --endpoint users --method GET
  %(prog)s utils --help
            """
        )
        
        # Add global options
        parser.add_argument(
            '--verbose', '-v',
            action='store_true',
            help='Enable verbose output'
        )
        
        parser.add_argument(
            '--config',
            type=str,
            help='Path to configuration file'
        )
        
        # Add subcommands
        subparsers = parser.add_subparsers(
            dest='command',
            help='Available commands'
        )
        
        # Register all commands
        for cmd_name, cmd_class in COMMANDS.items():
            cmd_class.add_parser(subparsers)
        
        return parser
    
    def run(self, args: Optional[List[str]] = None) -> int:
        """Run the CLI with given arguments."""
        try:
            parsed_args = self.parser.parse_args(args)
            
            if not parsed_args.command:
                self.parser.print_help()
                return 1
            
            # Get and execute command
            command_class = get_command(parsed_args.command)
            if not command_class:
                self.logger.error(f"Unknown command: {parsed_args.command}")
                return 1
            
            command = command_class(self.config, self.logger)
            return command.execute(parsed_args)
            
        except KeyboardInterrupt:
            self.logger.info("Operation cancelled by user")
            return 130
        except Exception as e:
            self.logger.error(f"Unexpected error: {e}")
            if self.config.debug:
                raise
            return 1

def main():
    """Main entry point."""
    cli = CLI()
    sys.exit(cli.run())

if __name__ == '__main__':
    main()
```

### **2. Base Command Class**
```python
# cli/commands/base.py
from abc import ABC, abstractmethod
import argparse
from typing import Any, Dict
from ..config import CLIConfig
from ..utils.logging import Logger

class BaseCommand(ABC):
    """Base class for all CLI commands."""
    
    def __init__(self, config: CLIConfig, logger: Logger):
        self.config = config
        self.logger = logger
    
    @classmethod
    @abstractmethod
    def add_parser(cls, subparsers: argparse._SubParsersAction) -> None:
        """Add command-specific arguments to the parser."""
        pass
    
    @abstractmethod
    def execute(self, args: argparse.Namespace) -> int:
        """Execute the command logic."""
        pass
    
    def validate_args(self, args: argparse.Namespace) -> bool:
        """Validate command arguments."""
        return True
    
    def setup_environment(self, args: argparse.Namespace) -> None:
        """Setup environment for command execution."""
        pass
    
    def cleanup(self) -> None:
        """Cleanup resources after command execution."""
        pass
```

### **3. Command Implementation Example**
```python
# cli/commands/process.py
import argparse
from pathlib import Path
from typing import Dict, Any
from .base import BaseCommand

class ProcessCommand(BaseCommand):
    """Process data files with various operations."""
    
    @classmethod
    def add_parser(cls, subparsers: argparse._SubParsersAction) -> None:
        parser = subparsers.add_parser(
            'process',
            help='Process data files',
            description='Process data files with various operations like filtering, transformation, and aggregation.'
        )
        
        # Input/Output options
        parser.add_argument(
            '--input', '-i',
            type=str,
            required=True,
            help='Input file or directory path'
        )
        
        parser.add_argument(
            '--output', '-o',
            type=str,
            required=True,
            help='Output file or directory path'
        )
        
        # Processing options
        parser.add_argument(
            '--format',
            choices=['json', 'csv', 'xml', 'yaml'],
            default='json',
            help='Output format (default: json)'
        )
        
        parser.add_argument(
            '--filter',
            type=str,
            help='Filter expression for data processing'
        )
        
        # Performance options
        parser.add_argument(
            '--batch-size',
            type=int,
            default=1000,
            help='Batch size for processing (default: 1000)'
        )
        
        parser.add_argument(
            '--workers',
            type=int,
            default=1,
            help='Number of worker processes (default: 1)'
        )
    
    def validate_args(self, args: argparse.Namespace) -> bool:
        """Validate command arguments."""
        input_path = Path(args.input)
        if not input_path.exists():
            self.logger.error(f"Input path does not exist: {args.input}")
            return False
        
        output_path = Path(args.output)
        output_dir = output_path.parent
        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.logger.error(f"Cannot create output directory: {e}")
                return False
        
        if args.batch_size <= 0:
            self.logger.error("Batch size must be positive")
            return False
        
        if args.workers <= 0:
            self.logger.error("Number of workers must be positive")
            return False
        
        return True
    
    def execute(self, args: argparse.Namespace) -> int:
        """Execute the process command."""
        try:
            if not self.validate_args(args):
                return 1
            
            self.logger.info(f"Starting data processing...")
            self.logger.info(f"Input: {args.input}")
            self.logger.info(f"Output: {args.output}")
            self.logger.info(f"Format: {args.format}")
            
            # Setup environment
            self.setup_environment(args)
            
            # Execute processing logic
            result = self._process_data(args)
            
            if result:
                self.logger.info("✅ Data processing completed successfully")
                return 0
            else:
                self.logger.error("❌ Data processing failed")
                return 1
                
        except Exception as e:
            self.logger.error(f"Processing error: {e}")
            return 1
        finally:
            self.cleanup()
    
    def _process_data(self, args: argparse.Namespace) -> bool:
        """Execute the actual data processing logic."""
        # Implementation specific to your use case
        # This is where you'd integrate with your existing modules
        pass
```

## ⚙️ **Configuration Management**

### **1. CLI Configuration Class**
```python
# cli/config.py
import json
import os
from pathlib import Path
from typing import Any, Dict, Optional

class CLIConfig:
    """Configuration management for CLI operations."""
    
    def __init__(self, config_path: Optional[str] = None):
        self.config_path = config_path or self._find_config()
        self.config = self._load_config()
        self.debug = os.getenv('DEBUG', 'false').lower() == 'true'
    
    def _find_config(self) -> Optional[str]:
        """Find configuration file in common locations."""
        search_paths = [
            'cli_config.json',
            'config/cli_config.json',
            os.path.expanduser('~/.config/app/cli_config.json'),
            '/etc/app/cli_config.json'
        ]
        
        for path in search_paths:
            if Path(path).exists():
                return path
        
        return None
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file."""
        if not self.config_path:
            return self._get_default_config()
        
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: Could not load config from {self.config_path}: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration values."""
        return {
            'logging': {
                'level': 'INFO',
                'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            },
            'output': {
                'default_format': 'json',
                'indent': 2,
                'colorize': True
            },
            'processing': {
                'default_batch_size': 1000,
                'default_workers': 1,
                'timeout': 300
            }
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key."""
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
```

## 🎨 **Output Formatting & User Experience**

### **1. Output Formatters**
```python
# cli/utils/formatters.py
import json
import yaml
from typing import Any, Dict, List
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

class OutputFormatter:
    """Format CLI output in various styles."""
    
    def __init__(self, colorize: bool = True):
        self.console = Console(color_system="auto" if colorize else None)
    
    def format_json(self, data: Any, indent: int = 2) -> str:
        """Format data as JSON."""
        return json.dumps(data, indent=indent, default=str)
    
    def format_yaml(self, data: Any) -> str:
        """Format data as YAML."""
        return yaml.dump(data, default_flow_style=False, sort_keys=False)
    
    def format_table(self, data: List[Dict[str, Any]], headers: List[str]) -> str:
        """Format data as a table."""
        table = Table()
        
        for header in headers:
            table.add_column(header)
        
        for row in data:
            table.add_row(*[str(row.get(header, '')) for header in headers])
        
        return table
    
    def print_success(self, message: str) -> None:
        """Print success message."""
        self.console.print(f"✅ {message}", style="green")
    
    def print_error(self, message: str) -> None:
        """Print error message."""
        self.console.print(f"❌ {message}", style="red")
    
    def print_warning(self, message: str) -> None:
        """Print warning message."""
        self.console.print(f"⚠️  {message}", style="yellow")
    
    def print_info(self, message: str) -> None:
        """Print info message."""
        self.console.print(f"ℹ️  {message}", style="blue")
```

### **2. Progress Indicators**
```python
# cli/utils/progress.py
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
from rich.console import Console
from typing import Optional

class ProgressManager:
    """Manage progress indicators for long-running operations."""
    
    def __init__(self):
        self.console = Console()
        self.progress = None
    
    def start_progress(self, description: str, total: Optional[int] = None):
        """Start a progress indicator."""
        columns = [
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn()
        ]
        
        self.progress = Progress(*columns, console=self.console)
        self.progress.start()
        
        if total:
            self.task = self.progress.add_task(description, total=total)
        else:
            self.task = self.progress.add_task(description, total=None)
    
    def update_progress(self, advance: int = 1):
        """Update progress."""
        if self.progress and hasattr(self, 'task'):
            self.progress.advance(self.task, advance)
    
    def stop_progress(self):
        """Stop progress indicator."""
        if self.progress:
            self.progress.stop()
```

## 🔧 **Integration with Existing Modules**

### **1. Module Discovery & Registration**
```python
# cli/utils/module_discovery.py
import importlib
import inspect
from pathlib import Path
from typing import Dict, Type, Any
from .base import BaseCommand

class ModuleDiscoverer:
    """Discover and register available modules for CLI integration."""
    
    def __init__(self, project_root: str):
        self.project_root = Path(project_root)
        self.discovered_modules = {}
    
    def discover_modules(self, module_paths: List[str]) -> Dict[str, Any]:
        """Discover modules in specified paths."""
        for module_path in module_paths:
            full_path = self.project_root / module_path
            
            if full_path.exists():
                if full_path.is_file():
                    self._discover_file_module(full_path)
                elif full_path.is_dir():
                    self._discover_directory_modules(full_path)
        
        return self.discovered_modules
    
    def _discover_file_module(self, file_path: Path) -> None:
        """Discover a single file module."""
        try:
            module_name = file_path.stem
            spec = importlib.util.spec_from_file_location(module_name, file_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            # Look for CLI-compatible functions/classes
            self._analyze_module(module_name, module)
            
        except Exception as e:
            print(f"Warning: Could not load module {file_path}: {e}")
    
    def _discover_directory_modules(self, dir_path: Path) -> None:
        """Discover modules in a directory."""
        for item in dir_path.iterdir():
            if item.is_file() and item.suffix == '.py':
                self._discover_file_module(item)
    
    def _analyze_module(self, module_name: str, module: Any) -> None:
        """Analyze a module for CLI-compatible components."""
        for name, obj in inspect.getmembers(module):
            if inspect.isfunction(obj) or inspect.isclass(obj):
                # Check if it has CLI-compatible attributes
                if hasattr(obj, 'cli_help') or hasattr(obj, 'cli_args'):
                    self.discovered_modules[f"{module_name}.{name}"] = obj
```

### **2. Command Wrapper for Existing Functions**
```python
# cli/commands/wrapper.py
import argparse
from typing import Any, Callable, Dict
from .base import BaseCommand

class FunctionWrapperCommand(BaseCommand):
    """Wrap existing functions as CLI commands."""
    
    def __init__(self, config, logger, func: Callable, func_config: Dict[str, Any]):
        super().__init__(config, logger)
        self.func = func
        self.func_config = func_config
    
    @classmethod
    def create_from_function(cls, func: Callable, config: Dict[str, Any]) -> Type['FunctionWrapperCommand']:
        """Create a command class from an existing function."""
        
        class DynamicCommand(cls):
            @classmethod
            def add_parser(cls, subparsers: argparse._SubParsersAction) -> None:
                parser = subparsers.add_parser(
                    config.get('name', func.__name__),
                    help=config.get('help', func.__doc__ or 'No description available'),
                    description=config.get('description', func.__doc__ or 'No description available')
                )
                
                # Add arguments based on function signature
                import inspect
                sig = inspect.signature(func)
                
                for param_name, param in sig.parameters.items():
                    if param_name == 'self':
                        continue
                    
                    if param.default == param.empty:
                        parser.add_argument(
                            f'--{param_name}',
                            required=True,
                            help=f'{param_name} parameter'
                        )
                    else:
                        parser.add_argument(
                            f'--{param_name}',
                            default=param.default,
                            help=f'{param_name} parameter (default: {param.default})'
                        )
            
            def execute(self, args: argparse.Namespace) -> int:
                """Execute the wrapped function."""
                try:
                    # Extract arguments
                    kwargs = {}
                    for param_name in inspect.signature(func).parameters:
                        if hasattr(args, param_name):
                            kwargs[param_name] = getattr(args, param_name)
                    
                    # Execute function
                    result = self.func(**kwargs)
                    
                    if result:
                        self.logger.info("✅ Function executed successfully")
                        return 0
                    else:
                        self.logger.error("❌ Function execution failed")
                        return 1
                        
                except Exception as e:
                    self.logger.error(f"Function execution error: {e}")
                    return 1
        
        return DynamicCommand
```

## 🧪 **Testing & Validation**

### **1. CLI Testing Framework**
```python
# tests/test_cli.py
import pytest
from unittest.mock import Mock, patch
from cli.main import CLI
from cli.config import CLIConfig

class TestCLI:
    """Test CLI functionality."""
    
    def setup_method(self):
        """Setup test environment."""
        self.config = Mock(spec=CLIConfig)
        self.cli = CLI()
    
    def test_help_output(self, capsys):
        """Test help output generation."""
        with pytest.raises(SystemExit) as exc_info:
            self.cli.run(['--help'])
        
        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert 'Application CLI' in captured.out
    
    def test_unknown_command(self):
        """Test handling of unknown commands."""
        result = self.cli.run(['unknown'])
        assert result == 1
    
    def test_command_execution(self):
        """Test successful command execution."""
        # Mock a command
        mock_command = Mock()
        mock_command.execute.return_value = 0
        
        with patch('cli.commands.get_command', return_value=Mock(return_value=mock_command)):
            result = self.cli.run(['test'])
            assert result == 0
            mock_command.execute.assert_called_once()
```

## 📚 **Best Practices Summary**

### **✅ Do's**
- **Modular design**: Separate commands into individual modules
- **Consistent interface**: Use consistent argument patterns across commands
- **Error handling**: Implement graceful error handling with clear messages
- **Configuration**: Use external configuration files for flexibility
- **Documentation**: Provide comprehensive help and examples
- **Testing**: Include tests for CLI functionality
- **Progress feedback**: Show progress for long-running operations

### **❌ Don'ts**
- **Monolithic commands**: Avoid putting all logic in a single command
- **Hardcoded values**: Don't hardcode paths, settings, or limits
- **Poor error messages**: Avoid generic or unhelpful error messages
- **No validation**: Don't skip input validation
- **Inconsistent patterns**: Avoid different argument styles between commands
- **No help**: Don't forget to include help text and examples

### **🔧 Implementation Checklist**
- [ ] Create modular command structure
- [ ] Implement base command class
- [ ] Add argument validation
- [ ] Include comprehensive help text
- [ ] Implement error handling
- [ ] Add progress indicators
- [ ] Create configuration management
- [ ] Include logging and debugging
- [ ] Add unit tests
- [ ] Document usage examples
- [ ] Consider backward compatibility
- [ ] Implement command discovery

## 🚀 **Quick Start Template**

```python
#!/usr/bin/env python3
"""
Quick CLI template - replace with your specific functionality.
"""

import argparse
import sys
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(description="Your CLI Description")
    parser.add_argument('--input', '-i', required=True, help='Input file')
    parser.add_argument('--output', '-o', required=True, help='Output file')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    
    args = parser.parse_args()
    
    try:
        # Your logic here
        print(f"Processing {args.input} -> {args.output}")
        # ... implementation ...
        print("✅ Success!")
        return 0
    except Exception as e:
        print(f"❌ Error: {e}")
        return 1

if __name__ == '__main__':
    sys.exit(main())
```

---

**Remember**: A well-designed CLI should be intuitive, powerful, and maintainable. Focus on user experience, error handling, and extensibility from the start.