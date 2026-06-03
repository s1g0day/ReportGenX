# AGENTS.md - ReportGenX Codebase Guide

> For AI coding agents operating in this Electron + FastAPI report generator.

## Project Overview

ReportGenX is a cross-platform desktop app for generating security vulnerability reports.
- **Frontend**: Electron + vanilla JavaScript (no framework)
- **Backend**: FastAPI (Python 3.9+) with SQLite database
- **Architecture**: Template-driven plugin system with dynamic form rendering

## Build & Run Commands

### Development

```bash
# Install dependencies
npm install                              # Frontend (Electron)
cd backend && pip install -r requirements.txt  # Backend (Python)

# Run development
cd backend && uvicorn api:app --host 127.0.0.1 --port 8000 --reload  # Backend
npm run start                            # Electron (in another terminal)
```

### Production Build

```bash
# Build Python backend executable
pyinstaller --noconfirm backend/api.spec

# Build Electron app
npm run dist                             # All platforms
npm run dist -- --win --x64              # Windows x64 only
npm run dist -- --mac --arm64            # macOS ARM only
```

### Version Sync

```bash
npm run sync-version    # Syncs version from package.json to backend/shared-config.json
```

## Project Structure

```
├── backend/                 # Python FastAPI backend
│   ├── api.py              # Main API entry point
│   ├── core/               # Core business logic
│   │   ├── base_handler.py # Template handler base class
│   │   ├── data_reader_db.py # SQLite database reader
│   │   ├── document_editor.py # Word document editor
│   │   ├── document_image_processor.py # Image processor
│   │   ├── exceptions.py   # Custom exceptions
│   │   ├── handler_utils.py # Handler utilities
│   │   ├── logger.py       # Logging system
│   │   ├── report_merger.py # Report merger
│   │   ├── summary_generator.py # Summary generator
│   │   └── template_manager.py # Template manager
│   ├── templates/          # Report template plugins
│   │   └── {template_id}/  # Each template is a folder
│   │       ├── schema.yaml # Form definition
│   │       ├── handler.py  # Business logic
│   │       └── template.docx
│   │   # Current templates: vuln_report, intrusion_report, penetration_test, Attack_Defense, single_vuln_report, intranet_vuln (内网测试报告)
│   ├── widgets/             # Shared widget directory (vuln_list.js + style.css)
│   └── data/               # SQLite database
├── src/                    # Frontend source
│   ├── index.html
│   ├── styles.css
│   └── js/
│       ├── api.js          # API client (window.AppAPI)
│       ├── config.js       # Configuration module
│       ├── form-renderer.js # Dynamic form generator
│       ├── main.js         # App entry point
│       ├── template-manager.js # Template management
│       ├── toolbox.js      # Toolbox utilities
│       ├── utils.js        # Common utilities
│       ├── vuln-manager.js # Vulnerability manager
│       └── managers/       # Feature modules
│           └── crud-manager.js # Generic CRUD abstraction
├── main.js                 # Electron main process
└── preload.js              # Electron preload script
```

## Core Module Boundary (Critical)

- `backend/core/`: **single implementation layer**. All real backend behavior lives here.
- `core/`: **SDK facade layer** for templates/plugins, mostly re-exporting `backend.core.*` for compatibility.

### Import Rules

- Backend runtime/internal code (`backend/api.py`, `backend/plugin_host/*`, backend tests): use `backend.core.*`.
- Template/plugin handlers may import `core.*` as a stable SDK entry.

### Maintenance Rule

When changing business logic, update `backend/core/*` only. Do not add independent implementations under `core/*`.

## Code Style Guidelines

### Python (Backend)

- **Imports**: stdlib -> third-party -> local, separated by blank lines
- **Type hints**: Required for function signatures
- **Naming**: snake_case for functions/variables, PascalCase for classes
- **Docstrings**: Google style with Args/Returns sections
- **Error handling**: Use HTTPException for API errors, log with `logger.error()`

```python
# Example pattern
from typing import Dict, Any, Optional
from fastapi import HTTPException

def process_data(data: Dict[str, Any], config: Optional[Dict] = None) -> Dict:
    """
    Process input data.
    
    Args:
        data: Input dictionary
        config: Optional configuration
        
    Returns:
        Processed result dictionary
    """
    if not data:
        raise HTTPException(status_code=400, detail="Empty data")
    return {"success": True, "result": data}
```

### JavaScript (Frontend)

- **No framework**: Vanilla JS with modular pattern
- **Global namespace**: Modules exposed via `window.AppXxx` (e.g., `window.AppAPI`)
- **Async/await**: Preferred over .then() chains
- **DOM queries**: Use `document.getElementById()` or `document.querySelector()`
- **Event handling**: Use addEventListener, support keyboard shortcuts

```javascript
// Module pattern example
window.AppMyModule = {
    init() {
        // Setup code
    },
    
    async fetchData() {
        try {
            const result = await AppAPI._request('/api/endpoint');
            return result;
        } catch (e) {
            console.error('Fetch failed:', e);
            throw e;
        }
    }
};
```

### Electron (main.js)

- **IPC**: Use ipcMain.handle() for async operations
- **Security**: contextIsolation enabled, sandbox disabled only for preload
- **Process management**: Clean up Python backend on app quit

## Template System

### Creating a New Template

1. Create folder: `backend/templates/{template_id}/`
2. Add `schema.yaml` - defines form fields and output config
3. Add `handler.py` - exports PLUGIN descriptor with `execute()` function
4. Add `template.docx` - Word template with `#placeholder#` markers

### Handler Pattern

Current standard: **PLUGIN descriptor + GenerationContext** (no class inheritance).

```python
from backend.core.generation_context import GenerationContext

PLUGIN = {"id": "my_template", "execute": None}

def execute(data, output_dir, template_manager, config, template_id):
    template_dir = template_manager.get_template_dir(template_id)
    info = template_manager.get_template_info(template_id)
    ctx = GenerationContext(template_dir, info, config, output_dir)

    doc = ctx.load_document()
    ctx.replace_text({
        "#title#": data.get("title", ""),
        "#content#": data.get("content", ""),
    })
    path = ctx.save_document(doc, ctx.build_output_path(
        data.get("unit_name", "unknown"), "report.docx"))
    ctx.postprocess(path, data, log_prefix="my_template")
    return {"success": True, "report_path": path, "message": "OK", "errors": []}

PLUGIN["execute"] = execute
```

> See `backend/templates/single_vuln_report/handler.py` (simplest) or `backend/templates/intranet_vuln/handler.py` (full-featured) for complete examples.

## API Conventions

### Response Format

```python
# Success
{"success": True, "message": "...", "data": {...}}

# Error
{"success": False, "error": "...", "detail": "..."}
```

### Endpoint Patterns

- `GET /api/templates` - List resources
- `GET /api/templates/{id}/schema` - Get single resource
- `POST /api/templates/{id}/generate` - Actions
- `PUT /api/vulnerabilities/{id}` - Update
- `DELETE /api/vulnerabilities/{id}` - Delete

## Key Dependencies

### Python
- fastapi, uvicorn - Web framework
- python-docx - Word document manipulation
- pydantic - Data validation
- pyyaml - YAML config parsing
- Pillow - Image processing

### Node.js
- electron - Desktop framework
- electron-builder - Packaging

## Common Pitfalls

1. **Path handling**: Use `os.path.join()`, handle Windows/Unix differences
2. **Encoding**: Always use `encoding='utf-8'` for file operations
3. **Process cleanup**: Backend process must be killed on app quit (see main.js)
4. **Template hot-reload**: New API routes require app restart
5. **Image validation**: Always verify image format before processing

## Testing Notes

- No automated test suite currently
- Manual testing via Electron UI
- Backend can be tested standalone: `uvicorn api:app --reload`

## CI/CD

GitHub Actions workflow (`.github/workflows/release.yml`):
- Triggers on `v*.*.*` tags
- Builds for Windows (x64/arm64), macOS (x64/arm64)
- Publishes to GitHub Releases
