# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Windows automation project for laptop security inspection using Python. The application automates the process of collecting screenshots and system information for security compliance documentation.

## Architecture

The project consists of two main Python files:

- **main.py**: Main application with Tkinter GUI that coordinates the RPA workflow
- **docxrpa.py**: Document generation module using python-docx for creating inspection reports

### Key Components

1. **GUI Layer** (main.py:280-367)
   - RPA_GUI class provides the interface for starting automation
   - Overlay system to prevent user interaction during automation
   - Progress tracking and result reporting

2. **Automation Tasks** (main.py:223-279)
   - Sequential execution of inspection tasks
   - Screenshot capture of various system states
   - Integration with document generation

3. **Document Generation** (docxrpa.py)
   - Creates structured Word documents with 2x18 table format
   - Inserts screenshots into specific table cells
   - Handles image processing via PIL

## Common Commands

```bash
# Run the main application
python main.py

# Run the Flet-based UI (alternative interface)
python docxrpa.py

# Install dependencies
pip install -r requirements.txt
```

## Dependencies

The project requires Windows-specific automation libraries:
- pyautogui: Screen automation and screenshot capture
- pygetwindow: Window management
- pywinauto: Windows UI automation
- python-docx: Word document generation
- flet: Alternative GUI framework (in docxrpa.py)

## Important Implementation Details

### Window Automation Pattern
The project uses a consistent pattern for automating Windows applications:
1. Launch application via subprocess
2. Wait for window to appear
3. Capture window using pygetwindow
4. Take screenshot of specific region
5. Close window when done

### Screenshot Integration
Screenshots are captured as PIL Image objects and inserted directly into Word document cells using BytesIO streams (docxrpa.py:74-93).

### Task Execution Flow
Tasks are defined with:
- name: Display name for progress tracking
- action: Lambda function returning screenshot
- doc_row/doc_col: Target position in document table

## Security Inspection Items

The application checks 18 security items including:
- System information (OS version, installation date)
- MAC address
- Screen saver settings
- Antivirus status and logs
- Installed programs
- Real-time protection status

## Output

Generates Word documents named: `노트북_점검_보고서_YYYYMMDD_HHMMSS.docx`