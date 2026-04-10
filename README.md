CATIA Companion
Revision 1.0.0 | Released: 2026-04-10
 
A lightweight internal tool designed to streamline common CATIA V5 workflows,
including drawing and part conversion, font installation, and drafting standard setup.
 
Features:
  - Convert CATDrawing files to PDF
  - Convert CATPart/CATProduct files to STEP
  - Copy font files to the CATIA installation folder
  - Copy ISO.xml drafting standard to the CATIA installation folder

Requirements:
  - Python 3.10 or newer
  - Windows (the tool targets CATIA V5 via COM automation and reads the Windows Registry)
  - A running CATIA V5 instance is required for file-conversion features

Installation / Development Setup:
  1. Clone the repository:
       git clone https://github.com/RayDutchman/CATIA-Companion.git
       cd CATIA-Companion
  2. Create and activate a virtual environment:
       python -m venv .venv
       .venv\Scripts\activate
  3. Install runtime dependencies:
       pip install -r requirements.txt

Building (Windows .exe):
  Prerequisites: pip install pyinstaller
  Build command:  pyinstaller build.spec
  Output:         dist\CATIA Companion\CATIA Companion.exe
  The ISO.xml and ChangFangSong.ttf resource files are automatically placed
  next to the executable by the spec file.
 
Developed by: CHEN Weibo
Contact: thucwb@gmail.com
 
For internal use only. Not for redistribution.
