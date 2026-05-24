# Tongshi Lesson Materials Code

This folder contains the runnable scripts for generating the course materials.

## Setup

```powershell
pip install -r requirements.txt
npm install
```

## Run

```powershell
python build_supporting_materials.py
python build_appendix_xlsx.py
npm run build:pptx
```

Generated files are written to `outputs/ai_complex_systems_materials/` inside this folder.
