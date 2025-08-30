# Auto PPT Generator

Turn bulk text/markdown into a PowerPoint that follows your uploaded template’s look & feel. Optional LLM (OpenAI/Anthropic/Gemini) for smarter slide mapping. No AI images.

## Features

- Paste text + optional guidance
- Upload `.pptx/.potx` template
- Optional LLM provider + API key (never stored) to produce a JSON slide outline
- Heuristic fallback when no LLM used
- Reuses template styling via slide layouts and reuses template images (if present)

## Quick Start

```bash
pip install -r requirements.txt
uvicorn app:app --reload
```
