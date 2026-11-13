# Fit Check

Prototype for the NBAY 6170 capstone (Team 14): an AI agent that reads an Amazon dress page and its reviews and gives a fit call for the shopper's usual size.

- **Live page:** https://varunkulkarni0995-svg.github.io/fit-check/
- **Status:** v0.12. Runs live once a key is added in Settings (any provider and model per agent, with an automatic backup; Gemini recommended). Without a key it uses sample data.
- **Keys:** add them on the Settings page. They are saved only in your browser and are never stored in this repo.

## Agents
| Step | Recommended (main → backup) | Eval Nov 13 |
|---|---|---|
| Reader | Groq · openai/gpt-oss-20b → Gemini 3.5 Flash | 3/3 test pages, ~3 s |
| Web researcher | Serper search → Firecrawl read (Jina backup) → Decider model summary | tested in API Bench, Nov 13 |
| Analyst | Groq · openai/gpt-oss-120b → Gemini 3.5 Flash | 3/3 test pages, ~2 s |
| Decider | Groq · qwen/qwen3.8-27b → Gemini 3.5 Flash | 2/2 test pages, ~3 s (Gemini passed Almir's 10/10) |

One file, no build step: open `index.html` in a browser.

## Prompts
Every agent prompt can be edited on the Settings page (✎ Edit) without a rebuild. Edits are saved in that browser.

## Logs and KPIs
`apps-script/Code.gs` is the Google Sheet script. It logs every API request and response (files replaced by their size, keys never sent), each check, feedback and main button tap, and builds a KPIs tab.
