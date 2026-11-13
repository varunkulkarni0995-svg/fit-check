# Fit Check

Prototype for the NBAY 6170 capstone (Team 14): an AI agent that reads an Amazon dress page and its reviews and gives a fit call for the shopper's usual size.

- **Live page:** https://varunkulkarni0995-svg.github.io/fit-check/
- **Status:** v0.7. Runs live once a key is added in Settings (Gemini recommended for every step). Without a key it uses sample data.
- **Keys:** add them on the Settings page. They are saved only in your browser and are never stored in this repo.

## Agents
| Step | Recommended | Status |
|---|---|---|
| Reader | Gemini 3.5 Flash | not tested yet |
| Web researcher | Gemini 3.5 Flash + Google Search (can be switched off) | not tested yet |
| Analyst | Gemini 3.5 Flash | not tested yet |
| Decider | Gemini 3.5 Flash (backup: Nemotron 3 Ultra on OpenRouter) | tested 10/10, Nov 12 |

One file, no build step: open `index.html` in a browser.

## Prompts
Every agent prompt can be edited on the Settings page (✎ Edit) without a rebuild. Edits are saved in that browser.

## Logs and KPIs
 is the Google Sheet script. It logs every API request and response (files replaced by their size, keys never sent), each check, feedback and main button tap, and builds a KPIs tab.
