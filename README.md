# Fit Check

Prototype for the NBAY 6170 capstone (Team 14): an AI agent that reads a clothing product page (Amazon or any store) and its reviews and gives a fit call for the shopper's usual size.

- **Live page:** https://varunkulkarni0995-svg.github.io/fit-check/
- **Status:** v0.16.1. Long uploads are read in parts (about 7,000 characters or 3 page images each, split at review boundaries) so no answer gets cut off; Runs live once a key is added in Settings (any provider and model per agent, with an automatic backup; Gemini recommended). Without a key it uses sample data.
- **Keys:** add them on the Settings page. They are saved only in your browser and are never stored in this repo.

## Agents
| Step | Recommended (main → backup) | Eval Nov 13 |
|---|---|---|
| Reader | Groq · openai/gpt-oss-20b → Gemini 3.5 Flash | 3/3 test pages, ~3 s |
| Web researcher | Serper search → Firecrawl read (Jina backup) → Decider model summary | tested in API Bench, Nov 13 |
| Review collector | Team memory (Sheet) → Serper search (review pages ranked first) → Firecrawl with scroll + "load more" clicks → schema.org review data read directly (no AI), else Groq · openai/gpt-oss-120b with a short prompt → Gemini backup | v0.15: max 3 pages, stops at 20 reviews, skips pages with no reviews before any AI call |
| Analyst | Groq · openai/gpt-oss-120b → Gemini 3.5 Flash | v0.14 new prompt: tags reviews, the app counts (old evals need a rerun) |
| Decider | Groq · qwen/qwen3.8-27b → Gemini 3.5 Flash | 2/2 test pages, ~3 s (Gemini passed Almir's 10/10) |

One file, no build step: open `index.html` in a browser.

## How matching works (v0.14)
- The **Analyst tags** each review (how it runs, body words, problems with exact quotes, comparisons to other brands, size mapping). The **app counts** from the tags, so counts are the same every time.
- **Built-in tolerance:** the item's fit tolerance sets who counts as "shoppers like you". High: a size either side and height within 3 in. Medium: same size, height within 2 in. Low: same size only. Editable in Settings → Matching rules.
- **Nearby sizes:** a review one size up that says "tight" counts for your size; one size down that says "loose" counts too.
- **Brands you wear:** reviews saying "fits like Uniqlo" count as shoppers like you when Uniqlo is in your profile.
- **Reviews disagree:** when both "tight" and "loose" are reported for the same area, the result shows both sides and the call becomes Check first.
- **Shoppers-like-you filter:** off by default for now (every review counts); turn on in Settings → Matching rules.
- **Fallback:** if fewer than the minimum (default 10) match you, the call uses all shoppers and says so.
- **Team memory:** every check saves its reviews to the Sheet's Reviews tab; later checks of the same product reuse them, and other products from the brand give a brand-level sizing line.

## Prompts
Every agent prompt can be edited on the Settings page (✎ Edit) without a rebuild. Edits are saved in that browser.

## Logs and KPIs
`apps-script/Code.gs` is the Google Sheet script. It logs every API request and response (files replaced by their size, keys never sent), each check, feedback and main button tap, and builds a KPIs tab. It also stores team memory (Reviews tab). After changing the script: Deploy → Manage deployments → Edit → New version → Deploy (the URL stays the same).
