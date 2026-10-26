# Umycs

Project for the Umycs data in SeaTable (product portfolios and related tables).

## SeaTable
- Access goes through the SeaTable MCP server configured in `.mcp.json`.
- Credentials come from environment variables `SEATABLE_SERVER_URL` and
  `SEATABLE_API_TOKEN`. Never commit the token or paste it into chat.
- Base name: PDM

## Workflow
- Commit and push at the end of each session; pull before working on another computer.

## TPDS translations
- `scripts/Generate_TPDS_i18n.py` generates one TPDS PDF per language in `TPDS view`.`Language` (EN = source, no translation).
- Template labels come from table `Labels` (Key, EN, DE, ...); other texts come from table `Translation memory`
  (Source, Field, one column per language, `Reviewed languages`). Unknown texts are appended to the memory by the script.
- The script in the repo holds placeholders for the DocuGenerate key and Graph client secret; never commit the real ones.
