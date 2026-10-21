# Umycs

Project for the Umycs data in SeaTable (product portfolios and related tables).

## SeaTable
- Access goes through the SeaTable MCP server configured in `.mcp.json`.
- Credentials come from environment variables `SEATABLE_SERVER_URL` and
  `SEATABLE_API_TOKEN`. Never commit the token or paste it into chat.
- Base name: <fill in>

## Data
- `bakery-portfolio.csv`: Bakery enzyme portfolio (17 products), import-ready for SeaTable.
  Columns: Product, Portfolio, Segment, Application, Active Components, Wave,
  Genetic Status, Physical Form, Pack Size (kg).

## Workflow
- Commit and push at the end of each session; pull before working on another computer.
