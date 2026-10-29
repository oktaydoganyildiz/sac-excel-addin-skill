# SAC Excel add-in skill

An AI agent skill for **SAP Analytics Cloud, add-in for Microsoft Excel**. It helps coding and chat agents (Claude Code, Codex, Cursor, Joule and others that read `SKILL.md`) with:

- **`SAP.*` formulas.** Write, convert and debug `SAP.GETDATA`, `SAP.SETDATA`, `SAP.GETDIMENSIONFILTER`, `SAP.OVERWRITEDIMENSIONFILTER` and the other 12 functions.
- **Analysis for Office migration.** Check data source fit and feature gaps, map `SAPGetData` → `SAP.GETDATA`, and decide what happens to AO VBA macros.
- **Planning in Excel.** Mass data entry, versions, publish/revert, data action and multi action triggers.
- **VBA automation.** Use the `SAPOFXLA.xlam` bridge from SAP Note 3594453, including what it cannot do.
- **Setup and troubleshooting.** Non-standard tenant domains, trusted origins, permissions, memory limits and high-volume mode.

Every statement was checked against SAP Help and SAP Notes on 2026-10-29 (add-in version 2026.23). Sources are listed in [`skills/sac-excel-addin/references/sources.md`](skills/sac-excel-addin/references/sources.md).

## Install

```bash
npx skills add oktaydoganyildiz/sac-excel-addin-skill --skill sac-excel-addin
```

Or copy `skills/sac-excel-addin/` into your agent's skills folder, for example `~/.claude/skills/`.

## Layout

```
skills/sac-excel-addin/
├── SKILL.md                     core rules, formula map, safety
├── references/
│   ├── formulas.md              all 16 SAP.* functions with arguments
│   ├── migration-from-analysis-office.md
│   ├── vba-api.md
│   ├── planning.md
│   ├── analysis-tables.md
│   ├── setup-and-troubleshooting.md
│   └── sources.md
└── evals/evals.json             test prompts with expected and forbidden content
```

## Safety

The skill contains no scripts, credentials or network calls; it is documentation only. It tells the agent to:
- state what will change and get the user's confirmation before suggesting anything that writes planning data, runs data or multi actions, or publishes versions;
- leave Microsoft 365 and SAC admin changes to an administrator;
- never put credentials into workbooks or macros.

## Author

**Oktay Doganyildiz**, SAP BW/4HANA, Datasphere and Business Data Cloud consultant.
[LinkedIn](https://www.linkedin.com/in/oktaydoganyildiz/) · [GitHub](https://github.com/oktaydoganyildiz)

Issues and corrections are welcome via GitHub issues. SAP releases the add-in about monthly, so check What's New if a statement looks outdated.

## Disclaimer

This is a community project. It is not affiliated with, endorsed by or supported by SAP SE. SAP, SAP Analytics Cloud and Analysis for Office are trademarks of SAP SE. Microsoft and Excel are trademarks of Microsoft Corporation.

## License

[Apache License 2.0](LICENSE). Copyright 2026 Oktay Doganyildiz.
