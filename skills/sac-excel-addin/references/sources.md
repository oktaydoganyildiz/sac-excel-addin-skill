# Sources

All sources below were checked on **2026-10-29**. At that point the add-in was at version **2026.23**, released 2026-10-26. The add-in ships about monthly, so re-check What's New before relying on the version-specific statements in this skill.

## Official SAP documentation [H]

Product: *SAP Analytics Cloud, add-in for Microsoft Excel*, https://help.sap.com/docs/SAP_ANALYTICS_CLOUD_OFFICE

| Deliverable | ID | Used for |
|---|---|---|
| What's New | `ccbf41c7f6ff49caa7af5b4b8c2312e7` | Version history 2025.03–2026.23 |
| Onboarding | `2b73edbc8f8f4189a36bc7a2e038185c` | System requirements, tenant configuration, deployment, permissions |
| Analyzing Data | `c637c9ff5d61457eb415ce161e38e57b` | Tables, filters, hierarchies, source limits, preferences, shortcuts |
| Planning | `9bafb6f9d4034781b8040fc29a2cd99f` | Planning, data/multi action triggers |
| Analyze and Plan with Formulas | `29624d6020b74bf8817bb7df398e5cb6` | All 16 `SAP.*` functions |
| Security for the add-in | `0ef628d47e504522819bf3ba934f6d3f` | Security, data protection |
| Troubleshooting & FAQ | `4554bec9b964401db071a89f779e8319` | FAQ, SAP support |

Analysis for Microsoft Office, edition for SAP Analytics Cloud 2.8 SP31: https://help.sap.com/docs/SAP_ANALYSIS_OFFICE_CLOUD (user guide `b25c777a2c8c4e3da2f472cce24f289c`). Used for the AO function list, `SAPGetData` syntax and AO VBA API method names.

SAP Analytics Cloud user guide, *Microsoft Office Integration*: https://help.sap.com/docs/SAP_ANALYTICS_CLOUD/00f68c2e08b941f081002fd3691d86a7/31601f28442a40978d770783820aa234.html

## SAP Notes [N]

- **3594453**: *XLA file to support VBA with SAP Analytics Cloud, add-in for Microsoft Excel* (v3, 2026-07-08). Attachment `VbaApiDocumentation.pdf`. Requires an S-user. This skill lists only method names and short descriptions; read the note for the full reference.
- **3323835**: performance of `SAP.GETDATA` formulas (referenced from SAP Help, not read directly).
- **2770570**: MDS result-record limit (referenced from SAP Help, not read directly).
- **3652483**: preview version of the add-in (referenced from What's New, not read directly).

## SAP-employee blog [B]

- *Comparison of SAP Analysis for Microsoft Office (SAC edition) vs. SAC add-in for Microsoft Excel*, technology-blog-posts-by-sap, published 2024-08-22, updated July 2025: https://community.sap.com/t5/technology-blog-posts-by-sap/comparison-of-sap-analysis-for-microsoft-office-sac-edition-vs-sac-add-in/ba-p/13783155. Used for product positioning, the data-source matrix, the AO maintenance outlook and feature gaps.

## SAP Community Q&A [Q]

- *SAC VBA add-in – getting stuck on Sign in and Sign out function; macro freezing*, 2025-09-04: https://community.sap.com/t5/technology-q-a/sac-vba-add-in-getting-stuck-on-sign-in-and-sign-out-function-macro/qaq-p/14206609
- *SAC Excel add-in VBA macro issues* (BW AO → SAC add-in migration blocked by VBA), 2026-08-24: https://community.sap.com/t5/technology-q-a/sac-excel-add-in-vba-macro-issues/qaq-p/14470312

## Not verified

- Live tenant tests: none were run for this release of the skill. Statements come from documentation, not from hands-on testing.
- Items marked *(unverified)* in the reference files.
