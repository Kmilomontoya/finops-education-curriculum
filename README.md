# FinOps Education Curriculum Generator

A Python artifact that converts **CSV results exported from the Microsoft FinOps assessment / FinOps Review** into a prioritized FinOps education curriculum.

The assessment CSV is the primary analytical input. An optional people/roles workbook supplies the roster used for presenter lists and attendance planning.

## FinOps purpose

The artifact supports **FinOps Education & Enablement** by turning assessment evidence into a structured learning plan. It also connects assessment results with the broader FinOps capability catalogue used by the curriculum.

```text
Microsoft FinOps assessment CSV(s)
              |
              v
 Parse recommendations, priority and weight
              |
              v
 Aggregate assessment evidence
              |
              v
 Calculate implementation-specific gap score
              |
              v
 Prioritize domains and capabilities
              |
              v
 Generate weekly FinOps curriculum
              |
              v
 Track delivery / attendance
              |
              v
 Reassess and reprioritize
```

## Important methodological distinction

Microsoft supplies the assessment results and recommendations. **The curriculum ordering and scoring formula are implemented by this artifact; they are not represented as a Microsoft scoring standard.**

Current implementation:

```text
Priority points:
High   = 3
Medium = 2
Low    = 1

Individual gap score =
Priority points * 100 + Weight

Group capability score =
mean(individual gap scores)
```

Divergence is the range between the maximum and minimum priority-point values when at least two responses are available.

See [Methodology](docs/methodology.md).

## Curriculum baseline

The current code defines 31 learning opportunities:

- 5 foundation sessions
- 4 domain-introduction sessions
- 22 capability sessions

Foundation sessions remain first. Assessed domains are then ordered by average capability gap; domains without assessment evidence are placed afterward. Within each domain, the domain introduction precedes capabilities ordered by gap score.

## Requirements

- Python 3.10+
- `openpyxl`

```bash
pip install -r requirements.txt
```

## Quick start

1. Copy `.env.example` to `.env`.
2. Put one or more unmodified assessment CSV exports in `input/`.
3. Optionally add an `.xlsx` people/roles roster.
4. Run interactively:

```bash
python generar_curriculo_finops.py
```

or directly:

```bash
python generar_curriculo_finops.py --auto
```

## Inputs

### Primary input — assessment CSV

The parser expects the structure exported by the Microsoft assessment used by this artifact, including recommendation fields such as:

- `Category`
- `Link-Text`
- `Priority`
- `Weight`
- `Context`

It also reads domain-level section/maturity information when present.

### Optional roster

An `.xlsx` file can provide people, FinOps roles and titles. The tool uses it for the `PERSONAS` sheet, presenter dropdown and default invitee count.

No real roster is required for the analytical prioritization.

## Output workbook

The generated workbook contains:

- `RESUMEN`
- `CURRICULO`
- `PERSONAS`
- `FUENTES`

It combines the prioritized learning agenda with editable delivery fields such as presenter, status, invitees, attendees and notes.

## Data protection

Do not publish:

- assessment CSVs containing identifiable participant information;
- real people/roles rosters;
- generated operational workbooks;
- `.env`;
- organization-specific filenames or paths.

The public repository contains synthetic samples only.

## Documentation

- [Methodology](docs/methodology.md)
- [Architecture](docs/architecture.md)
- [Configuration](docs/configuration.md)
- [Data model](docs/data_model.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Changelog](CHANGELOG.md)
- [Notice](NOTICE.md)

## Version

**1.0.0 — sanitized certification/portfolio baseline**

No open-source license is asserted by this package. See `NOTICE.md`.
