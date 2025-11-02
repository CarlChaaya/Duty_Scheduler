# Duty Scheduler

A production-ready Flask service that generates balanced duty rosters for medical internship programs. The optimiser ensures minimum staffing levels, fair rotations, and configurable spacing between shifts. Downloadable Excel schedules are created on demand so coordinators can share them immediately.

## Features
- Constraint-programming solver (PuLP) with fairness metrics and weekend weighting
- Web form for date range, staffing requirements, and intern capacity
- Automatic Excel export styled for distribution
- Configurable via environment variables (`FLASK_SECRET_KEY`, `PORT`)

## Quick Start
```bash
python -m venv .venv
. .venv/Scripts/Activate.ps1  # or source .venv/bin/activate on macOS/Linux
pip install -r requirements.txt
set FLASK_SECRET_KEY=change-me  # use export on macOS/Linux
flask --app app run
```

To build a production deployment:
1. Set `FLASK_SECRET_KEY` and `PORT` in your hosting environment
2. Run `gunicorn --workers 4 --bind 0.0.0.0:$PORT app:app`

## Project Layout
- `app.py` – Flask entrypoint, request validation, and file-response lifecycle
- `scheduler.py` – Optimisation model, Excel export helper utilities
- `templates/` – Jinja templates for the HTML form
- `temp/` – Transient folder used to stage the generated XLSX (ignored by Git)

## Roadmap
- [x] MVP intern scheduler with fairness constraints
- [ ] Upload roster demand data instead of manual entry
- [ ] Container image + CI deployment workflow
- [ ] Reporting dashboard for coverage statistics

Contribution guidelines and issues will be tracked in GitHub Projects once the next milestone starts.
