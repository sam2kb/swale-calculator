# swale-calculator

Generates the stormwater retention swale calculator Excel workbook (`.xlsx`).

## Why this exists

This workbook provides a quick, consistent check that proposed swales supply the required stormwater retention volume for a residential lot. It standardizes assumptions, formulas, and layout so every project is calculated the same way across the team.

It is used by [Witty Finch Engineering](https://wittyfinch.com) in their residential plan sets.

## How the math works

- Site stats:
	- Acres = `Sq. feet / 43,560`.
	- Percent = each area divided by lot size.
- Required retention (cf): uses the larger of:
	- `0.5 inch over total lot area`, and
	- `1.0 inch over impervious area`.
- Swale volume (cf):
	- `V-Shape` (accounts for sloped short sides): `h × Wt × (2Lt + Lb) / 6`, where
		- `h = depth`,
		- `Wt = top width`,
		- `Lt = top length`,
		- `Lb = max(0, Lt - Wt)`.
	- `Trapezoid` (frustum): `h/3 × (A1 + A2 + sqrt(A1×A2))`, where
		- `h = depth`,
		- `A1 = bottom width × bottom length`,
		- `A2 = top width × top length`.
- Provided retention (cf): sum of swales.

## Retention controls

Under **Stormwater Retention**:

- `Retention basis`: `MAX`, `LOT ONLY`, or `IMP ONLY`.
- `Side slope ratio (H:V)`: global slope constant used by swale geometry formulas.
- `Input highlight`: `ON/OFF` toggle for yellow background on editable input cells.

## Built-in guardrails

The workbook includes validation checks to prevent invalid geometry entries.

- `Trapezoid` top width and top length must each be at least `2.0 ft`.
- Bottom width/length cells are calculated and locked from manual edits.

## Requirements

- Python 3.10+

## Build

```bash
chmod +x build.sh
./build.sh
```

The script will:

- create `.venv` if needed,
- install dependencies from `requirements.txt`,
- generate the workbook at `build/swale_calculator.xlsx`.

## Run script directly (optional)

```bash
python3 residential.py --out build/swale_calculator.xlsx
```
