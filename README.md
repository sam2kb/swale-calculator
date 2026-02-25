# swale-calculator

Generates the swale calculator Excel workbook (`.xlsx`).

## Why this exists

This workbook is used to quickly check whether proposed swales provide enough retention volume for a residential lot.
It standardizes the calculation so teams can use the same assumptions, formulas, and layout in every project.

## How the math works

- Site stats:
	- Acres = `Sq. feet / 43,560`.
	- Percent = each area divided by lot size.
- Required retention (cf): uses the larger of:
	- `0.5 inch over total lot area`, and
	- `1.0 inch over impervious area`.
- Swale volume (cf):
	- `V-Shape`: `0.5 × top width × depth × length`
	- `Trapezoid` (frustum): `h/3 × (A1 + A2 + sqrt(A1×A2))`, where
		- `h = depth`,
		- `A1 = bottom width × bottom length`,
		- `A2 = top width × top length`.
- Provided retention (cf): sum of swales.

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
