# Rack_Unit Factor Lookup Locations in Code

## Summary
The `rack_unit` column is looked up from dataframes in multiple locations throughout the code. Here are the key locations:

---

## 1. **Setup DOE Design** (`doe.py`, lines 53-55)
**Function:** `setup_doe_design()`
**Purpose:** Extract min, max, and mean values of rack_unit from the combined DOE dataframe

```python
# Line 53-55 in setup_doe_design()
rack_unit_min = doe_df['rack_unit'].min()
rack_unit_max = doe_df['rack_unit'].max()
rack_unit_mean = doe_df['rack_unit'].mean()
```

**Context:** This happens after combining low and high speed fan dataframes:
```python
doe_df = pd.concat([low_df, high_df], ignore_index=True)
```

**Data Flow:**
- Source: `doe_df` (combined from balanced_low_df + balanced_high_df)
- Lookup type: Statistical aggregation (min/max/mean)
- Used for: Display and documentation

---

## 2. **Create Full-Factorial Design** (`doe.py`, lines 99-100)
**Function:** `create_full_factorial_design()`
**Purpose:** Create discrete rack_unit levels (1-42) instead of using actual values

```python
# Lines 99-100
# Create discrete rack_unit levels as integers from 1 to 42
rack_units = list(range(1, 43))  # Integer increments 1 through 42
```

**Context:** Used in design generation loop:
```python
for tman in transceiver_mfrs:
    for rack_unit in rack_units:          # Lines 107
        for speed in speed_ranges:
            design_rows.append({
                'Rack_Unit': rack_unit,   # Line 112
                ...
            })
```

**Data Flow:**
- Source: `doe_df` is passed but NOT directly used for rack_unit lookup
- Instead: Hardcoded range(1, 43) is used
- Purpose: Create design space with all possible rack positions (1-42)

---

## 3. **Fit Full DOE Model** (`doe.py`, lines 205-209)
**Function:** `fit_doe_model()`
**Purpose:** Extract `rack_unit` column and rename it to `Rack_Unit` for model fitting

```python
# Lines 205-209
# Prepare data for model - encode categorical variables
model_df = doe_df[['SFP_manufacturer', 'Fan_Speed_Range', 'rack_unit', 'Interface_Temp']].copy()
model_df.columns = ['Transceiver_Manufacturer', 'Fan_Speed_Range', 'Rack_Unit', 'ttemp']

# Convert Rack_Unit to integer for categorical treatment
model_df['Rack_Unit'] = model_df['Rack_Unit'].astype(int)
```

**Context:** The column is then used in the model formula:
```python
# Line 213
formula = 'ttemp ~ C(Transceiver_Manufacturer) + C(Rack_Unit) + C(Fan_Speed_Range) + ...'
```

**Data Flow:**
- Source: `doe_df` (model data passed to function)
- Columns extracted: `SFP_manufacturer`, `Fan_Speed_Range`, `rack_unit`, `Interface_Temp`
- Column renamed: `rack_unit` → `Rack_Unit`
- Converted to: `int` type for categorical treatment
- Used in: OLS regression formula with `C(Rack_Unit)` for categorical encoding

---

## 4. **Fit Reduced DOE Model** (`doe.py`, lines 473-475)
**Function:** `fit_reduced_doe_model()`
**Purpose:** Same as full model - extract and rename rack_unit

```python
# Lines 473-475
model_df = doe_df[['SFP_manufacturer', 'Fan_Speed_Range', 'rack_unit', 'Interface_Temp']].copy()
model_df.columns = ['Transceiver_Manufacturer', 'Fan_Speed_Range', 'Rack_Unit', 'ttemp']
model_df['Rack_Unit'] = model_df['Rack_Unit'].astype(int)
```

**Context:** Identical to full model preparation for regression

---

## 5. **Create Interaction Plots** (`doe.py`, lines 698-705)
**Function:** `create_interaction_plots()`
**Purpose:** Get unique levels of `Rack_Unit` for visualization

```python
# Lines 698-705 (relevant to Rack_Unit handling)
# Get unique levels for each factor
levels_factor1 = sorted(model_df[factor1_raw].unique())
levels_factor2 = sorted(model_df[factor2_raw].unique())

# For continuous factors (like Rack_Unit), use a sample of levels
if len(levels_factor1) > 6 and isinstance(levels_factor1[0], (int, float)):
    # For continuous factors, sample evenly distributed levels
    step = max(1, len(levels_factor1) // 6)
    levels_factor1 = levels_factor1[::step][:6]
```

**Data Flow:**
- Source: `model_df` (contains renamed `Rack_Unit` column)
- Lookup: `model_df[factor_name].unique()` to get all unique values
- Sampling: For continuous factors like Rack_Unit, evenly distributed sample (max 6 levels)
- Purpose: Create representative interaction plot with subset of levels

---

## 6. **Duplicate Observations Grouping** (`doe.py`, line 297)
**Function:** `_calculate_lack_of_fit()`
**Purpose:** Group observations by all factors for lack-of-fit test

```python
# Line 297
group_cols = ['Transceiver_Manufacturer', 'Rack_Unit', 'Fan_Speed_Range']
```

**Context:** Used for grouping:
```python
model_df.groupby(group_cols).apply(...)  # To find duplicate combinations
```

**Data Flow:**
- Source: `model_df` (contains `Rack_Unit` column)
- Purpose: Identify rows with identical factor combinations for replicate observations
- Used in: Lack-of-fit (LOF) ANOVA calculation

---

## Data Type Conversions

### Original Data (from CSV/dataframe):
```
rack_unit: float64 (raw values from CSV, e.g., 1.0, 2.0, ..., 42.0)
```

### After Renaming:
```
Rack_Unit: float64 initially (after rename)
```

### After Conversion:
```
Rack_Unit: int64 (after astype(int), used in OLS regression)
```

### In Design:
```
rack_units: list of int (hardcoded 1-42, not extracted from data)
```

---

## Key Insight: rack_unit vs Design Levels

**Important:** The code distinguishes between:

1. **Actual data values** (`doe_df['rack_unit']`):
   - Extracted from the combined balanced dataframe
   - Used in model fitting with actual observed values
   - Range: 1.0 to 42.0 (continuous-like but categorical in model)

2. **Design space** (`list(range(1, 43))`):
   - Hardcoded 1 through 42
   - Represents all possible rack positions
   - Used only in full-factorial design table generation
   - Does NOT necessarily match all values in data

This means the model can be trained on whatever rack_unit values exist in the data, while the design table represents the full factorial space (1-42).

---

## Summary Table

| Location | File | Lines | Operation | Source |
|----------|------|-------|-----------|--------|
| Statistics | doe.py | 53-55 | min/max/mean | doe_df['rack_unit'] |
| Design Generation | doe.py | 99-100 | Create levels | hardcoded range(1,43) |
| Model Prep (Full) | doe.py | 205-209 | Extract & rename | doe_df['rack_unit'] |
| Model Prep (Reduced) | doe.py | 473-475 | Extract & rename | doe_df['rack_unit'] |
| Interaction Plots | doe.py | 698-705 | Get unique values | model_df['Rack_Unit'] |
| LOF Test | doe.py | 297 | Grouping column | model_df['Rack_Unit'] |
