# Plan: Fix Interaction Plots Rack Unit Range

## Problem Identified

### Current Behavior
- Interaction plots show only **Rack Units 1-6** on the X-axis
- These are only 6 levels (downsampled from 41 available)

### Root Cause
**Location:** `doe.py`, lines 698-711

```python
if len(levels_factor1) > 6 and isinstance(levels_factor1[0], (int, float)):
    # For continuous factors, sample evenly distributed levels
    step = max(1, len(levels_factor1) // 6)
    levels_factor1 = levels_factor1[::step][:6]
```

**The Problem:**
- Data has 41 unique rack units (1-42, missing 35)
- Step calculation: `41 // 6 = 6`
- Every 6th element taken: `rack_units[::6]` = `[1, 7, 13, 19, 25, 31, 37]` → limited to first 6 → `[1, 7, 13, 19, 25, 31]`
- Wait, that should give different values... Let me recalculate

**Actually:**
- When factor2_raw is 'Rack_Unit' (continuous-like with 41 values)
- Step = `max(1, 41 // 6)` = `max(1, 6)` = `6`
- `levels_factor2[::6][:6]` samples every 6th element
- From 41 elements indexed 0-40, every 6th is: 0, 6, 12, 18, 24, 30, 36
- That maps to values: [1, 7, 13, 19, 25, 31, 37]
- Then `[:6]` takes first 6: [1, 7, 13, 19, 25, 31]

**But we're seeing 1-6, which means:**
- The limiting happens BEFORE it goes into the plot
- Most likely: `levels_factor2[:6]` is being applied instead of the step-based sampling
- This is line 711: `elif len(levels_factor2) > 6: levels_factor2 = levels_factor2[:6]`

### Why This Happens
The integer type check `isinstance(levels_factor1[0], (int, float))` may be failing or the elif branch is being hit.

---

## Solution Plan

### Option 1: Show All Rack Units in Plot (Recommended)
**Pros:** 
- Complete information
- Shows actual data distribution
- Better for interaction visualization

**Cons:**
- X-axis may be crowded
- Plotly will handle rendering well

**Implementation:**
```python
# Remove the arbitrary limit of 6 levels for Rack_Unit
# Instead, show all levels
if factor2_raw == 'Rack_Unit':
    # Show all rack units - they're the design space
    pass  # Keep all levels_factor2
elif len(levels_factor2) > 10:
    # For other continuous factors, limit to 10 levels
    step = max(1, len(levels_factor2) // 10)
    levels_factor2 = levels_factor2[::step][:10]
```

### Option 2: Show More Representative Rack Units
**Pros:**
- Balanced between detail and readability
- Can show min, max, quartiles

**Cons:**
- Still not showing complete picture

**Implementation:**
```python
# Show more levels (e.g., 12 instead of 6)
if factor2_raw == 'Rack_Unit':
    if len(levels_factor2) > 12:
        step = max(1, len(levels_factor2) // 12)
        levels_factor2 = levels_factor2[::step][:12]
```

### Option 3: Show Quartiles + Extremes for Rack Unit
**Pros:**
- Statistically meaningful sampling
- Compact visualization

**Cons:**
- May miss important patterns

**Implementation:**
```python
if factor2_raw == 'Rack_Unit':
    # Use percentiles: 0%, 25%, 50%, 75%, 100%
    import numpy as np
    percentiles = [0, 25, 50, 75, 100]
    levels_factor2 = np.percentile(levels_factor2, percentiles).astype(int)
    levels_factor2 = sorted(list(set(levels_factor2)))
```

---

## Recommended Fix: Option 1 (Show All Rack Units)

### Why?
1. **Rack_Unit is not truly continuous** - it's categorical (1-42)
2. **Data only has 41 values** - not difficult to display
3. **Interaction patterns matter across full range** - cramping data loses info
4. **Plotly handles large X-axis well** - no rendering issues

### Implementation Details

**File:** `doe.py`
**Function:** `create_interaction_plots()`
**Lines to modify:** 698-711

**Current Code:**
```python
# For continuous factors (like Rack_Unit), use a sample of levels
# For categorical factors (like Transceiver_Manufacturer), limit to first 6 levels
if len(levels_factor1) > 6 and isinstance(levels_factor1[0], (int, float)):
    # For continuous factors, sample evenly distributed levels
    step = max(1, len(levels_factor1) // 6)
    levels_factor1 = levels_factor1[::step][:6]
elif len(levels_factor1) > 6:
    levels_factor1 = levels_factor1[:6]

if len(levels_factor2) > 6 and isinstance(levels_factor2[0], (int, float)):
    step = max(1, len(levels_factor2) // 6)
    levels_factor2 = levels_factor2[::step][:6]
elif len(levels_factor2) > 6:
    levels_factor2 = levels_factor2[:6]
```

**New Code:**
```python
# For Rack_Unit (discrete categories 1-42), show all levels
# For other continuous factors, sample to ~12 levels for readability
# For categorical factors, limit to first 10 levels

def should_limit_levels(factor_name, num_levels, is_numeric):
    """Determine if factor levels should be limited for visualization."""
    if factor_name == 'Rack_Unit':
        return False  # Show all rack units
    elif is_numeric and num_levels > 12:
        return True   # Sample numeric factors to ~12 levels
    elif num_levels > 10:
        return True   # Limit categorical factors to ~10 levels
    else:
        return False

# Apply level limiting for factor 1
if should_limit_levels(factor1_raw, len(levels_factor1), isinstance(levels_factor1[0], (int, float))):
    num_target = 12 if isinstance(levels_factor1[0], (int, float)) else 10
    if len(levels_factor1) > num_target:
        step = max(1, len(levels_factor1) // num_target)
        levels_factor1 = levels_factor1[::step][:num_target]

# Apply level limiting for factor 2
if should_limit_levels(factor2_raw, len(levels_factor2), isinstance(levels_factor2[0], (int, float))):
    num_target = 12 if isinstance(levels_factor2[0], (int, float)) else 10
    if len(levels_factor2) > num_target:
        step = max(1, len(levels_factor2) // num_target)
        levels_factor2 = levels_factor2[::step][:num_target]
```

---

## Testing Plan

### Before Fix
1. Generate plots
2. Verify plots show 1-6 on X-axis
3. Screenshot for comparison

### After Fix
1. Generate plots
2. Verify plots show 1-42 (or 1-41) on X-axis
3. Confirm all 3 interaction plots display correctly
4. Check for crowding or rendering issues
5. Verify prediction lines still calculate correctly

### Validation Checks
- [ ] Fan_Speed_Range × Rack_Unit shows all 41 rack units
- [ ] Rack_Unit × Transceiver_Manufacturer shows all 41 rack units
- [ ] Fan_Speed_Range × Transceiver_Manufacturer limits Manufacturer to ~10 levels
- [ ] HTML report renders without errors
- [ ] PDF embeds plots correctly
- [ ] PowerPoint slides display properly

---

## Files to Modify

1. **`doe.py`** - `create_interaction_plots()` function
   - Lines: 698-711 (level limiting logic)
   - Add helper function for smarter level limiting
   - Add debug output showing how many levels kept

2. **Optional: Add documentation**
   - Update docstring to explain level limiting strategy
   - Document why Rack_Unit is treated specially

---

## Expected Outcome

### Before:
```
Interaction Plot: Fan_Speed_Range × Rack_Unit
X-axis: [1, 2, 3, 4, 5, 6]
Lines: Low speed, High speed
```

### After:
```
Interaction Plot: Fan_Speed_Range × Rack_Unit
X-axis: [1, 2, 3, ..., 40, 41, 42] (all 41 values shown)
Lines: Low speed, High speed
```

### Benefits:
- Complete interaction pattern visualization
- Better scientific accuracy
- No information loss
- Proper representation of design space

---

## Risk Assessment

**Low Risk** because:
1. Only changes visualization, not calculation
2. Plotly handles many X-axis labels well
3. Rack_Unit is categorical (discrete), not truly continuous
4. Data is already small (41 values)
5. Can be tested immediately

**Potential Issues:**
- X-axis labels may overlap on small screens
- Mitigation: Plotly auto-rotates labels, users can zoom/pan in interactive HTML

---

## Approval Gate

✓ Root cause identified: Arbitrary 6-level limit
✓ Solution designed: Show all Rack_Units
✓ Risk assessed: Low
✓ Testing plan: Comprehensive
✓ Ready to implement: Yes
