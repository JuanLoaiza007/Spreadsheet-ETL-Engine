# Spreadsheet ETL Engine
A configurable ETL (Extract, Transform, Load) engine for spreadsheet-based data workflows.  
Designed to be flexible, safe, and user-friendly for non-technical users through configuration sheets instead of code changes.
Tested on Google Sheets + Google Apps Script.

## Overview
SheetFlow ETL is a generic data transformation engine built for spreadsheet environments.  
It enables dynamic mapping, filtering, and formula evaluation using structured configuration tables.

The engine is designed so that:

- Developers maintain the core logic.
- End-users only modify configuration sheets.
- No direct code interaction is required.

## Key Features
- Configuration-driven data transformation
- Safe expression evaluation
- Column mapping and renaming
- Conditional filters
- Constant and formula columns
- Self-reference column formulas
- Syntax validation
- Parentheses and operator integrity checks
- Non-technical user friendly dashboard execution
- Clear error reporting

## Architecture Concept
The system separates responsibilities into three layers:
1. **Source Layer**  
   Raw input data.
2. **Mapping Layer**  
   Rules, filters, formulas, and transformations.
3. **Output Layer**  
   Final structured dataset.

This separation allows high flexibility without modifying the core engine.

## Configuration Structure
### Dashboard Sheet
Defines sheet names and execution configuration.
| Key    | Value  |
|--------|--------|
| source | Source |
| map    | Map    |
| output | Output |

### Mapping Sheet
Each row defines a transformation rule.

#### Prefixes
| Prefix         | Meaning                    |
|----------------|----------------------------|
| `_filter:`     | Filter rule                |
| `eval:`        | Evaluated expression       |
| `constant:`    | Static value               |
| `formula:`     | Spreadsheet formula        |
| `src[column]`  | Reference source column    |
| `self[column]` | Reference generated column |


## Example Rules

### Direct Mapping
```
Name -> src[Name]
```

### Constant
```
Status -> constant:Active
```

### Formula
```
Total -> formula:=A2+B2
```

### Filter
```
_filter:Age -> eval: src(Age) >= 18
```

## Safety Mechanisms
- Operator validation
- Parentheses balancing
- Column existence verification
- Safe evaluation logic
- Controlled expression parsing
- Error isolation and reporting

## Target Users
- Administrative staff
- Analysts
- Operations teams
- Non-technical spreadsheet users
- Developers needing configurable ETL pipelines

## Use Cases
- Data normalization
- Report preparation
- Spreadsheet migration
- Administrative automation
- Data cleaning pipelines
- Internal data workflows

## Design Principles
- Configuration over code
- Safety over flexibility
- Transparency over magic
- Generic and reusable
- Minimal user friction
- Clear error communication

## Limitations
- Not intended for massive datasets
- Depends on spreadsheet performance limits
- Expression evaluation is intentionally restricted

## Roadmap Ideas
- Advanced expression parser
- Multi-sheet joins
- Type validation
- Logging dashboard
- Execution history
- Plugin transformation system

## License
This project is licensed under a permissive open-source license.  
See the `LICENSE` file for details.

## Technical Foundations
For detailed explanations of the parsing logic, expression validation,
and regular expression design, see the `/docs` directory.