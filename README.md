# Fuel Log Generator

Generate Excel workbooks for tracking fuel expenses and travel logs.

## Features

- Creates a properly formatted Excel log book for travel expenses
- Supports configuration via JSON file or command line arguments
- Automatically calculates odometer readings across months
- Properly highlights weekends
- Adds company details, vehicle information, and employee data
- Handles date ranges from any start to end month/year

## Quick Start

```bash
# Run with default settings
python fuel_log_v2.py

# Run with a configuration file
python fuel_log_v2.py --config config.json

# Run with command line arguments
python fuel_log_v2.py --start-date 2024-08-01 --end-date 2025-03-31 --output "my_log_book.xlsx"
```

## Configuration

You can configure the generator either through a JSON file or command line arguments.

### Configuration File Example

```json
{
  "start_date": "2024-08-01",
  "end_date": "2025-03-31",
  "initial_odometer": 17569,
  "inr_per_km": 10,
  "work_related_km": 110,
  "trip_purpose": "Official",
  "client_name": "Blink Charging",
  "is_work_travel": "Y",
  "personal_travel": "",
  "employee": {
    "name": "Ashish Kumar",
    "id": "BLINKIN065",
    "department": "Technology",
    "manager": "Ajay Singh"
  },
  "vehicle": {
    "make": "Hyundai",
    "model": "Xcent",
    "year": "2018",
    "registration": "Delhi",
    "engine_size": "1199 CC"
  },
  "output_file_path": "Financial_Year_2024_25_Log_Book.xlsx"
}
```

### Command Line Arguments

```
--config, -c       Path to JSON configuration file
--output, -o       Output Excel file path
--start-date       Start date in YYYY-MM-DD format
--end-date         End date in YYYY-MM-DD format
--initial-odometer Initial odometer reading
--km-per-day       Work-related kilometers per day
--rate-per-km      Rate per kilometer in INR
```

## Output

The generator produces an Excel workbook with:
- One sheet per month in the specified date range
- Company headers and employee details
- Vehicle information
- Daily travel log with dates, odometer readings, and expense calculations
- Weekend days highlighted in red
- Monthly totals
