# Monthly Report Database - Flask Web UI

## Setup

1. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the application:
   ```
   python app.py
   ```

3. Open your browser to:
   ```
   http://localhost:5000
   ```

## Features

- **View Artifacts**: See all artifacts with their associated data sources
- **Add/Edit Artifacts**: Create new artifacts or modify existing ones
- **Delete Artifacts**: Mark artifacts as deleted (soft delete)
- **View Scans**: See all security scans with vulnerability counts
- **Add/Edit Scans**: Create new scan records or modify existing ones
- **Delete Scans**: Permanently remove scan records

## Database

The application connects to the `monthlyReport.db` SQLite database in the parent directory.

## Notes

- This is a single-user local application with no authentication
- The application runs in debug mode for development
- All database constraints from your schema are enforced
- Flash messages provide feedback on successful/failed operations
