## 4. JSON Configuration System

The system uses a `config.json` file for easy customization. It will be auto-created on first run with sensible defaults.

### Key Configuration Options:

#### Work Schedule:
```json
"work_schedule": {
    "work_days_per_week": 5,
    "work_days": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "start_time": "07:00",
    "end_time": "17:00",
    "include_weekends": false,
    "weekend_premium": 1.3
}
```

#### Team Defaults:
```json
"team_defaults": {
    "max_daily_hours": 8.0,
    "overtime_allowed": true,
    "max_overtime_hours": 4.0,
    "base_hourly_rate": 35.0,
    "overtime_multiplier": 1.5,
    "default_mowing_rate_sqm_per_hour": 1200
}
```

#### Equipment Rates:
```json
"equipment_rates": {
    "commercial_mower": {
        "sqm_per_hour": 1500,
        "efficiency_factor": 1.0,
        "maintenance_cost_per_hour": 5.0
    },
    "ride_on_mower": {
        "sqm_per_hour": 1200,
        "efficiency_factor": 1.0,
        "maintenance_cost_per_hour": 3.5
    }
}
```

### Configuration Management:
- **Web Interface**: Visit `http://localhost:5000/config` to modify settings
- **Direct Edit**: Modify `config.json` file directly
- **API Access**: Use `/api/config` endpoints for programmatic updates# requirements.txt
flask==2.3.3
pandas==2.0.3
numpy==1.24.3
requests==2.31.0
plotly==5.15.0
openpyxl==3.1.2
ortools==9.7.2996
folium==0.14.0
sqlite3-tools==0.2.1

# Installation Instructions

## 1. System Requirements
- Python 3.8 or higher
- 4GB RAM minimum (8GB recommended for 400 parks)
- 1GB free disk space

## 2. Installation Steps

### Step 1: Install Python Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Get Weather API Key (Optional)
1. Sign up at https://openweathermap.org/api
2. Get your free API key
3. Replace "your_weather_api_key" in the code with your actual key

### Step 3: Run the Application
```bash
python mowing_scheduler.py
```

### Step 4: Access the Web Interface
Open your browser and go to: http://localhost:5000

## 3. Excel File Format for Parks Import

### Required Columns:
- **name**: Park name (e.g., "Central Park")
- **suburb**: Suburb name (e.g., "Northgate") 
- **difficulty**: Difficulty rating 1-5 (1=easy, 5=very difficult)
- **priority**: Priority level 1-5 (1=low, 5=urgent)

### Size Columns (at least one required):
- **size_sqm**: Park size in square meters (e.g., 12000)
- **size_hectares**: Park size in hectares (e.g., 1.2)

### Optional Columns:
- **id**: Unique park identifier (auto-generated if not provided)
- **latitude**: GPS latitude coordinate (for weather & routing optimization)
- **longitude**: GPS longitude coordinate (for weather & routing optimization)
- **last_mowed**: Last mowed date (YYYY-MM-DD format)
- **estimated_hours**: Manual time estimate (will be auto-calculated)

### Minimum Excel Example:
| name | suburb | size_sqm | difficulty | priority |
|------|--------|----------|------------|----------|
| Central Park | Northgate | 12000 | 3 | 4 |
| Sports Field | Chermside | 25000 | 4 | 5 |
| Playground Area | Kedron | 3000 | 2 | 2 |

### Full Excel Example with Optional Columns:
| id | name | suburb | latitude | longitude | size_sqm | size_hectares | difficulty | priority | last_mowed | estimated_hours |
|----|------|--------|----------|-----------|----------|---------------|------------|----------|------------|-----------------|
| 1 | Central Park | Northgate | -27.3812 | 153.0416 | 25000 | 2.5 | 3 | 4 | 2024-11-15 | 2.1 |
| 2 | Sports Complex | Chermside | -27.3854 | 153.0355 | 50000 | 5.0 | 4 | 5 | 2024-11-10 | 4.2 |

**Note**: The system works perfectly fine without latitude/longitude coordinates. It will use suburb-based routing and general weather data for the area.

## 4. Key Features Included

### AI Scheduling Optimization
- Uses Google OR-Tools for constraint satisfaction
- Considers team assignments to suburbs
- Handles overtime rules and limits
- Weather-aware scheduling
- Priority and urgency-based optimization

### Team Management
- Assign teams to specific suburbs
- Configure overtime allowances
- Set maximum daily and overtime hours
- Track team utilization and costs

### Weather Integration
- Real-time weather API integration
- Automatic rescheduling for poor conditions
- Weather suitability scoring
- Risk assessment and recommendations

### Excel Integration
- Import parks from Excel files
- Export schedules to formatted Excel
- Automatic formatting and charts
- Column auto-sizing and styling

### Gantt Chart Visualization
- Interactive timeline view
- Color-coded by priority
- Team-based organization
- Hover details with costs and overtime info

### Comprehensive Reporting
- Team utilization analysis
- Suburb-based statistics
- Priority completion tracking
- Cost analysis and forecasting

## 5. API Endpoints

The system provides REST APIs for integration:

- GET /api/parks - Retrieve all parks
- GET /api/teams - Retrieve all teams  
- POST /api/teams - Update team configuration
- GET /api/report - Get comprehensive analytics report
- GET /api/weather_impact - Get weather impact analysis
- POST /optimize - Run scheduling optimization
- GET /export - Download Excel schedule

## 6. Database

Uses SQLite for data persistence with tables for:
- parks: Park information and maintenance history
- teams: Team configurations and assignments
- scheduled_jobs: Generated schedules and assignments

## 7. Customization Options

### Adding New Constraints
Modify the `optimize_schedule` method in `SchedulingOptimizer` class to add:
- Equipment-specific requirements
- Seasonal adjustments
- Customer preferences
- Maintenance windows

### Custom Weather Rules
Modify `calculate_mowing_suitability` in `WeatherService` class for:
- Local weather patterns
- Seasonal considerations
- Equipment-specific weather limits

### Reporting Extensions
Add new reports in `ReportGenerator` class for:
- Equipment utilization
- Fuel consumption tracking
- Customer satisfaction metrics
- Maintenance scheduling

## 8. Scaling for Larger Operations

For operations beyond 400 parks:
- Enable database clustering
- Implement caching with Redis
- Use asynchronous processing
- Consider microservices architecture

## 9. Production Deployment

For production use:
- Use a production WSGI server (Gunicorn)
- Set up proper logging
- Configure environment variables
- Use PostgreSQL instead of SQLite
- Implement user authentication
- Set up automated backups

Example production start:
```bash
gunicorn -w 4 -b 0.0.0.0:5000 mowing_scheduler:app
```

## 10. Support and Maintenance

The system includes:
- Comprehensive error handling
- Detailed logging
- Automatic database backups
- Performance monitoring hooks
- Health check endpoints