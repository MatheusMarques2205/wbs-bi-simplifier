# WBS BI Simplifier

## About
Tool to simplify WBS data from MS Project for Power BI visualization, focused on wind energy projects. Transforms thousands of tasks into 19 main milestones.

## How to Use

### Requirements
- Python 3.x
- Microsoft Access
- Power BI Desktop
- Python packages: pyodbc, pandas, sqlalchemy

### Step by Step

1. Configure Python file:
   - Open "WBS (Work Breakdown Structure).py"
   - Update database path:
   ```python
   DATABASE_PATH = "your_database_path.mdb"
   ```

2. Run the program:
   - Open terminal
   - Run: `python "WBS (Work Breakdown Structure).py"`

3. In Power BI:
   - Import both tables using Power Query
   - Create relationship between tables using TASK_ID
   - Create your visualizations

## Main Milestones
The program categorizes tasks into:
1. NTP
2. Anchor Cages EX Works
3. Anchor Cages Delivery at POD
4. Anchor Cages Delivery at Site
5. WTG Ex Works Local
6. Foundation Ready
7. WTG Delivery at Site
8. Main Crane Mobilization
9. Pre Erection
10. Full Erection
11. Main Crane Demobilization
12. Electrical and Mechanical Completion
13. SCADA Installation
14. Pre Commissioning
15. Substation Energized
16. Commissioning
17. Run Test
18. COD

## Support
For questions or issues, please open an issue on GitHub.
