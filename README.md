# JIRA Data Fetcher

This script connects to a JIRA server, retrieves information about boards, sprints, and issues, and exports the data to an Excel file. The script also calculates the total work time for each assignee in each sprint.

## Requirements

- Python 3.x
- JIRA Python library
- pandas library
- openpyxl library

## Installation

First, ensure you have Python 3 installed on your machine. Then, install the required libraries using pip:

```bash
pip install jira pandas openpyxl
```

## Configuration

Before running the script, update the following variables in the script with your JIRA server details and credentials:

```python
jira_server = 'https://domainname.com:8085'
jira_user = 'your_username'
jira_password = 'your_password'
```

## Running the Script

To run the script, navigate to the directory containing the script and execute the following command:

```bash
python fetch_jira_data.py
```

This will connect to the specified JIRA server, retrieve the data, and export it to an Excel file named `jira_tasks.xlsx` in the same directory.

## Output

The output Excel file `jira_tasks.xlsx` will contain separate sheets for each sprint, with detailed information about each issue, including:

- Board name
- Sprint name
- Issue key
- Issue type
- Priority
- Reporter
- Status
- Summary
- Resolution
- Original estimate
- Time spent
- Work ratio
- Progress
- Assignee

Additionally, the total time spent by each assignee in each sprint will be calculated and included in the data.

## Contact

If you encounter any issues or have any questions, please contact me at [nakhoda.id@gmail.com].