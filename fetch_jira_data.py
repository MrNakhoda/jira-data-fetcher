from jira import JIRA
import pandas as pd

def seconds_to_workday_time(seconds):
    # Convert seconds to workday time (6-hour workdays)
    if seconds is None:
        seconds = 0
    workday_seconds = 6 * 3600  # Number of seconds in a 6-hour workday
    workdays = seconds // workday_seconds
    remaining_seconds = seconds % workday_seconds

    hours = remaining_seconds // 3600
    minutes = (remaining_seconds % 3600) // 60

    result = ''
    if workdays > 0:
        result += f"{workdays}D "
    if hours > 0:
        result += f"{hours}H "
    if minutes > 0:
        result += f"{minutes}M "

    return result.strip()

# JIRA user information and URL
jira_server = 'https://domainname.com:8085'
jira_user = 'your_username'
jira_password = 'your_password'

# Connect to JIRA
options = {'server': jira_server}
jira = JIRA(options, basic_auth=(jira_user, jira_password))

# Retrieve all boards
boards = jira.boards()

# Final data
all_data = {}

# Extract information from each board
for board in boards:
    # Retrieve sprints from each board
    sprints = jira.sprints(board.id, state='active,closed')  # Get all sprints, regardless of state
    for sprint in sprints:
        # Retrieve tasks from each sprint
        issues = jira.search_issues(f'sprint = {sprint.id}', maxResults=False)
        data = []
        for issue in issues:
            time_spent = issue.fields.timespent if hasattr(issue.fields, 'timespent') else 0
            original_estimate = issue.fields.timeoriginalestimate if hasattr(issue.fields, 'timeoriginalestimate') else 0
            
            task_data = {
                'Board': board.name,
                'Sprint': sprint.name,
                'Key': issue.key,
                'Type': issue.fields.issuetype.name,
                'Priority': issue.fields.priority.name if issue.fields.priority else 'None',
                'Reporter': issue.fields.reporter.displayName if issue.fields.reporter else 'None',
                'Status': issue.fields.status.name,
                'Summary': issue.fields.summary,
                'Resolution': issue.fields.resolution.name if issue.fields.resolution else 'Unresolved',
                'Original Estimate': seconds_to_workday_time(original_estimate),
                'Time Spent': seconds_to_workday_time(time_spent),
                'Time Spent (Seconds)': time_spent,
                'Work Ratio': issue.fields.workratio,
                'Progress': issue.fields.progress.progress,
                'Assignee': issue.fields.assignee.displayName if issue.fields.assignee else 'Unassigned'
            }
            data.append(task_data)

        # Add data to the dictionary with the sprint name
        all_data[sprint.name] = data

# Calculate the total work time for each person in each sprint
for sprint_name, sprint_data in all_data.items():
    assignee_time = {}
    for task in sprint_data:
        assignee = task['Assignee']
        time_spent_seconds = task['Time Spent (Seconds)'] if 'Time Spent (Seconds)' in task else 0
        time_spent_seconds = time_spent_seconds or 0  # Ensure it's an integer
        if assignee in assignee_time:
            assignee_time[assignee] += time_spent_seconds
        else:
            assignee_time[assignee] = time_spent_seconds
    
    for assignee, time_spent in assignee_time.items():
        sprint_data.append({
            'Assignee': assignee,
            'Total Time Spent (Seconds)': time_spent,
            'Total Time Spent (Formatted)': seconds_to_workday_time(time_spent)
        })

# Create an Excel file with separate sheets for each sprint
with pd.ExcelWriter('jira_tasks.xlsx', engine='openpyxl') as writer:
    for sprint_name, data in all_data.items():
        df = pd.DataFrame(data)
        
        # Separate information for each person
        unique_assignees = df['Assignee'].unique()
        writer.book.create_sheet(sprint_name)
        start_row = 0
        for assignee in unique_assignees:
            assignee_data = df[df['Assignee'] == assignee]
            assignee_data.to_excel(writer, sheet_name=sprint_name, startrow=start_row, index=False)
            start_row += len(assignee_data) + 2  # Add space between each person's data

print("Data has been successfully exported to jira_tasks.xlsx")
