import pandas as pd
import os
import matplotlib.pyplot as plt

def summarize_tasks(file_path="tasks.xlsx"):
    if not os.path.exists(file_path):
        print("No task data found.")
        return

    df = pd.read_excel(file_path)

    # Convert duration to total hours
    def convert_to_hours(duration_str):
        h, m, s = map(int, duration_str.split(':'))
        return h + m / 60 + s / 3600

    df['Duration Hours'] = df['Duration'].apply(convert_to_hours)

    summary = df.groupby('Task Type')['Duration Hours'].sum().reset_index()
    summary.columns = ['Task Type', 'Total Hours']

    # Add a total row
    total_hours = summary['Total Hours'].sum()
    total_summary = pd.DataFrame([{'Task Type': 'Total', 'Total Hours': total_hours}])
    summary = pd.concat([summary, total_summary], ignore_index=True)

    # Print summary to console
    print(summary)

    # Save summary to Excel
    summary.to_excel("task_summary.xlsx", index=False)

    # Plotting
    plt.figure(figsize=(10, 6))
    plt.bar(summary['Task Type'], summary['Total Hours'], color=['blue', 'green', 'red', 'purple'])
    plt.xlabel('Task Type')
    plt.ylabel('Total Hours')
    plt.title('Total Hours by Task Type')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.savefig('task_summary_plot.png')  # Save the plot as a .png file
    plt.show()

if __name__ == "__main__":
    summarize_tasks()
