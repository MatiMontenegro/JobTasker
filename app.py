import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QInputDialog
from datetime import datetime
import subprocess

class TaskApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Task Logger")
        self.setGeometry(100, 100, 400, 300)

        self.startButton = QPushButton("Start Task", self)
        self.startButton.clicked.connect(self.start_task)
        self.startButton.setGeometry(50, 50, 150, 50)

        self.stopButton = QPushButton("Stop Task", self)
        self.stopButton.clicked.connect(self.stop_task)
        self.stopButton.setGeometry(50, 120, 150, 50)

        self.changeButton = QPushButton("Change Task", self)
        self.changeButton.clicked.connect(self.change_task)
        self.changeButton.setGeometry(50, 190, 150, 50)

        self.openLogButton = QPushButton("Open Log", self)
        self.openLogButton.clicked.connect(self.open_log)
        self.openLogButton.setGeometry(220, 50, 150, 50)

        self.task_type = ""
        self.details = ""
        self.start_time = None

    def start_task(self):
        self.start_time = datetime.now()
        self.task_type, ok = QInputDialog.getItem(self, "Task Type", "Select Task Type:", ["Development", "Support", "Meetings"], 0, False)
        if ok:
            self.details, ok = QInputDialog.getText(self, "Task Details", "Enter Task Details:")
            if ok:
                self.statusBar().showMessage(f"Task started: {self.task_type}")

    def stop_task(self):
        if not self.start_time:
            self.statusBar().showMessage("No task is currently running.")
            return

        end_time = datetime.now()
        duration = end_time - self.start_time

        # Convert duration to a more readable format
        total_seconds = duration.total_seconds()
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        duration_str = f"{hours:02}:{minutes:02}:{seconds:02}"

        data = {
            "Date": [end_time.strftime("%Y-%m-%d")],
            "Start Time": [self.start_time.strftime("%H:%M:%S")],
            "End Time": [end_time.strftime("%H:%M:%S")],
            "Duration": [duration_str],
            "Task Type": [self.task_type],
            "Details": [self.details]
        }
        new_data_df = pd.DataFrame(data)

        file_path = "tasks.xlsx"

        try:
            if os.path.exists(file_path):
                existing_df = pd.read_excel(file_path)
                updated_df = pd.concat([existing_df, new_data_df], ignore_index=True)
            else:
                updated_df = new_data_df

            updated_df.to_excel(file_path, index=False)
            self.statusBar().showMessage("Task stopped and recorded.")
        except PermissionError:
            self.statusBar().showMessage("Permission denied: Unable to access tasks.xlsx.")
        finally:
            self.start_time = None

    def change_task(self):
        self.stop_task()  # Stop the current task before changing
        self.start_task()  # Start a new task

    def open_log(self):
        file_path = "tasks.xlsx"
        if os.path.exists(file_path):
            subprocess.Popen(['start', file_path], shell=True)
        else:
            self.statusBar().showMessage("Log file not found.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TaskApp()
    window.show()
    sys.exit(app.exec_())
