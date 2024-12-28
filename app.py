import sys
import requests
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QListWidgetItem, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal, QTimer
from design import Ui_MainWindow


class YouTubeSongManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect buttons and enter key to search
        self.ui.search_button.clicked.connect(self.start_search)
        self.ui.add_to_excel_button.clicked.connect(self.add_to_excel)

        # Connect enter key to search method
        self.ui.search_input.returnPressed.connect(self.start_search)

        # List to store search results
        self.search_results = []

        # Initialize the custom spinner (QPainter) for the search button
        self.spinner_angle = 0  # Starting angle for the spinner
        self.spinner_timer = QTimer(self)
        self.spinner_timer.timeout.connect(self.update_spinner)

        # Set focus on search input field once the window opens
        self.ui.search_input.setFocus()

        # Initialize the thread as None (to keep track of it)
        self.search_thread = None

    def start_search(self):
        # Disable the button during search
        self.ui.search_button.setEnabled(False)

        # Start the spinner animation
        self.spinner_timer.start(50)  # Update spinner every 50ms

        # Run the search process in a separate thread to avoid freezing the UI
        if self.search_thread is not None and self.search_thread.isRunning():
            self.search_thread.terminate()  # Terminate any running thread if it's still active

        self.search_thread = SearchThread(self.ui.search_input.text())
        self.search_thread.results_signal.connect(self.on_search_results)
        self.search_thread.finished.connect(
            self.on_search_finished
        )  # Connect finished signal
        self.search_thread.start()

    def update_spinner(self):
        # Update the angle of the spinner
        self.spinner_angle += 10
        if self.spinner_angle >= 360:
            self.spinner_angle = 0

        # Force a repaint of the search button to show the spinner
        self.ui.search_button.update()

    def on_search_results(self, results):
        # Enable the search button back
        self.ui.search_button.setEnabled(True)

        # Stop the spinner animation
        self.spinner_timer.stop()

        # Clear previous results
        self.ui.results_list.clear()
        self.search_results = results

        # Display the search results
        for result in results:
            list_item = QListWidgetItem(result["title"])
            self.ui.results_list.addItem(list_item)

    def on_search_finished(self):
        # This function is called when the search thread finishes
        print("Search completed")

    def add_to_excel(self):
        selected_items = self.ui.results_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select at least one song!")
            return

        # Use the existing Excel file
        excel_file = "songs.xlsx"  # Use the uploaded file path
        try:
            df = pd.read_excel(excel_file, sheet_name="Active")
        except FileNotFoundError:
            QMessageBox.critical(
                self, "Error", "Excel file not found! Please ensure it exists."
            )
            return

        # Add selected items to the DataFrame
        for item in selected_items:
            for result in self.search_results:
                if result["title"] == item.text():
                    # Append to the existing DataFrame
                    new_row = pd.DataFrame(
                        {"Title": [result["title"]], "YouTube Link": [result["url"]]}
                    )
                    df = pd.concat([df, new_row], ignore_index=True)
                    break

        # Save DataFrame to Excel while maintaining the current structure
        try:
            with pd.ExcelWriter(
                excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay"
            ) as writer:
                df.to_excel(writer, index=False, sheet_name="Active")
            QMessageBox.information(
                self, "Success", "Songs added to Excel successfully!"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save to Excel: {e}")


class SearchThread(QThread):
    results_signal = pyqtSignal(list)

    def __init__(self, search_query):
        super().__init__()
        self.search_query = search_query

    def run(self):
        API_KEY = "AIzaSyDmXg_MlBEvUb3oAtMpj-fi4Fet80b21fM"
        url = f"https://www.googleapis.com/youtube/v3/search?part=snippet&q={self.search_query}&type=video&maxResults=15&key={API_KEY}"

        try:
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()

            results = []
            for item in data.get("items", []):
                video_title = item["snippet"]["title"]
                video_id = item["id"]["videoId"]
                video_url = f"https://www.youtube.com/watch?v={video_id}"

                results.append({"title": video_title, "url": video_url})

            # Emit results signal to update the UI
            self.results_signal.emit(results)

        except requests.exceptions.RequestException as e:
            self.results_signal.emit([])  # If error, return empty list


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = YouTubeSongManager()
    window.show()
    sys.exit(app.exec())
