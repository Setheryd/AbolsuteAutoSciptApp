import win32com.client as win32  # type:ignore
import os
import sys
import subprocess
from datetime import datetime, timedelta
from PySide6.QtWidgets import (  # type: ignore
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QMessageBox,
    QHBoxLayout,
    QSizePolicy,
    QScrollArea,
    QLabel,
    QTextEdit,
    QSpacerItem,
    QFrame,
    QFileDialog,
    QTableWidget,
    QTabWidget,
    QTableWidgetSelectionRange,
    QComboBox,
    QSpinBox,
    QTableWidgetItem,
    QKeySequenceEdit,
    QSplitter,
)

from PySide6.QtCore import Qt, QProcess, Slot, QTimer  # type: ignore
from PySide6.QtGui import QMovie, QIcon, QPixmap, QPainter, QColor, QGuiApplication  # type: ignore
import pandas as pd
from io import StringIO
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import mplcursors # type: ignore
import matplotlib.pyplot as plt

# Categorized items (already sorted alphabetically)
daily_items = sorted([
    "Employee Birthday Email"
])
daily_items.insert(0, "Run All")  # Ensure "Run All" is at the top

weekly_items = sorted([
    "Caregiver ID Exp",
    "IN Emp EVAL EXP",
    "IN Emp In-Services EXP",
    "IN PAT SUP EXP",
    "Pending Admission",
    "Pending Caregiver Assignment",
    "Pending IHCC Admission",
    "Pending PERS Installation",
    "SB EMP EVAL EXP",
    "SB Emp Inservices Exp",
    "SB ID EXP",
    "SB PAT SUP EXP"
])
weekly_items.insert(0, "Run All")  # Ensure "Run All" is at the top

monthly_items = sorted([
    "Age Notification",
    "Employee Attrition",
    "Expired NOAs",
    "Inventory Request",
    "Next Months Expired NOAs",
    "Patient Attrition"
])
monthly_items.insert(0, "Run All")  # Ensure "Run All" is at the top


class ScriptButtonWidget(QWidget):
    def __init__(self, script_name, button, has_status=True):
        super().__init__()
        self.script_name = script_name
        self.button = button
        self.has_status = has_status

        # Status label
        self.status_label = QLabel()
        if self.has_status:
            self.status_label.setFixedSize(24, 24)  # Increased size for better visibility
            self.status_label.setAlignment(Qt.AlignCenter)
        else:
            self.status_label.hide()

        if self.has_status:
            # Initial status: Pending
            self.status = 'pending'
            self.update_status_icon()

        # Layout
        layout = QHBoxLayout()
        layout.addWidget(self.button)
        if self.has_status:
            layout.addWidget(self.status_label)
        layout.setAlignment(Qt.AlignLeft)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.setLayout(layout)

    def make_pixmap_white(self, pixmap):
        """
        Returns a white-tinted version of the given pixmap.
        """
        white_pixmap = QPixmap(pixmap.size())
        white_pixmap.fill(Qt.transparent)
        painter = QPainter(white_pixmap)
        painter.setCompositionMode(QPainter.CompositionMode_Source)
        painter.drawPixmap(0, 0, pixmap)
        painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
        painter.fillRect(white_pixmap.rect(), QColor('white'))
        painter.end()
        return white_pixmap

    def update_status_icon(self):
        if not self.has_status:
            return

        # Construct absolute path to the icon
        script_dir = os.path.dirname(os.path.abspath(__file__))
        if self.status == 'pending':
            icon_path = os.path.join(script_dir, 'resources', 'pending.png')
        elif self.status == 'running':
            icon_path = os.path.join(script_dir, 'resources', 'running.png')
        elif self.status == 'success':
            icon_path = os.path.join(script_dir, 'resources', 'success.png')
        elif self.status == 'failed':
            icon_path = os.path.join(script_dir, 'resources', 'failed.png')
        else:
            icon_path = ''

        if os.path.exists(icon_path):
            pixmap = QPixmap(icon_path)
            if pixmap.isNull():
                print(f"Failed to load pixmap from {icon_path}")
                self.status_label.setText(self.status.capitalize())
            else:
                # Tint 'pending.png' and 'running.png' to white
                if self.status in ['pending', 'running']:
                    pixmap = self.make_pixmap_white(pixmap)

                self.status_label.setPixmap(pixmap.scaled(
                    self.status_label.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                ))
        else:
            print(f"Icon path does not exist: {icon_path}")
            self.status_label.setText(self.status.capitalize())

    def set_status(self, status):
        self.status = status
        self.update_status_icon()

class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig, self.ax = plt.subplots(figsize=(width, height), dpi=dpi)
        super(MplCanvas, self).__init__(self.fig)

class CustomTableWidget(QTableWidget):
    """
    Custom QTableWidget to handle Ctrl+C for copying selected data.
    """
    
    def keyPressEvent(self, event):
        """Handle key press events, including Ctrl+C."""
        if event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:  # Check for Ctrl+C
            self.copy_selected_data()
        else:
            super().keyPressEvent(event)  # Pass event to the base class

    def copy_selected_data(self):
        """
        Copy selected cells from QTableWidget to clipboard.
        """
        # Get the selected ranges from the QTableWidget itself
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "No cells are selected to copy.")
            return

        # Get the range of selected cells
        clipboard_content = ""
        for selected_range in selected_ranges:
            rows = range(selected_range.topRow(), selected_range.bottomRow() + 1)
            cols = range(selected_range.leftColumn(), selected_range.rightColumn() + 1)

            for row in rows:
                row_content = []
                for col in cols:
                    item = self.item(row, col)
                    row_content.append(item.text() if item else '')  # Handle empty cells
                clipboard_content += "\t".join(row_content) + "\n"

        # Set clipboard content
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_content.strip())  # Strip trailing newlines

        QMessageBox.information(self, "Copied", "Selected data has been copied to the clipboard.")


class MainApp(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the window properties
        self.setWindowTitle("Absolute Caregivers Auto Scripting App")
        # Get the available screen geometry, excluding taskbars or docks
        screen = QGuiApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()

        # Set window size to slightly less than the available screen size to fit properly
        # Subtract a few pixels (like 10-15) to make sure everything fits on screen
        self.setGeometry(0, 0, screen_geometry.width(), screen_geometry.height() - 20)



        # Create a QTabWidget for switching between tabs
        self.tab_widget = QTabWidget(self)

        # Primary Scripts Tab
        self.scripts_tab = QWidget()
        self.setup_scripts_tab()
        self.tab_widget.addTab(self.scripts_tab, "Scripts")

        # Dashboard Tab (if needed)
        self.secondary_tab = QWidget()
        self.dashboard_tab()
        self.tab_widget.addTab(self.secondary_tab, "Dashboard")

        # Graph Tab (if needed)
        self.graph_tab_widget = QWidget()
        self.graph_tab()
        self.tab_widget.addTab(self.graph_tab_widget, "Graph")

        # Main layout to hold the QTabWidget
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.tab_widget)
        self.setLayout(main_layout)

        # Initialize script mapping and other necessary variables
        #self.initialize_script_mapping()

        # Initialize QProcess
        self.process = None

        # Initialize selected subcategory button
        self.selected_subcategory_button = None

        # Timer for timeout
        self.timer = QTimer(self)
        self.timer.setInterval(60000)  # 60 seconds
        self.timer.timeout.connect(self.handle_timeout)

    def setup_scripts_tab(self):
        """Set up the layout for the Scripts tab."""
        # Main layout for the Scripts tab
        scripts_layout = QHBoxLayout(self.scripts_tab)
        scripts_layout.setContentsMargins(20, 20, 20, 20)
        scripts_layout.setSpacing(20)

        # Left-side layout for category buttons (top-aligned)
        self.category_layout = QVBoxLayout()
        self.category_layout.setSpacing(10)
        self.category_layout.setAlignment(Qt.AlignTop)

        # "Categories" header
        categories_header = QLabel("Categories")
        categories_header.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #A0C4FF;
            }
        """)
        self.category_layout.addWidget(categories_header, alignment=Qt.AlignCenter)

        # Create main category buttons as toggles
        self.daily_button = self.create_category_button("Daily", daily_items, "daily")
        self.weekly_button = self.create_category_button("Weekly", weekly_items, "weekly")
        self.monthly_button = self.create_category_button("Monthly", monthly_items, "monthly")

        self.category_layout.addWidget(self.daily_button)
        self.category_layout.addWidget(self.weekly_button)
        self.category_layout.addWidget(self.monthly_button)

        # Container for category buttons
        self.category_container = QWidget()
        self.category_container.setLayout(self.category_layout)
        self.category_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        scripts_layout.addWidget(self.category_container)

        # Middle layout for the script buttons
        self.script_layout = QVBoxLayout()
        self.script_layout.setSpacing(10)
        self.script_layout.setAlignment(Qt.AlignTop)

        # "Scripts" header
        scripts_header = QLabel("Scripts")
        scripts_header.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #A0C4FF; 
                margin-top: 6px;
            }
        """)
        self.script_layout.addWidget(scripts_header, alignment=Qt.AlignCenter)

        # Container for sub-category buttons
        self.button_container = QWidget()
        self.button_layout = QVBoxLayout()
        self.button_layout.setContentsMargins(0, 0, 0, 0)
        self.button_layout.setSpacing(10)
        self.button_layout.setAlignment(Qt.AlignTop)
        self.button_container.setLayout(self.button_layout)

        # Scroll area to hold the sub-category buttons
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.button_container)
        self.scroll_area.setMinimumWidth(350)
        self.script_layout.addWidget(self.scroll_area)

        # Add the script layout to the main scripts_layout
        scripts_layout.addLayout(self.script_layout)

        # Indicator layout (Rightmost column)
        self.indicator_layout = QVBoxLayout()
        self.indicator_layout.setAlignment(Qt.AlignTop)
        self.indicator_layout.setSpacing(10)

        # Animation Label
        self.animation_label = QLabel(self)
        self.animation_label.setFixedSize(100, 100)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        loading_gif_path = os.path.join(script_dir, "resources", "loading.gif")
        if not os.path.exists(loading_gif_path):
            print(f"Loading GIF not found at {loading_gif_path}. Please ensure the file exists.")
        self.animation_movie = QMovie(loading_gif_path)
        self.animation_label.setMovie(self.animation_movie)
        self.animation_label.setAlignment(Qt.AlignCenter)
        self.animation_label.setScaledContents(True)
        self.animation_label.hide()  # Hide initially

        # Cancel Button
        self.cancel_button = QPushButton("Cancel", self)
        self.cancel_button.setFixedWidth(250)
        self.cancel_button.clicked.connect(self.cancel_process)
        self.cancel_button.hide()  # Hide initially

        # Log Area
        self.log_label = QLabel("Script Output:")
        self.log_label.hide()  # Hide initially
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setFixedWidth(400)
        self.log_text.setFixedHeight(400)
        self.log_text.hide()  # Hide initially
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2e2e2e;
                color: white;
                font-family: Consolas;
                font-size: 12pt;
            }
        """)

        # Add to indicator layout
        self.indicator_layout.addWidget(self.animation_label, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.cancel_button, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.log_label, alignment=Qt.AlignLeft)
        self.indicator_layout.addWidget(self.log_text, alignment=Qt.AlignCenter)

        # Spacer to push indicators to the top
        self.indicator_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Container for indicators
        self.indicator_container = QWidget()
        self.indicator_container.setLayout(self.indicator_layout)
        self.indicator_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        scripts_layout.addWidget(self.indicator_container)

        # State to track expanded categories
        self.expanded_categories = {
            "Daily": False,
            "Weekly": False,
            "Monthly": False
        }

        # Dictionary to store script buttons (if needed for further management)
        self.script_buttons = {}

        # Update styles initially
        self.update_category_styles()
        # # Adjust size to fit the content
        # self.adjustSize()

        # Dictionary to map script names to their paths
        self.scripts = {
            # Daily Scripts
            "Employee Birthday Email": os.path.join(script_dir, "daily_tasks", "birthday.py"),
            # Weekly Scripts
            "Caregiver ID Exp": os.path.join(script_dir, "weekly_tasks", "in_emp_id_exp.py"),
            "IN Emp EVAL EXP": os.path.join(script_dir, "weekly_tasks", "indy_emp_eval.py"),
            "IN Emp In-Services EXP": os.path.join(script_dir, "weekly_tasks", "in_emp_inservices_exp.py"),
            "IN PAT SUP EXP": os.path.join(script_dir, "weekly_tasks", "in_pat_sup_exp.py"),
            "Pending Admission": os.path.join(script_dir, "weekly_tasks", "pending_admission.py"),
            "Pending Caregiver Assignment": os.path.join(script_dir, "weekly_tasks", "pending_caregiver_assignment.py"),
            "Pending IHCC Admission": os.path.join(script_dir, "weekly_tasks", "pending_IHCC_admission.py"),
            "Pending PERS Installation": os.path.join(script_dir, "weekly_tasks", "pending_PERS_Installation.py"),
            "SB EMP EVAL EXP": os.path.join(script_dir, "weekly_tasks", "sb_emp_eval.py"),
            "SB Emp Inservices Exp": os.path.join(script_dir, "weekly_tasks", "sb_emp_inservices_exp.py"),
            "SB ID EXP": os.path.join(script_dir, "weekly_tasks", "sb_emp_id_exp.py"),
            "SB PAT SUP EXP": os.path.join(script_dir, "weekly_tasks", "sb_pat_sup_exp.py"),
            # Monthly Scripts
            "Age Notification": os.path.join(script_dir, "monthly_tasks", "age.py"),
            "Expired NOAs": os.path.join(script_dir, "monthly_tasks", "NOA_exp.py"),
            "Next Months Expired NOAs": os.path.join(script_dir, "monthly_tasks", "next_month_NOA_exp.py"),
        }

        # Dictionary to store script buttons
        self.script_buttons = {}

        # Dictionary to store script execution results
        self.script_results = {}

        # Timer for timeout
        self.timer = QTimer(self)
        self.timer.setInterval(60000)  # 45 seconds
        self.timer.timeout.connect(self.handle_timeout)
        
    def dashboard_tab(self):
        """Set up the layout for the secondary tab."""
        self.secondary_layout = QVBoxLayout(self.secondary_tab)

        # QSplitter for resizable columns
        splitter = QSplitter(Qt.Horizontal)

        # Left-side layout for the custom program buttons
        self.file_list_layout = QVBoxLayout()
        self.file_list_layout.setAlignment(Qt.AlignTop)

        # Define custom labels for buttons and the corresponding programs/scripts to run
        custom_programs = {
            "Active Patients per Month": "Admission_by_Month.py",
            "View Patient Data": "Patient_Data.py",
            "Active Contractor per Month": "Caregiver_by_Month.py",
            "Generate Report 4": "Report4.py",
            "Generate Report 5": "Report5.py",
            "Run Custom Report 6": "Report6.py"
        }

        # Container widget to hold the buttons
        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)

        # Create buttons with custom labels to run the corresponding programs
        for button_label, program in custom_programs.items():
            button = QPushButton(button_label)
            button.clicked.connect(lambda checked, prog=program: self.run_python_script(prog))
            button_layout.addWidget(button)

        # Ensure the button container has the right size to show all buttons
        button_container.setLayout(button_layout)
        button_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        button_container.adjustSize()

        # Scroll area to make the button list scrollable
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(button_container)

        # Create a widget for the left-side content (buttons)
        left_widget = QWidget()
        left_widget.setLayout(QVBoxLayout())
        left_widget.layout().addWidget(scroll_area)

        # Add the left widget to the splitter
        splitter.addWidget(left_widget)

        # Right-side layout for displaying the DataFrame (QTableWidget)
        self.table_widget = CustomTableWidget()
        self.table_widget.setColumnCount(0)
        self.table_widget.setRowCount(0)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                background-color: #f0f0f0;
                color: black;
            }
        """)

        # Add the table widget to the splitter
        splitter.addWidget(self.table_widget)

        # Set initial sizes for the splitter (left size for buttons, right size for table)
        splitter.setSizes([300, 700])  # Adjust these values as necessary

        # Add the splitter to the secondary layout and give it a stretch factor
        self.secondary_layout.addWidget(splitter)
        self.secondary_layout.setStretch(0, 1)  # Make sure the splitter takes up most space

        # Horizontal layout for the bottom buttons (Save DataFrame, Copy Selected)
        bottom_buttons_layout = QHBoxLayout()
        bottom_buttons_layout.setAlignment(Qt.AlignCenter)

        # Buttons for saving and copying DataFrame
        self.save_dataframe_button = QPushButton("Save DataFrame")
        self.save_dataframe_button.clicked.connect(self.save_dataframe)

        self.copy_selected_button = QPushButton("Copy Selected")
        self.copy_selected_button.clicked.connect(self.table_widget.copy_selected_data)

        # Add the buttons to the bottom button layout
        bottom_buttons_layout.addWidget(self.save_dataframe_button)
        bottom_buttons_layout.addWidget(self.copy_selected_button)

        # Add the bottom buttons layout to the secondary layout
        self.secondary_layout.addLayout(bottom_buttons_layout)

        # Set the layout to the secondary tab
        self.secondary_tab.setLayout(self.secondary_layout)




        
    def graph_tab(self):
        """Set up the layout for the graph tab."""
        self.graph_layout = QVBoxLayout(self.graph_tab_widget) 

        # Create dropdowns for X-axis and Y-axis selection
        self.x_axis_combo = QComboBox(self)
        self.y_axis_combo = QComboBox(self)

        # Add a label for X-axis
        self.x_axis_label = QLabel("Select X-axis:")
        self.graph_layout.addWidget(self.x_axis_label)
        self.graph_layout.addWidget(self.x_axis_combo)

        # Add a label for Y-axis
        self.y_axis_label = QLabel("Select Y-axis:")
        self.graph_layout.addWidget(self.y_axis_label)
        self.graph_layout.addWidget(self.y_axis_combo)

        # Add a spinbox to control the number of labels on the X-axis
        self.label_spinbox = QSpinBox(self)
        self.label_spinbox.setRange(1, 20)  # Set a reasonable range for label steps
        self.label_spinbox.setValue(6)  # Default value for step count
        self.label_spinbox.setSuffix(" labels")  # Add a label suffix

        self.label_spinbox_label = QLabel("Number of X-axis labels to display:")
        self.graph_layout.addWidget(self.label_spinbox_label)
        self.graph_layout.addWidget(self.label_spinbox)
        
        # Add a button to plot the graph
        self.plot_button = QPushButton("Plot Graph", self)
        self.plot_button.clicked.connect(self.plot_graph)
        self.graph_layout.addWidget(self.plot_button)

        # Set up the matplotlib canvas
        self.canvas = FigureCanvas(Figure(figsize=(5, 3)))
        self.graph_layout.addWidget(self.canvas)

        # Initialize an empty Axes
        self.canvas.ax = self.canvas.figure.add_subplot(111)

        self.graph_tab_widget.setLayout(self.graph_layout)

    
    def plot_graph(self):
        """Plot the graph with the selected X and Y axis, and enable click to show data values."""
        # Clear the previous plot
        self.canvas.ax.clear()

        # Get the selected X and Y columns
        x_column = self.x_axis_combo.currentText()
        y_column = self.y_axis_combo.currentText()

        # Get the number of labels to show from the spinbox
        max_labels = self.label_spinbox.value()

        # Plot the graph
        x_data = self.df[x_column]
        y_data = self.df[y_column]

        # Scatter plot with labels
        scatter_plot = self.canvas.ax.scatter(x_data, y_data, picker=True)

        # Set axis labels
        self.canvas.ax.set_xlabel(x_column)
        self.canvas.ax.set_ylabel(y_column)

        # Ensure proper tick placement for X-axis labels
        num_points = len(x_data)
        step = max(1, num_points // max_labels)  # Calculate step based on number of labels

        # Set X-axis ticks and labels
        tick_indices = range(0, num_points, step)
        self.canvas.ax.set_xticks([x_data[i] for i in tick_indices])
        self.canvas.ax.set_xticklabels([x_data[i] for i in tick_indices], ha="right")
        
        self.canvas.figure.tight_layout()

        # Redraw the canvas
        self.canvas.draw()

        # Connect the click event to show data point value
        self.canvas.mpl_connect("pick_event", self.on_click)

    def on_click(self, event):
        """Handle click events on the plot and display the clicked data point value."""
        if event.artist != self.canvas.ax.collections[0]:
            return

        # Get the index of the clicked data point
        ind = event.ind[0]
        x_column = self.x_axis_combo.currentText()
        y_column = self.y_axis_combo.currentText()

        # Get the data point values
        x_value = self.df[x_column].iloc[ind]
        y_value = self.df[y_column].iloc[ind]

        # Show the data point value in a message box
        QMessageBox.information(self, "Data Point", f"X: {x_value}\nY: {y_value}")
        
    def create_scrollable_file_list(self, file_names):
        """Create a scrollable area containing buttons for the given file names."""
        # Container widget for the file buttons
        container_widget = QWidget()
        layout = QVBoxLayout(container_widget)

        # Add file buttons to the layout
        for file_name in file_names:
            button = QPushButton(file_name)
            button.clicked.connect(lambda checked, file=file_name: self.run_python_script(file))
            layout.addWidget(button)

        # Scroll area to make the file list scrollable
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container_widget)
        
        # Set a fixed width for the scroll area and allow vertical expansion
        # scroll_area.setFixedWidth(400)  # Set the fixed width as needed
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Fixed width, expands vertically

        return scroll_area

    def display_dataframe(self, df):
        """
        Display a Pandas DataFrame in the QTableWidget on the dashboard.
        """
        if df.empty:
            QMessageBox.warning(self, "No Data", "The DataFrame is empty.")
            return

        self.df = df  # Store the DataFrame for future export
        self.table_widget.clear()

        # Set row and column count based on the DataFrame
        self.table_widget.setRowCount(df.shape[0])
        self.table_widget.setColumnCount(df.shape[1])

        # Set the headers
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        # Populate the table with the DataFrame data
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                self.table_widget.setItem(i, j, item)

        # Resize the columns to fit the contents
        self.table_widget.resizeColumnsToContents()

        # Populate the X and Y axis combo boxes with column names in the third tab
        self.x_axis_combo.clear()
        self.y_axis_combo.clear()

        for col in df.columns:
            self.x_axis_combo.addItem(col)
            self.y_axis_combo.addItem(col)

    def save_dataframe(self):
        """Open a file dialog to save the DataFrame."""
        if not hasattr(self, 'df'):
            QMessageBox.warning(self, "Error", "No DataFrame to save.")
            return

        # Open a file dialog to choose the save location
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save DataFrame", "", "CSV Files (*.csv);;All Files (*)", options=options)

        if file_name:
            try:
                # Save the DataFrame as a CSV file
                self.df.to_csv(file_name, index=False)
                QMessageBox.information(self, "Success", f"DataFrame saved successfully to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save DataFrame: {str(e)}")

    def run_python_script(self, file_name):
        """Execute the Python script corresponding to the file name and display the DataFrame output."""
        try:
            # Determine the absolute path to the 'scripts' subfolder
            script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_extraction")
            # Path to the Python script
            script_path = os.path.join(script_dir, file_name)

            # Check if the file exists
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found: {file_name}")
                return

            # Run the Python script as a subprocess
            process = subprocess.Popen([sys.executable, script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()

            # Check if the process finished without errors
            if process.returncode == 0:
                # Capture stdout output from the process
                output = stdout.decode().strip()

                # Log the raw output for debugging
                print(f"Raw Output from {file_name}:\n{output}")

                # Check if output contains valid CSV structure
                if ',' in output and '\n' in output:
                    try:
                        # Try to parse the output as CSV into a DataFrame
                        df = pd.read_csv(StringIO(output))  # Read DataFrame from the script output
                        self.display_dataframe(df)  # Display the DataFrame in the dashboard
                    except pd.errors.ParserError as pe:
                        QMessageBox.critical(self, "Error", f"CSV Parsing error: {str(pe)}")
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to parse DataFrame output: {str(e)}")
                else:
                    QMessageBox.critical(self, "Error", f"Output from {file_name} is not in CSV format.")
            else:
                # If the process failed, show the error message
                error_message = stderr.decode().strip()
                print(f"Error Output from {file_name}:\n{error_message}")
                QMessageBox.critical(self, "Error", f"Script '{file_name}' failed to execute.\nError: {error_message}")
        except Exception as e:
            # Handle any other unexpected exceptions
            QMessageBox.critical(self, "Error", f"An error occurred while executing the script: {str(e)}")
  
    def create_category_button(self, text, items, identifier):
        button = QPushButton(text, self)
        button.setFixedWidth(300)  # Set a fixed width for all category buttons
        button.setCheckable(True)  # Make the button checkable
        button.clicked.connect(lambda: self.toggle_category(text, items))
        button.setObjectName(identifier)  # Assign a unique object name
        
        # Remove the individual style sheet
        # button.setStyleSheet("""...""")
        
        # Set size policy to prevent vertical stretching
        button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        
        return button

    def toggle_category(self, category, items):
        if self.expanded_categories[category]:
            # If expanded, collapse and clear sub-category buttons
            self.clear_buttons()
            self.expanded_categories[category] = False
            self.set_category_button_checked(category, False)
            self.current_category = None  # Reset current category
        else:
            # If collapsed, expand and show sub-category buttons
            self.clear_buttons()
            self.current_category = category  # Store current category
            self.show_buttons(items)
            # Update the state to show it's expanded
            self.expanded_categories[category] = True
            # Collapse other categories
            for cat in self.expanded_categories:
                if cat != category:
                    self.expanded_categories[cat] = False
                    self.set_category_button_checked(cat, False)
        # Update button styles to reflect current selection
        self.update_category_styles()
        # Adjust window size based on content
        # self.adjustSize()

    def clear_buttons(self):
        """
        Clear all sub-category buttons from the layout.
        """
        for i in reversed(range(self.button_layout.count())):
            widget = self.button_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        # Also reset the selected subcategory button
        self.selected_subcategory_button = None

        # Clear script_buttons mapping
        self.script_buttons.clear()

    def show_buttons(self, items):
        """
        Create and display sub-category buttons based on the selected category.

        Args:
            items (list): The list of sub-category items to create buttons for.
        """
        for item in items:
            button = QPushButton(item, self)
            button.setFixedWidth(300)  # Set a fixed width for all sub-category buttons

            # Make the button checkable
            button.setCheckable(True)

            # Handle the "Run All" button
            if item == "Run All":
                button.setObjectName("run_all")
                if self.current_category == "Weekly":
                    button.clicked.connect(self.run_all_weekly_items)
                elif self.current_category == "Daily":
                    button.clicked.connect(self.run_all_daily_items)
                elif self.current_category == "Monthly":  # Add this condition
                    button.clicked.connect(self.run_all_monthly_items)
                else:
                    button.clicked.connect(self.show_message)
                # Wrap in ScriptButtonWidget without status indicator
                script_widget = ScriptButtonWidget(item, button, has_status=False)
            else:
                # Map the script name to the button
                if item in self.scripts:
                    script_path = self.scripts[item]
                    button.clicked.connect(lambda checked, path=script_path, name=item: self.run_single_script(path, name))
                else:
                    button.clicked.connect(self.show_message)

                # Wrap in ScriptButtonWidget with status indicator
                script_widget = ScriptButtonWidget(item, button, has_status=True)
                self.script_buttons[item] = script_widget  # Store the composite widget

            # Update stylesheet to handle the checked state
            button.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    border-radius: 8px;
                    padding: 10px;
                    text-align: center;  /* Center text horizontally */
                    color: black;
                }
                QPushButton:hover {
                    background-color: #e6e6e6;
                }
                QPushButton:checked {
                    background-color: #a0c4ff;
                    border: 1px solid #89a4ff;
                    color: black;
                }
            """)

            # Set size policy to prevent vertical stretching
            button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            # Connect the selection handler
            button.clicked.connect(lambda checked, btn=button: self.select_subcategory(btn))

            # Add the widget to the layout with horizontal center alignment
            self.button_layout.addWidget(script_widget, alignment=Qt.AlignLeft)

    def select_subcategory(self, button):
        """
        Handles the selection of a sub-category button.
        Ensures that only one sub-category button is checked at a time.

        Args:
            button (QPushButton): The button that was clicked.
        """
        # If another subcategory button is already selected, uncheck it
        if self.selected_subcategory_button and self.selected_subcategory_button != button:
            self.selected_subcategory_button.setChecked(False)

        # Update the selected subcategory button
        if button.isChecked():
            self.selected_subcategory_button = button
        else:
            self.selected_subcategory_button = None

    def show_message(self):
        """
        Display a message box with the button text.
        """
        button = self.sender()
        if button:
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Button Clicked")
            msg_box.setText(f"You clicked: {button.text()}")
            msg_box.exec()
    def run_all_daily_items(self):
        try:
            # Get the list of script paths for daily scripts
            script_items = daily_items[1:]  # Skip "Run All"
            self.pending_scripts = [(item, self.scripts[item]) for item in script_items if item in self.scripts]

            # Initialize script results
            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            # Reset status icons
            for item in self.script_buttons:
                self.script_buttons[item].set_status('pending')

            # Disable the sub-category buttons
            self.set_buttons_enabled(False)

            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            self.log_text.clear()

            # Start the first script
            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    def run_all_weekly_items(self):
        try:
            # Get the list of script paths for weekly scripts
            script_items = weekly_items[1:]  # Skip "Run All"
            self.pending_scripts = [(item, self.scripts[item]) for item in script_items if item in self.scripts]

            # Initialize script results
            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            # Reset status icons
            for item in self.script_buttons:
                self.script_buttons[item].set_status('pending')

            # Disable the sub-category buttons
            self.set_buttons_enabled(False)
            # Show the animation and logs
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            self.log_text.clear()
            # Start the first script
            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    def run_all_monthly_items(self):
        try:
            # Get the list of script paths for monthly scripts
            script_items = monthly_items[1:]  # Skip "Run All"
            self.pending_scripts = [(item, self.scripts[item]) for item in script_items if item in self.scripts]

            # Initialize script results
            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            # Reset status icons
            for item in self.script_buttons:
                if item in self.scripts:  # Ensure it's a monthly script
                    self.script_buttons[item].set_status('pending')

            # Disable the sub-category buttons
            self.set_buttons_enabled(False)

            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            self.log_text.clear()

            # Start the first script
            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    def run_next_script(self):
        if self.pending_scripts:
            next_script_name, next_script_path = self.pending_scripts.pop(0)
            self.current_script_name = next_script_name
            self.run_script(next_script_path, next_script_name)
        else:
            # All scripts have been executed
            self.log_text.append("<span style='color: green;'>All scripts executed.</span>")
            # Show summary
            self.show_summary()

    def run_script(self, script_path, script_name):
        try:
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                self.run_next_script()
                return

            # Update status to 'running'
            self.script_results[script_name] = 'running'
            self.script_buttons[script_name].set_status('running')

            # Highlight the active button
            self.highlight_active_button(script_name)

            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            self.process.setArguments(['-u', script_path])

            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)

            # Start the timer for timeout
            self.timer.start()

            # Show loading GIF and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()

            self.process.start()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.run_next_script()

    def run_single_script(self, script_path, script_name):
        try:
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return

            # Reset status icons
            self.script_buttons[script_name].set_status('pending')
            self.script_results = {script_name: 'pending'}

            # Disable the sub-category buttons
            self.set_buttons_enabled(False)

            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()

            # Clear previous logs
            self.log_text.clear()

            # Update status to 'running'
            self.script_buttons[script_name].set_status('running')
            self.highlight_active_button(script_name)
            self.current_script_name = script_name

            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])

            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)

            # Start the timer for timeout
            self.timer.start()

            # Start the process
            self.process.start()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred while executing the script:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    @Slot()
    def handle_stdout(self):
        """
        Handle standard output from the running process and append it to the log area.
        """
        if self.process:
            data = bytes(self.process.readAllStandardOutput()).decode('utf-8')
            # Replace newline characters with HTML line breaks for proper formatting
            formatted_data = data.replace('\n', '<br>')
            self.log_text.append(f"<span style='color: white;'>{formatted_data}</span>")

    @Slot()
    def handle_stderr(self):
        """
        Handle standard error from the running process and append it to the log area.
        """
        if self.process:
            data = bytes(self.process.readAllStandardError()).decode('utf-8')
            # Replace newline characters with HTML line breaks for proper formatting
            formatted_data = data.replace('\n', '<br>')
            self.log_text.append(f"<span style='color: red;'>{formatted_data}</span>")

    @Slot(int, QProcess.ExitStatus)
    def process_finished(self, exitCode, exitStatus):
        self.timer.stop()
        if exitCode == 0:
            self.log_text.append(f"<span style='color: green;'>Script '{self.current_script_name}' executed successfully.</span>")
            # Update status to 'success'
            self.script_results[self.current_script_name] = 'success'
            self.script_buttons[self.current_script_name].set_status('success')
        else:
            self.log_text.append(f"<span style='color: red;'>Script '{self.current_script_name}' failed with exit code {exitCode}.</span>")
            # Update status to 'failed'
            self.script_results[self.current_script_name] = 'failed'
            self.script_buttons[self.current_script_name].set_status('failed')

        # Hide loading GIF and cancel button after each script
        self.animation_movie.stop()
        self.animation_label.hide()
        self.cancel_button.hide()
        self.set_buttons_enabled(True)

        if hasattr(self, 'pending_scripts') and self.pending_scripts:
            self.run_next_script()
        else:
            self.log_text.append("<span style='color: green;'>All scripts executed.</span>")
            # Show summary if in "Run All" mode
            if hasattr(self, 'pending_scripts'):
                self.show_summary()

    def show_summary(self):
        summary = "Script Execution Summary:\n\n"
        for script_name, result in self.script_results.items():
            summary += f"{script_name}: {result.capitalize()}\n"

        QMessageBox.information(self, "Summary", summary)

    def cancel_process(self):
        if self.process and self.process.state() == QProcess.Running:
            self.process.kill()
            self.process = None
            self.timer.stop()
            if hasattr(self, 'pending_scripts'):
                self.pending_scripts = []
            self.animation_movie.stop()
            self.animation_label.hide()
            self.cancel_button.hide()
            self.set_buttons_enabled(True)
            self.log_text.append("<span style='color: orange;'>Script execution has been cancelled.</span>")
            # Update status to 'failed'
            self.script_results[self.current_script_name] = 'failed'
            self.script_buttons[self.current_script_name].set_status('failed')
            # Show summary
            if hasattr(self, 'pending_scripts'):
                self.show_summary()

    def set_buttons_enabled(self, enabled):
        """
        Enable or disable all sub-category buttons.

        Args:
            enabled (bool): True to enable, False to disable.
        """
        for i in range(self.button_layout.count()):
            widget = self.button_layout.itemAt(i).widget()
            if isinstance(widget, ScriptButtonWidget):
                widget.button.setEnabled(enabled)
            elif isinstance(widget, QPushButton):
                # For "Run All" buttons
                widget.setEnabled(enabled)

    def update_category_styles(self):
        """
        Update each category button based on its expanded state.
        """
        self.daily_button.setChecked(self.expanded_categories["Daily"])
        self.weekly_button.setChecked(self.expanded_categories["Weekly"])
        self.monthly_button.setChecked(self.expanded_categories["Monthly"])

    def set_category_button_checked(self, category, checked):
        """
        Helper method to set the checked state of a category button.

        Args:
            category (str): The name of the category.
            checked (bool): True to check, False to uncheck.
        """
        if category == "Daily":
            self.daily_button.setChecked(checked)
        elif category == "Weekly":
            self.weekly_button.setChecked(checked)
        elif category == "Monthly":
            self.monthly_button.setChecked(checked)

    def highlight_active_button(self, script_name):
        """
        Highlights the active button with the specified color and updates its status to 'running'.
        """
        # Reset styles for all buttons
        for item in self.script_buttons:
            widget = self.script_buttons[item]
            widget.button.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    border-radius: 8px;
                    padding: 10px;
                    text-align: center;
                    color: black;
                }
                QPushButton:hover {
                    background-color: #e6e6e6;
                }
                QPushButton:checked {
                    background-color: #a0c4ff;
                    border: 1px solid #89a4ff;
                    color: black;
                }
            """)

        # Highlight the active button with #A0C4FF
        widget = self.script_buttons.get(script_name)
        if widget:
            widget.button.setStyleSheet("""
                QPushButton {
                    background-color: #A0C4FF;
                    border: 1px solid #cccccc;
                    border-radius: 8px;
                    padding: 10px;
                    text-align: center;
                    color: black;
                }
                QPushButton:hover {
                    background-color: #89a4ff;
                }
            """)

    def handle_timeout(self):
        if self.process and self.process.state() == QProcess.Running:
            self.process.kill()
            self.process = None
            self.timer.stop()
            if hasattr(self, 'pending_scripts'):
                self.pending_scripts = []
            self.animation_movie.stop()
            self.animation_label.hide()
            self.cancel_button.hide()
            self.set_buttons_enabled(True)
            self.log_text.append(f"<span style='color: orange;'>Script '{self.current_script_name}' timed out.</span>")
            # Update status to 'failed'
            self.script_results[self.current_script_name] = 'failed'
            self.script_buttons[self.current_script_name].set_status('failed')

            # Hide loading GIF and cancel button
            self.animation_movie.stop()
            self.animation_label.hide()
            self.cancel_button.hide()
            self.set_buttons_enabled(True)

            if hasattr(self, 'pending_scripts') and self.pending_scripts:
                self.run_next_script()
            else:
                self.show_summary()


# Main entry point
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Apply the global button style
    button_style = """
    QPushButton {
        background-color: #ffffff;
        border: 2px solid #cccccc;
        border-radius: 10px;
        padding: 10px;
        text-align: center;
        color: black;
        font-weight: bold;
        font-size: 14px;  /* Adjust as needed */
    }
    QPushButton:hover {
        background-color: #dcdcdc;
    }
    QPushButton:checked {
        background-color: #a0c4ff;
        border: 2px solid #89a4ff;
        color: black;
    }
    """
    app.setStyleSheet(button_style)
    
    window = MainApp()
    window.show()
    sys.exit(app.exec())
