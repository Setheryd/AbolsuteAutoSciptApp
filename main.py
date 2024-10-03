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
)
from PySide6.QtCore import Qt, QProcess, Slot, QTimer  # type: ignore
from PySide6.QtGui import QMovie, QIcon, QPixmap, QPainter, QColor  # type: ignore


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


class MainApp(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the window properties
        self.setWindowTitle("Absolute Caregivers Auto Scripting App")
        self.setGeometry(100, 100, 1200, 800)  # Adjusted height for better layout

        # Main layout
        self.main_layout = QHBoxLayout()
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(20)

        # Left-side layout for category buttons (top-aligned)
        self.category_layout = QVBoxLayout()
        self.category_layout.setContentsMargins(20, 20, 20, 20)
        self.category_layout.setSpacing(20)
        self.category_layout.setAlignment(Qt.AlignTop)  # Ensure top alignment

        # Add "Categories" header with additional styling
        categories_header = QLabel("Categories")
        categories_header.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #A0C4FF;  
                padding: 5px;
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

        # Create a widget to hold the category buttons and add it to the left side
        self.category_container = QWidget()
        self.category_container.setLayout(self.category_layout)
        self.category_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.main_layout.addWidget(self.category_container)

        # Middle layout (right_layout)
        self.right_layout = QVBoxLayout()
        self.right_layout.setContentsMargins(0, 0, 0, 0)
        self.right_layout.setSpacing(10)
        self.right_layout.setAlignment(Qt.AlignTop)

        # Add "Scripts" header with additional styling
        scripts_header = QLabel("Scripts")
        scripts_header.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #A0C4FF; 
                padding: 5px;
            }
        """)
        self.right_layout.addWidget(scripts_header, alignment=Qt.AlignCenter)

        # Right-side container for sub-category buttons (top-aligned)
        self.button_container = QWidget()
        self.button_layout = QVBoxLayout()
        self.button_layout.setContentsMargins(0, 0, 0, 0)
        self.button_layout.setSpacing(10)
        self.button_layout.setAlignment(Qt.AlignTop)  # Ensure top alignment
        self.button_container.setLayout(self.button_layout)

        # Create a scroll area to hold the sub-category buttons
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.button_container)
        self.scroll_area.setMinimumWidth(350)  # Increased width to accommodate status icons

        # Add the scroll area to the right layout
        self.right_layout.addWidget(self.scroll_area)

        # Create a widget to hold the right_layout
        self.right_container = QWidget()
        self.right_container.setLayout(self.right_layout)
        self.right_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.main_layout.addWidget(self.right_container)

        # Indicator layout (Rightmost column)
        self.indicator_layout = QVBoxLayout()
        self.indicator_layout.setAlignment(Qt.AlignTop)
        self.indicator_layout.setSpacing(10)

        # Animation Label
        self.animation_label = QLabel(self)
        self.animation_label.setFixedSize(100, 100)  # Increased size for better visibility
        script_dir = os.path.dirname(os.path.abspath(__file__))
        loading_gif_path = os.path.join(script_dir, "resources", "loading.gif")
        if not os.path.exists(loading_gif_path):
            print(f"Loading GIF not found at {loading_gif_path}. Please ensure the file exists.")
        self.animation_movie = QMovie(loading_gif_path)
        self.animation_label.setMovie(self.animation_movie)
        self.animation_label.setAlignment(Qt.AlignCenter)  # Center the GIF within the label
        self.animation_label.setScaledContents(True)  # Scale the GIF to fit the label
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

        # Add a spacer to push indicators to the top
        self.indicator_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Create a widget to hold the indicator_layout
        self.indicator_container = QWidget()
        self.indicator_container.setLayout(self.indicator_layout)
        self.indicator_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.main_layout.addWidget(self.indicator_container)

        # Set the main layout for the widget
        self.setLayout(self.main_layout)

        # State to track expanded categories
        self.expanded_categories = {
            "Daily": False,
            "Weekly": False,
            "Monthly": False
        }

        # Initialize QProcess
        self.process = None

        # Initialize selected subcategory button
        self.selected_subcategory_button = None

        # Update styles initially
        self.update_category_styles()

        # Adjust size to fit the content
        self.adjustSize()

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

    def create_category_button(self, text, items, identifier):
        """
        Create a category button that acts as a toggle.

        Args:
            text (str): The display text of the button.
            items (list): The list of sub-category items associated with this category.
            identifier (str): A unique identifier for the button (used as objectName).

        Returns:
            QPushButton: The configured category button.
        """
        button = QPushButton(text, self)
        button.setFixedWidth(300)  # Set a fixed width for all category buttons
        button.setCheckable(True)  # Make the button checkable
        button.clicked.connect(lambda: self.toggle_category(text, items))
        button.setObjectName(identifier)  # Assign a unique object name

        # Enhance button appearance with border-radius and black text
        button.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 2px solid #cccccc;
                border-radius: 10px;
                padding: 10px;
                text-align: center;  /* Center text horizontally */
                color: black;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #dcdcdc;
            }
            QPushButton:checked {
                background-color: #a0c4ff;
                border: 2px solid #89a4ff;
                color: black;
            }
        """)

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
        self.adjustSize()

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
            self.button_layout.addWidget(script_widget, alignment=Qt.AlignHCenter)

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
    window = MainApp()
    window.show()
    sys.exit(app.exec())
