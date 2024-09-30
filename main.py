# main_app.py

import win32com.client as win32# type: ignore
import os
import sys
import subprocess
from datetime import datetime, timedelta
from PySide6.QtWidgets import (# type: ignore
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
    QSpacerItem
)
from PySide6.QtCore import Qt, QProcess, Slot # type: ignore
from PySide6.QtGui import QMovie # type: ignore

# Categorized items (already sorted alphabetically)
daily_items = sorted([
    "Employee Birthday Email"
])
daily_items.insert(0, "Run All")  # Ensure "Run All" is at the top

weekly_items = sorted([
    "Caregiver ID Exp",
    "Clinician Info Source Pull",
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


class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # Set up the window properties
        self.setWindowTitle("Absolute Caregivers Auto Scripting App")
        self.setGeometry(100, 100, 1200, 800)  # Increased width for better layout
        
        # Main layout
        self.main_layout = QHBoxLayout()
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        self.main_layout.setSpacing(20)
        
        # Left-side layout for category buttons (top-aligned)
        self.category_layout = QVBoxLayout()
        self.category_layout.setContentsMargins(0, 0, 0, 0)
        self.category_layout.setSpacing(10)
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
        self.category_container.setFixedWidth(200)  # Set fixed width for the category container
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
        self.scroll_area.setFixedWidth(350)
        self.scroll_area.setFixedHeight(600)

        # Add the scroll area to the right layout
        self.right_layout.addWidget(self.scroll_area)
        
        # Create a widget to hold the right_layout and set fixed width
        self.right_container = QWidget()
        self.right_container.setLayout(self.right_layout)
        self.right_container.setFixedWidth(400)  # Adjust as needed
        self.main_layout.addWidget(self.right_container)

        # Indicator layout (Rightmost column)
        self.indicator_layout = QVBoxLayout()
        self.indicator_layout.setAlignment(Qt.AlignTop)
        self.indicator_layout.setSpacing(10)
        
        # Animation Label
        self.animation_label = QLabel(self)
        self.animation_label.setFixedSize(100, 100)  # Increased size for better visibility
        loading_gif_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "loading.gif")
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
        
        # Create a widget to hold the indicator_layout and set fixed width
        self.indicator_container = QWidget()
        self.indicator_container.setLayout(self.indicator_layout)
        self.indicator_container.setFixedWidth(350)  # Adjust as needed
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
        button.setFixedWidth(150)  # Set a fixed width for all category buttons
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
        """
        Toggle the visibility of sub-category buttons based on the selected category.
        
        Args:
            category (str): The name of the category being toggled.
            items (list): The list of sub-category items to display if expanded.
        """
        if self.expanded_categories[category]:
            # If expanded, collapse and clear sub-category buttons
            self.clear_buttons()
            self.expanded_categories[category] = False
            self.set_category_button_checked(category, False)
        else:
            # If collapsed, expand and show sub-category buttons
            self.clear_buttons()
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
    
    # Inside the show_buttons method in MainApp class

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
            
            # Connect specific actions based on the button text or object name
            if item == "Caregiver ID Exp":
                button.setObjectName("caregiver_id_exp")  # Assign a unique object name
                button.clicked.connect(self.run_caregiver_id_exp)
            elif item == "IN Emp EVAL EXP":
                button.setObjectName("in_emp_eval_exp")  # Assign a unique object name
                button.clicked.connect(self.run_indy_emp_eval)
            elif item == "SB EMP EVAL EXP":  # New condition added
                button.setObjectName("sb_emp_eval_exp")  # Assign a unique object name
                button.clicked.connect(self.run_sb_emp_eval)  # Connect to existing method
            elif item == "SB ID EXP":
                button.setObjectName("sb_id_exp")
                button.clicked.connect(self.run_sb_id_exp)
            elif item == "Pending Admission":
                button.setObjectName("pending_admission")
                button.clicked.connect(self.run_pending_admission)
            elif item == "Pending Caregiver Assignment":
                button.setObjectName("pending_caregiver_assignment")
                button.clicked.connect(self.run_pending_caregiver_assignment)
            elif item == "Pending IHCC Admission":
                button.setObjectName("pending_IHCC_admission")
                button.clicked.connect(self.run_pending_IHCC_admission)
            elif item == "Pending PERS Installation":
                button.setObjectName("pending_PERS_installation")  # Assign a unique object name
                button.clicked.connect(self.run_pending_PERS_installation)
            elif item == "SB ID EXP":  # Add this condition
                button.setObjectName("sb_id_exp")
                button.clicked.connect(self.run_sb_id_exp)
            else:
                button.clicked.connect(self.show_message)
            
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
            
            # Add the button to the layout with horizontal center alignment
            self.button_layout.addWidget(button, alignment=Qt.AlignHCenter)



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
    
    def run_caregiver_id_exp(self):
        """
        Execute the in_emp_id_exp.py script when the "Caregiver ID Exp" button is clicked.
        """
        try:
            # Determine the path to the in_care_id_exp.py script
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "weekly_tasks", "in_emp_id_exp.py")
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred while executing the script:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

            # Inside the MainApp class in main_app.py

    def run_sb_id_exp(self):
        """
        Execute the sb_emp_id_exp.py script when the "SB ID EXP" button is clicked.
        """
        try:
            # Determine the path to the sb_emp_id_exp.py script
            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 
                "weekly_tasks", 
                "sb_emp_id_exp.py"
            )
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at '{script_path}'")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Exception", 
                f"An error occurred while executing the script:\n{str(e)}"
            )
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    
    def run_indy_emp_eval(self):
        """
        Execute the indy_emp_eval.py script when the "IN Emp EVAL EXP" button is clicked.
        """
        try:
            # Determine the path to the indy_emp_eval.py script
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "weekly_tasks", "indy_emp_eval.py")
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred while executing the script:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()


    def run_sb_emp_eval(self):
        """
        Execute the sb_emp_eval.py script when the "SB EMP EVAL EXP" button is clicked.
        """
        try:
            # Determine the path to the sb_emp_eval.py script
            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 
                "weekly_tasks", 
                "sb_emp_eval.py"
            )
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at '{script_path}'")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Exception", 
                f"An error occurred while executing the script:\n{str(e)}"
            )
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()
            

    def run_pending_admission(self):
        """
        Execute the pending_admission.py script when the "Pending Admission" button is clicked.
        """
        try:
            # Determine the path to the pending_admission.py script
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "weekly_tasks", "pending_admission.py")
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred while executing the script:\n{str(e)}")
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    def run_pending_caregiver_assignment(self):
        """
        Execute the pending_caregiver_assignment.py script when the 
        "Pending Caregiver Assignment" button is clicked.
        """
        try:
            # Determine the path to the pending_caregiver_assignment.py script
            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 
                "weekly_tasks", 
                "pending_caregiver_assignment.py"
            )
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Exception", 
                f"An error occurred while executing the script:\n{str(e)}"
            )
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()
    
    def run_pending_IHCC_admission(self):
        """
        Execute the pending_IHCC_admission.py script when the "Pending IHCC Admission" button is clicked.
        """
        try:
            # Determine the path to the pending_IHCC_admission.py script
            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 
                "weekly_tasks", 
                "pending_IHCC_admission.py"
            )
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at '{script_path}'")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Exception", 
                f"An error occurred while executing the script:\n{str(e)}"
            )
            self.set_buttons_enabled(True)
            self.animation_label.hide()
            self.cancel_button.hide()
            self.log_text.hide()
            self.log_label.hide()

    def run_pending_PERS_installation(self):
        """
        Execute the pending_PERS_Installation.py script when the "Pending PERS Installation" button is clicked.
        """
        try:
            # Determine the path to the pending_PERS_Installation.py script
            script_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 
                "weekly_tasks", 
                "pending_PERS_Installation.py"
            )
            
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at '{script_path}'")
                return
            
            # Disable the sub-category buttons to prevent multiple executions
            self.set_buttons_enabled(False)
            
            # Show the animation and cancel button
            self.animation_label.show()
            self.animation_movie.start()
            self.cancel_button.show()
            self.log_text.show()
            self.log_label.show()
            
            # Clear previous logs
            self.log_text.clear()
            
            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            # Use the '-u' flag to force unbuffered output
            self.process.setArguments(['-u', script_path])
            
            # Connect signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)
            
            # Start the process
            self.process.start()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Exception", 
                f"An error occurred while executing the script:\n{str(e)}"
            )
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
        """
        Handle the completion of the process.
        
        Args:
            exitCode (int): The exit code of the process.
            exitStatus (QProcess.ExitStatus): The exit status of the process.
        """
        # Hide the animation and cancel button
        self.animation_movie.stop()
        self.animation_label.hide()
        self.cancel_button.hide()
        
        # Re-enable the sub-category buttons
        self.set_buttons_enabled(True)
        
        # Check if process exited successfully
        if exitCode == 0:
            # Log success in the log area
            self.log_text.append("<span style='color: green;'>Script executed successfully.</span>")
        else:
            # Log failure in the log area
            self.log_text.append(f"<span style='color: red;'>Script failed with exit code {exitCode}.</span>")
    
    def cancel_process(self):
        """
        Cancel the running process.
        """
        if self.process and self.process.state() == QProcess.Running:
            self.process.kill()
            self.process = None
            # Hide the animation and cancel button
            self.animation_movie.stop()
            self.animation_label.hide()
            self.cancel_button.hide()
            # Re-enable the sub-category buttons
            self.set_buttons_enabled(True)
            # Log cancellation
            self.log_text.append("<span style='color: orange;'>Script execution has been cancelled.</span>")
    
    def set_buttons_enabled(self, enabled):
        """
        Enable or disable all sub-category buttons.
        
        Args:
            enabled (bool): True to enable, False to disable.
        """
        for i in range(self.button_layout.count()):
            widget = self.button_layout.itemAt(i).widget()
            if isinstance(widget, QPushButton):
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


# Main entry point
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec())
