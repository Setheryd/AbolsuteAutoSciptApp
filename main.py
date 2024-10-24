# main.py

# =============================================================================
# Imports
# =============================================================================

# Standard Library Imports
import os
import sys
import threading
from datetime import datetime, timedelta
from io import StringIO

# Third-Party Imports
import win32com.client as win32  # type: ignore
import pandas as pd
import matplotlib.pyplot as plt
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
    QComboBox,
    QSpinBox,
    QTableWidgetItem,
    QSplitter,
    QAbstractItemView,
    QProgressBar,
    QGraphicsBlurEffect,
)
from PySide6.QtCore import Qt, Slot, QTimer, QPropertyAnimation, QObject, Signal, QThread  # type: ignore
from PySide6.QtGui import QMovie, QPixmap, QPainter, QColor, QGuiApplication  # type: ignore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from ui.ability_ui import setup_ability_mode_tabs  # Import the AbilityTab class

# Importing all the necessary scripts as modules
from daily_tasks.birthday import main as birthday_main

from weekly_tasks.in_emp_id_exp import main as in_emp_id_exp_main
from weekly_tasks.in_emp_inservices_exp import main as in_emp_inservices_exp_main
from weekly_tasks.in_pat_sup_exp import main as in_pat_sup_exp_main
from weekly_tasks.indy_emp_eval import main as indy_emp_eval_main
from weekly_tasks.pending_admission import main as pending_admission_main
from weekly_tasks.pending_caregiver_assignment import (
    main as pending_caregiver_assignment_main,
)
from weekly_tasks.pending_IHCC_admission import main as pending_IHCC_admission_main
from weekly_tasks.pending_PERS_installation import (
    main as pending_PERS_installation_main,
)
from weekly_tasks.sb_emp_eval import main as sb_emp_eval_main
from weekly_tasks.sb_emp_id_exp import main as sb_emp_id_exp_main
from weekly_tasks.sb_emp_inservices_exp import main as sb_emp_inservices_exp_main
from weekly_tasks.sb_pat_sup_exp import main as sb_pat_sup_exp_main

from monthly_tasks.age import main as age_main
from monthly_tasks.employee_attrition_email import main as employee_attrition_email_main
from monthly_tasks.employee_attrition import main as employee_attrition_main
from monthly_tasks.next_month_NOA_exp import main as next_month_NOA_exp_main
from monthly_tasks.NOA_exp import main as NOA_exp_main
from monthly_tasks.patient_attrition_email import main as patient_attrition_email_main
from monthly_tasks.patient_attrition import main as patient_attrition_main

# =============================================================================
# Constants
# =============================================================================

# Categorized Items
daily_items = sorted(["Employee Birthday Email"])
daily_items.insert(0, "Run All")  # Ensure "Run All" is at the top

weekly_items = sorted(
    [
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
        "SB PAT SUP EXP",
    ]
)
weekly_items.insert(0, "Run All")  # Ensure "Run All" is at the top

monthly_items = sorted(
    [
        "Age Notification",
        "Employee Attrition",
        "Expired NOAs",
        "Inventory Request",
        "Next Months Expired NOAs",
        "Patient Attrition",
    ]
)
monthly_items.insert(0, "Run All")  # Ensure "Run All" is at the top

# =============================================================================
# Helper Classes
# =============================================================================


class ScriptButtonWidget(QWidget):
    """
    Composite widget containing a script button and an optional status indicator.
    """

    def __init__(self, script_name, button, has_status=True):
        super().__init__()
        self.script_name = script_name
        self.button = button
        self.has_status = has_status

        # Status Label
        self.status_label = QLabel()
        if self.has_status:
            self.status_label.setFixedSize(
                24, 24
            )  # Increased size for better visibility
            self.status_label.setAlignment(Qt.AlignCenter)
            self.status = "pending"
            self.update_status_icon()
        else:
            self.status_label.hide()

        # Layout Configuration
        layout = QHBoxLayout()
        layout.addWidget(self.button)
        if self.has_status:
            layout.addWidget(self.status_label)
        layout.setAlignment(Qt.AlignLeft)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.setLayout(layout)

    def get_resource_path(self, relative_path):
        """Get the absolute path to the resource, works for PyInstaller executable."""
        try:
            # PyInstaller creates a temporary folder and stores the path in _MEIPASS
            base_path = sys._MEIPASS
        except AttributeError:
            # If not running as an executable, use the current script directory
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

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
        painter.fillRect(white_pixmap.rect(), QColor("white"))
        painter.end()
        return white_pixmap

    def update_status_icon(self):
        """
        Updates the status icon based on the current status.
        """
        if not self.has_status:
            return

        # Construct absolute path to the icon

        icon_mapping = {
            "pending": "pending.png",
            "running": "running.gif",  # Assuming running is a gif for animation
            "success": "success.png",
            "failed": "failed.png",
        }
        icon_file = icon_mapping.get(self.status, "")
        icon_path = self.get_resource_path(os.path.join("resources", icon_file))

        if os.path.exists(icon_path):
            pixmap = QPixmap(icon_path)
            if pixmap.isNull():
                print(f"Failed to load pixmap from {icon_path}")
                self.status_label.setText(self.status.capitalize())
            else:
                if self.status == "running":
                    # Use a QMovie for running animation
                    self.movie = QMovie(icon_path)
                    self.status_label.setMovie(self.movie)
                    self.movie.start()
                else:
                    # Tint 'pending.png' to white if necessary
                    if self.status in ["pending"]:
                        pixmap = self.make_pixmap_white(pixmap)

                    self.status_label.setPixmap(
                        pixmap.scaled(
                            self.status_label.size(),
                            Qt.KeepAspectRatio,
                            Qt.SmoothTransformation,
                        )
                    )
        else:
            print(f"Icon path does not exist: {icon_path}")
            self.status_label.setText(self.status.capitalize())

    def set_status(self, status):
        """
        Sets the status and updates the status icon.

        Args:
            status (str): New status ('pending', 'running', 'success', 'failed').
        """
        self.status = status
        self.update_status_icon()


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=10, height=6, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.ax = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)

        # Fix the size based on figsize and dpi
        self.setFixedSize(width * dpi, height * dpi)

        # Prevent the canvas from expanding
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)


class CustomTableWidget(QTableWidget):
    """
    Custom QTableWidget to handle Ctrl+C for copying selected data.
    """

    def keyPressEvent(self, event):
        """Handle key press events, including Ctrl+C."""
        if event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:
            self.copy_selected_data()
        else:
            super().keyPressEvent(event)

    def copy_selected_data(self):
        """
        Copy selected cells from QTableWidget to clipboard.
        """
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "No cells are selected to copy.")
            return

        clipboard_content = ""
        for selected_range in selected_ranges:
            rows = range(selected_range.topRow(), selected_range.bottomRow() + 1)
            cols = range(selected_range.leftColumn(), selected_range.rightColumn() + 1)

            for row in rows:
                row_content = []
                for col in cols:
                    item = self.item(row, col)
                    row_content.append(item.text() if item else "")
                clipboard_content += "\t".join(row_content) + "\n"

        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_content.strip())

        QMessageBox.information(
            self, "Copied", "Selected data has been copied to the clipboard."
        )


# =============================================================================
# Worker Class for Script Execution
# =============================================================================


class Worker(QObject):
    finished = Signal(str, str)  # script_name, result
    log_output = Signal(str)
    log_error = Signal(str)

    @Slot(object, str)
    def run(self, script_func, script_name):
        """
        Executes the given script function and emits signals based on execution status.

        Args:
            script_func (callable): The script's main function to execute.
            script_name (str): The name of the script.
        """
        try:
            # Redirect stdout and stderr
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            sys.stdout = StringIO()
            sys.stderr = StringIO()

            # Execute the script function
            script_func()

            # Capture outputs
            stdout_output = sys.stdout.getvalue()
            stderr_output = sys.stderr.getvalue()

            # Restore stdout and stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr

            # Emit logs
            if stdout_output:
                self.log_output.emit(stdout_output)
            if stderr_output:
                self.log_error.emit(stderr_output)
                self.finished.emit(script_name, "failed")
            else:
                self.finished.emit(script_name, "success")

        except Exception as e:
            # Restore stdout and stderr in case of exception
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.log_error.emit(str(e))
            self.finished.emit(script_name, "failed")


# =============================================================================
# Main Application Class
# =============================================================================


class MainApp(QWidget):
    """
    Main application window for the Absolute Caregivers Auto Scripting App.
    """

    def __init__(self):
        super().__init__()

        # Initialize Attributes
        self.current_script_name = None
        self.is_ability_mode = False
        self.threads = []  # To keep references to QThreads
        self.pending_scripts = []  # Initialize pending_scripts

        # Timer for Timeout
        self.timer = QTimer(self)
        self.timer.setInterval(60000)  # 60 seconds
        self.timer.timeout.connect(self.handle_timeout)

        # Set up Window Properties
        self.setup_window()

        # Initialize Scripts Mapping
        self.initialize_scripts_mapping()

        # Initialize UI Components
        self.initialize_ui()

    def setup_window(self):
        """
        Configures the main window properties.
        """
        self.setWindowTitle("Absolute Caregivers Auto Scripting App")
        screen = QGuiApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setGeometry(0, 0, screen_geometry.width(), screen_geometry.height() - 20)

    def get_resource_path(self, relative_path):
        """Get the absolute path to the resource, works for PyInstaller executable."""
        try:
            # PyInstaller creates a temporary folder and stores the path in _MEIPASS
            base_path = sys._MEIPASS
        except AttributeError:
            # If not running as an executable, use the current script directory
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

    def initialize_scripts_mapping(self):
        """
        Maps script names to their corresponding functions.
        """
        self.scripts_mapping = {
            # Daily Scripts
            "Employee Birthday Email": birthday_main,
            # Weekly Scripts
            "Caregiver ID Exp": in_emp_id_exp_main,
            "IN Emp EVAL EXP": indy_emp_eval_main,
            "IN Emp In-Services EXP": in_emp_inservices_exp_main,
            "IN PAT SUP EXP": in_pat_sup_exp_main,
            "Pending Admission": pending_admission_main,
            "Pending Caregiver Assignment": pending_caregiver_assignment_main,
            "Pending IHCC Admission": pending_IHCC_admission_main,
            "Pending PERS Installation": pending_PERS_installation_main,
            "SB EMP EVAL EXP": sb_emp_eval_main,
            "SB Emp Inservices Exp": sb_emp_inservices_exp_main,
            "SB ID EXP": sb_emp_id_exp_main,
            "SB PAT SUP EXP": sb_pat_sup_exp_main,
            # Monthly Scripts
            "Age Notification": age_main,
            "Employee Attrition": employee_attrition_main,
            "Expired NOAs": NOA_exp_main,
            "Inventory Request": None,  # Placeholder if not implemented
            "Next Months Expired NOAs": next_month_NOA_exp_main,
            "Patient Attrition": patient_attrition_main,
            "Employee Attrition Email": employee_attrition_email_main,
            "Patient Attrition Email": patient_attrition_email_main,
        }

        # Add "Run All" functionality to execute all scripts in each category
        self.run_all_functions = {
            "Daily": self.run_all_daily_items,
            "Weekly": self.run_all_weekly_items,
            "Monthly": self.run_all_monthly_items,
        }

    def initialize_ui(self):
        """
        Initializes and sets up all UI components.
        """
        # Create a layout for the entire window

        tab_style = """
            QTabWidget::pane {
                border: none;
                background-color: #2e2e2e;
            }

            QTabBar::tab {
                background: #3c3c3c;
                color: white;
                padding: 10px 20px;
                border: 1px solid #444;
                border-bottom: none;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                margin-right: 2px;
                font-family: "Segoe UI", sans-serif;
                font-size: 14px;
            }

            QTabBar::tab:hover {
                background: #505050;
            }

            QTabBar::tab:selected {
                background: #207544;
                font-weight: bold;
                color: #d9e6f2;
                color: white;
                margin-top: 0px;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }
        """

        self.main_layout = QVBoxLayout(self)

        # Create a toggle button for switching between Absolute and Ability
        self.toggle_button = QPushButton("Absolute", self)
        self.toggle_button.setCheckable(True)
        self.toggle_button.setStyleSheet(
            """
            QPushButton {
                background-color: #207544;
                border-radius: 15px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
                color: white;
            }
            QPushButton:checked {
                background-color: #7e07a0;
            }
        """
        )
        self.toggle_button.clicked.connect(self.switch_mode)

        # Add the toggle button to the top of the layout
        self.main_layout.addWidget(self.toggle_button, alignment=Qt.AlignCenter)

        # Create a QTabWidget for switching between tabs
        self.tab_widget = QTabWidget(self)

        # Apply the custom style sheet to the QTabWidget
        self.tab_widget.setStyleSheet(tab_style)

        # Setup the default Absolute side with tabs on the left
        self.setup_scripts_tab()
        self.setup_dashboard_tab()
        self.setup_graph_tab()

        # Add the tab widget to the layout
        self.main_layout.addWidget(self.tab_widget)

        self.setLayout(self.main_layout)

        # Initial state is Absolute
        self.is_ability_mode = False

    def switch_mode(self):
        """
        Switches between 'Absolute' mode and 'Ability' mode.
        """
        if self.toggle_button.isChecked():
            # Switch to Ability mode
            self.toggle_button.setText("Ability")
            self.is_ability_mode = True

            # Hide the Absolute tabs and show Ability tabs on the right side
            self.tab_widget.clear()  # Clear the existing tabs
            setup_ability_mode_tabs(self)  # Setup Ability tabs

        else:
            # Switch back to Absolute mode
            self.toggle_button.setText("Absolute")
            self.is_ability_mode = False

            # Hide the Ability tabs and show Absolute tabs on the left side
            self.tab_widget.clear()  # Clear the existing tabs
            self.tab_widget.setStyleSheet(
                """QTabWidget::pane {
                border: none;
                background-color: #2e2e2e;
            }

            QTabBar::tab {
                background: #3c3c3c;
                color: white;
                padding: 10px 20px;
                border: 1px solid #444;
                border-bottom: none;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                margin-right: 2px;
                font-family: "Segoe UI", sans-serif;
                font-size: 14px;
            }

            QTabBar::tab:hover {
                background: #505050;
            }

            QTabBar::tab:selected {
                background: #207544;
                font-weight: bold;
                color: #d9e6f2;
                color: white;
                margin-top: 0px;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            
            QPushButton:checked {
                background-color: #207544;
                color: white;
            }
            """
            )
            self.setup_scripts_tab()  # Setup Absolute tabs
            self.setup_dashboard_tab()
            self.setup_graph_tab()

    # ---------------------------------
    # Scripts Tab Setup
    # ---------------------------------

    def setup_scripts_tab(self):
        """Set up the layout and components for the Scripts tab."""
        self.scripts_tab = QWidget()
        scripts_layout = QHBoxLayout(self.scripts_tab)
        scripts_layout.setContentsMargins(20, 20, 20, 20)
        scripts_layout.setSpacing(20)

        # Category Buttons Layout
        self.category_layout = QVBoxLayout()
        self.category_layout.setSpacing(10)
        self.category_layout.setAlignment(Qt.AlignTop)

        # "Categories" Header
        categories_header = QLabel("Categories")
        categories_header.setStyleSheet(
            """
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #207544;
            }
        """
        )
        self.category_layout.addWidget(categories_header, alignment=Qt.AlignCenter)

        # Create Main Category Buttons as Toggles
        self.daily_button = self.create_category_button("Daily", daily_items, "daily")
        self.weekly_button = self.create_category_button(
            "Weekly", weekly_items, "weekly"
        )
        self.monthly_button = self.create_category_button(
            "Monthly", monthly_items, "monthly"
        )

        self.category_layout.addWidget(self.daily_button)
        self.category_layout.addWidget(self.weekly_button)
        self.category_layout.addWidget(self.monthly_button)

        # Add Stretch to Push Buttons to the Top
        self.category_layout.addStretch()  # Corrected: Add stretch to the layout, not the button

        # Container for Category Buttons
        self.category_container = QWidget()
        self.category_container.setLayout(self.category_layout)
        self.category_container.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Preferred
        )
        scripts_layout.addWidget(self.category_container)

        # Scripts Buttons Layout
        self.script_layout = QVBoxLayout()
        self.script_layout.setSpacing(10)
        self.script_layout.setAlignment(Qt.AlignTop)

        # "Scripts" Header
        scripts_header = QLabel("Scripts")
        scripts_header.setStyleSheet(
            """
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #207544; 
                margin-top: 6px;
            }
        """
        )
        self.script_layout.addWidget(scripts_header, alignment=Qt.AlignCenter)

        # Container for Sub-Category Buttons
        self.button_container = QWidget()
        self.button_layout = QVBoxLayout()
        self.button_layout.setContentsMargins(0, 0, 0, 0)

        self.button_layout.setSpacing(10)
        self.button_layout.setAlignment(Qt.AlignTop)

        self.button_container.setLayout(self.button_layout)
        self.button_container.setStyleSheet("background-color: grey;")

        # Scroll Area for Sub-Category Buttons
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.scroll_area.setWidget(self.button_container)
        self.scroll_area.setStyleSheet("background-color: transparent;")
        self.scroll_area.setMinimumWidth(350)

        # Transparent background for scroll area

        self.script_layout.addWidget(self.scroll_area, alignment=Qt.AlignHCenter)

        # Add Scripts Layout to Main Scripts Layout
        scripts_layout.addLayout(self.script_layout)

        # Indicators Layout (Rightmost Column)
        self.indicator_layout = QVBoxLayout()
        self.indicator_layout.setAlignment(Qt.AlignTop)
        self.indicator_layout.setSpacing(10)

        # Animation Label
        self.animation_label = QLabel(self)
        self.animation_label.setFixedSize(100, 100)

        # Use the get_resource_path method to handle paths correctly
        loading_gif_path = self.get_resource_path(
            os.path.join("resources", "loading.gif")
        )

        if not os.path.exists(loading_gif_path):
            print(
                f"Loading GIF not found at {loading_gif_path}. Please ensure the file exists."
            )

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
        self.log_text.setStyleSheet(
            """
            QTextEdit {
                background-color: #2e2e2e;
                color: white;
                font-family: Consolas;
                font-size: 12pt;
            }
        """
        )

        # Add Components to Indicators Layout
        self.indicator_layout.addWidget(self.animation_label, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.cancel_button, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.log_label, alignment=Qt.AlignLeft)
        self.indicator_layout.addWidget(self.log_text, alignment=Qt.AlignCenter)

        # Spacer to Push Indicators to the Top
        self.indicator_layout.addSpacerItem(
            QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        )

        # Container for Indicators
        self.indicator_container = QWidget()
        self.indicator_container.setLayout(self.indicator_layout)
        self.indicator_container.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Preferred
        )
        scripts_layout.addWidget(self.indicator_container)

        # State to Track Expanded Categories
        self.expanded_categories = {"Daily": False, "Weekly": False, "Monthly": False}

        # Initialize Category Styles
        self.update_category_styles()

        # Add Scripts Tab to QTabWidget
        self.tab_widget.addTab(self.scripts_tab, "Scripts")

        # Initialize Script Buttons Dictionary
        self.script_buttons = {}

    def create_category_button(self, text, items, identifier):
        """
        Creates a category button with specified properties.

        Args:
            text (str): Button text.
            items (list): List of sub-category items.
            identifier (str): Unique identifier for the button.

        Returns:
            QPushButton: Configured QPushButton instance.
        """
        button = QPushButton(text, self)
        button.setFixedWidth(300)  # Set a fixed width for all category buttons
        button.setCheckable(True)  # Make the button checkable
        button.clicked.connect(lambda: self.toggle_category(text, items))
        button.setObjectName(identifier)  # Assign a unique object name
        button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        return button

    def toggle_category(self, category, items):
        """
        Toggles the expansion state of a category.

        Args:
            category (str): The category to toggle.
            items (list): List of sub-category items.
        """
        if self.expanded_categories[category]:
            # Collapse Category
            self.clear_buttons()
            self.expanded_categories[category] = False
            self.set_category_button_checked(category, False)
            self.current_category = None
        else:
            # Expand Category
            self.clear_buttons()
            self.current_category = category
            self.show_buttons(items)
            self.expanded_categories[category] = True

            # Collapse Other Categories
            for cat in self.expanded_categories:
                if cat != category:
                    self.expanded_categories[cat] = False
                    self.set_category_button_checked(cat, False)

            # Update Styles
            self.update_category_styles()

    def clear_buttons(self):
        """
        Clears all sub-category buttons from the layout.
        """
        for i in reversed(range(self.button_layout.count())):
            widget = self.button_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        self.selected_subcategory_button = None
        self.script_buttons.clear()

    def show_buttons(self, items):
        """
        Creates and displays sub-category buttons based on the selected category.

        Args:
            items (list): The list of sub-category items to create buttons for.
        """

        for item in items:
            button = QPushButton(item, self)
            button.setFixedWidth(300)  # Fixed width for consistency
            button.setCheckable(True)

            if item == "Run All":
                # Configure "Run All" Button
                button.setObjectName("run_all")
                button.clicked.connect(
                    lambda checked, cat=self.current_category: self.run_all_functions.get(
                        cat, self.show_message
                    )()
                )
                script_widget = ScriptButtonWidget(item, button, has_status=False)
            else:
                # Configure Individual Script Button
                if item in self.scripts_mapping and self.scripts_mapping[item]:
                    script_func = self.scripts_mapping[item]
                    button.clicked.connect(
                        lambda checked, func=script_func, name=item: self.run_single_script(
                            func, name
                        )
                    )
                    script_widget = ScriptButtonWidget(item, button, has_status=True)
                    self.script_buttons[item] = script_widget
                else:
                    button.clicked.connect(self.show_message)
                    script_widget = ScriptButtonWidget(item, button, has_status=False)

            # Apply Stylesheet
            button.setStyleSheet(
                """
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
                    background-color: #207544;
                    border: 1px solid #89a4ff;
                    color: white;
                }
            """
            )
            button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            # Connect Selection Handler
            button.clicked.connect(
                lambda checked, btn=button: self.select_subcategory(btn)
            )

            # Add to Layout
            self.button_layout.addWidget(script_widget, alignment=Qt.AlignLeft)

    def select_subcategory(self, button):
        """
        Handles the selection of a sub-category button, ensuring only one is checked.

        Args:
            button (QPushButton): The button that was clicked.
        """
        if (
            hasattr(self, "selected_subcategory_button")
            and self.selected_subcategory_button
            and self.selected_subcategory_button != button
        ):
            self.selected_subcategory_button.setChecked(False)

        if button.isChecked():
            self.selected_subcategory_button = button
        else:
            self.selected_subcategory_button = None

    def show_message(self):
        """
        Displays a message box with the button text.
        """
        button = self.sender()
        if button:
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Button Clicked")
            msg_box.setText(f"You clicked: {button.text()}")
            msg_box.exec()

    # ---------------------------------
    # Dashboard Tab Setup
    # ---------------------------------

    def setup_dashboard_tab(self):
        """
        Sets up the layout and components for the Dashboard tab.
        """
        self.dashboard_tab_widget = QWidget()
        self.tab_widget.addTab(self.dashboard_tab_widget, "Dashboard")

        self.secondary_layout = QVBoxLayout(self.dashboard_tab_widget)

        # QSplitter for Resizable Columns
        splitter = QSplitter(Qt.Horizontal)

        # Left-Side Layout for Custom Program Buttons
        self.setup_dashboard_left_side(splitter)

        # Right-Side Layout for DataFrame Display
        self.setup_dashboard_right_side(splitter)

        # Configure Splitter Sizes
        splitter.setSizes([300, 700])  # Adjust as necessary

        # Add Splitter to Secondary Layout
        self.secondary_layout.addWidget(splitter)
        self.secondary_layout.setStretch(0, 1)

        # Bottom Buttons Layout
        self.setup_dashboard_bottom_buttons()

    def setup_dashboard_left_side(self, splitter):
        """
        Sets up the left-side layout for the Dashboard tab.

        Args:
            splitter (QSplitter): The splitter to add the left-side widget to.
        """

        if self.is_ability_mode:
            custom_programs = {"Example Ability Script": "example_script.py"}
        else:
            custom_programs = {
                "Active Patients per Month": "Admission_by_Month.py",
                "View Patient Data": "patient_data_extractor.py",
                "Active Contractor per Month": "Caregiver_by_Month.py",
                "View Contractor Data": "caregiver_data_extractor.py",
                "View Employee Records Data": "employee_records_data_extractor.py",
                "Employee Records Hours": "Employee_Records_by_Month.py",
                "Active Admission Count by Service": "Active_Admission_by_Service.py",
                "View Billing Data": "billing_files_extractor.py",
                "Patient Tenure": "Patient_Tenure_by_Group.py",
                "Patient Attrition": "patient_attrition_email.py",
                "Employee Attrition Email": "employee_attrition_email.py",
                "Paysource Patient Count": "Paysource_Patient_count.py",
            }

        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)

        for button_label, program in custom_programs.items():
            button = QPushButton(button_label)
            button.clicked.connect(
                lambda checked, prog=program: self.run_python_script(prog)
            )
            button_layout.addWidget(button)

        button_container.setLayout(button_layout)
        button_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        button_container.adjustSize()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(button_container)

        left_widget = QWidget()
        left_widget.setLayout(QVBoxLayout())
        left_widget.layout().addWidget(scroll_area)

        splitter.addWidget(left_widget)

    def setup_dashboard_right_side(self, splitter):
        """
        Sets up the right-side layout for displaying the DataFrame and includes a Gaussian blur
        and loading bar while retrieving data.

        Args:
            splitter (QSplitter): The splitter to add the right-side widget to.
        """
        # Set up the table widget where DataFrame will be displayed
        self.table_widget = CustomTableWidget()
        self.table_widget.setColumnCount(0)
        self.table_widget.setRowCount(0)
        self.table_widget.setSelectionMode(
            QAbstractItemView.ExtendedSelection
        )  # Allow multiple selections
        self.table_widget.setSelectionBehavior(
            QAbstractItemView.SelectItems
        )  # Item selection behavior

        # Apply styles to show cell selection clearly
        self.table_widget.setStyleSheet(
            """
            QTableWidget {
                background-color: #f0f0f0;
                color: black;
            }
            QTableWidget::item:selected {
                background-color: #207544; /* Custom color for selected cells */
                color: white; /* Text color for selected cells */
               
            }
            QTableWidget::item {
                background-color: white; /* Normal cell background */
                color: black; /* Normal cell text color */
            }
        """
        )

        # Container to hold the table and blur/loading elements
        self.right_side_widget = QWidget()
        right_side_layout = QVBoxLayout(self.right_side_widget)

        # Loading bar
        self.loading_bar = QProgressBar(self.right_side_widget)
        self.loading_bar.setRange(0, 0)  # Indeterminate progress bar
        self.loading_bar.setTextVisible(False)
        self.loading_bar.setFixedHeight(30)
        self.loading_bar.setStyleSheet(
            """
            QProgressBar {
                background-color: #e0e0e0;
                border: none;
            }
            QProgressBar::chunk {
                background-color: #207544;
            }
        """
        )

        # "Retrieving Data" label
        self.retrieving_data_label = QLabel(
            "Retrieving data...", self.right_side_widget
        )
        self.retrieving_data_label.setAlignment(Qt.AlignCenter)
        self.retrieving_data_label.setStyleSheet(
            """
            QLabel {
                color: #207544;
                font-size: 16pt;
                font-weight: bold;
            }
        """
        )

        # Add table widget, loading bar, and label to layout
        right_side_layout.addWidget(self.table_widget)
        right_side_layout.addWidget(self.retrieving_data_label)
        right_side_layout.addWidget(self.loading_bar)

        # Initially hide the loading components
        self.loading_bar.hide()
        self.retrieving_data_label.hide()

        # Apply blur effect to the table when retrieving data
        self.blur_effect = QGraphicsBlurEffect(self)
        self.blur_effect.setBlurRadius(10)
        self.table_widget.setGraphicsEffect(self.blur_effect)
        self.blur_effect.setEnabled(False)

        splitter.addWidget(self.right_side_widget)

    def setup_dashboard_bottom_buttons(self):
        """
        Sets up the bottom buttons (Save DataFrame, Copy Selected) for the Dashboard tab.
        """
        bottom_buttons_layout = QHBoxLayout()
        bottom_buttons_layout.setAlignment(Qt.AlignRight)

        # Save DataFrame Button
        self.save_dataframe_button = QPushButton("Export to CSV")
        self.save_dataframe_button.clicked.connect(self.save_dataframe)

        # Add Buttons to Layout
        bottom_buttons_layout.addWidget(self.save_dataframe_button)

        # Add Bottom Buttons Layout to Secondary Layout
        self.secondary_layout.addLayout(bottom_buttons_layout)

    # ---------------------------------
    # Graph Tab Setup
    # ---------------------------------

    def setup_graph_tab(self):
        """
        Sets up the layout and components for the Graph tab.
        """
        self.graph_tab_widget = QWidget()
        self.tab_widget.addTab(self.graph_tab_widget, "Graph")

        self.graph_layout = QHBoxLayout(self.graph_tab_widget)
        self.graph_layout.setContentsMargins(20, 20, 20, 20)
        self.graph_layout.setSpacing(15)

        # Left-Side Controls Layout
        self.setup_graph_controls(self.graph_layout)

        # Right-Side Graph Area Layout
        self.setup_graph_area(self.graph_layout)

        # Apply Graph Styles
        self.apply_graph_styles()

    def setup_graph_controls(self, graph_layout):
        """
        Sets up the left-side controls for the Graph tab.

        Args:
            graph_layout (QHBoxLayout): The main graph layout to add controls to.
        """
        controls_layout = QVBoxLayout()
        controls_layout.setContentsMargins(20, 20, 20, 20)
        controls_layout.setSpacing(15)

        # Helper function to set label styles
        def set_label_style(label):
            color = "#7e07a0" if self.is_ability_mode else "#207544"
            label.setStyleSheet(
                f"""
                QLabel {{
                    font-size: 14pt;
                    font-weight: bold;
                    color: {color};
                }}
                """
            )

        # X-axis Selection
        self.x_axis_label = QLabel("Select X-axis:")
        set_label_style(self.x_axis_label)

        self.x_axis_combo = QComboBox(self)
        self.x_axis_combo.setStyleSheet(
            """
            QComboBox {
                padding: 5px;
                font-size: 12pt;
            }
        """
        )
        controls_layout.addWidget(self.x_axis_label)
        controls_layout.addWidget(self.x_axis_combo)

        # Y-axis Selection
        self.y_axis_label = QLabel("Select Y-axis:")
        set_label_style(self.y_axis_label)

        self.y_axis_combo = QComboBox(self)
        self.y_axis_combo.setStyleSheet(
            """
            QComboBox {
                padding: 5px;
                font-size: 12pt;
            }
        """
        )
        controls_layout.addWidget(self.y_axis_label)
        controls_layout.addWidget(self.y_axis_combo)

        # Label SpinBox
        self.label_spinbox_label = QLabel("Number of X-axis labels to display:")
        set_label_style(self.label_spinbox_label)

        self.label_spinbox = QSpinBox(self)
        self.label_spinbox.setRange(1, 20)
        self.label_spinbox.setValue(6)
        self.label_spinbox.setStyleSheet(
            """
            QSpinBox {
                padding: 5px;
                font-size: 12pt;
            }
        """
        )
        controls_layout.addWidget(self.label_spinbox_label)
        controls_layout.addWidget(self.label_spinbox)

        # Spacer to Push Plot Button to Bottom
        controls_layout.addStretch()

        # Plot Graph Button
        self.plot_button = QPushButton("Plot Graph", self)
        self.plot_button.setFixedWidth(150)

        button_color = "#7e07a0" if self.is_ability_mode else "#207544"
        button_hover = "#9b32bf" if self.is_ability_mode else "#28a05e"
        self.plot_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {button_color};
                border: none;
                border-radius: 8px;
                padding: 10px;
                font-size: 14pt;
                color: white;
            }}
            QPushButton:hover {{
                background-color: {button_hover};
            }}
        """
        )
        self.plot_button.clicked.connect(self.plot_graph)
        controls_layout.addWidget(self.plot_button, alignment=Qt.AlignCenter)

        # Add Controls Layout to Main Graph Layout
        graph_layout.addLayout(controls_layout, stretch=1)

    def setup_graph_area(self, graph_layout):
        """
        Sets up the right-side graph area for the Graph tab.

        Args:
            graph_layout (QHBoxLayout): The main graph layout to add the graph area to.
        """
        graph_area_layout = QVBoxLayout()
        graph_area_layout.setContentsMargins(20, 20, 20, 20)
        graph_area_layout.setSpacing(10)

        # Matplotlib Canvas with fixed size
        self.canvas = MplCanvas(
            self, width=10, height=6, dpi=100
        )  # Reduced width and height
        self.canvas.setStyleSheet(
            """
            background-color: #2e2e2e;
        """
        )
        graph_area_layout.addWidget(self.canvas)

        # Add Graph Area Layout to Main Graph Layout
        graph_layout.addLayout(graph_area_layout, stretch=3)

    def apply_graph_styles(self):
        """
        Applies consistent styles to the graph area.
        """
        # Set the figure and axes background
        self.canvas.figure.patch.set_facecolor(
            "#2e2e2e"
        )  # Dark background for the figure
        self.canvas.ax.set_facecolor("#2e2e2e")  # Dark background for the plot area

        # Set tick parameters
        self.canvas.ax.tick_params(axis="x", colors="white")  # White ticks for X-axis
        self.canvas.ax.tick_params(axis="y", colors="white")  # White ticks for Y-axis

        # Set axis labels color
        self.canvas.ax.xaxis.label.set_color("white")  # White X-axis label
        self.canvas.ax.yaxis.label.set_color("white")  # White Y-axis label

        # Set spines (axis lines) color to white
        for spine in self.canvas.ax.spines.values():
            spine.set_edgecolor("white")

        # Set tick label colors and sizes
        self.canvas.ax.tick_params(
            axis="both", which="major", labelsize=10, colors="white"
        )

        # Refresh the canvas
        self.canvas.draw()

    def plot_graph(self):
        """
        Plots the graph based on the selected parameters.
        """
        if not hasattr(self, "df") or self.df.empty:
            QMessageBox.warning(
                self,
                "No Data",
                "No data available to plot. Please run a script to load data.",
            )
            return

        # Clear Previous Plot
        self.canvas.ax.clear()

        # Get Selected Columns
        x_column = self.x_axis_combo.currentText()
        y_column = self.y_axis_combo.currentText()

        if not x_column or not y_column:
            QMessageBox.warning(
                self, "Selection Error", "Please select both X and Y axes."
            )
            return

        # Get Number of Labels to Display
        max_labels = self.label_spinbox.value()

        try:
            x_data = self.df[x_column]
            y_data = self.df[y_column]

            # Scatter Plot
            scatter_plot = self.canvas.ax.scatter(
                x_data, y_data, picker=True, color="#207544"
            )

            # Set Axis Labels with Explicit White Color
            self.canvas.ax.set_xlabel(x_column, color="white", fontsize=14)
            self.canvas.ax.set_ylabel(y_column, color="white", fontsize=14)

            # Configure X-axis Ticks and Labels
            num_points = len(x_data)
            step = max(1, num_points // max_labels)
            tick_indices = range(0, num_points, step)
            self.canvas.ax.set_xticks([x_data[i] for i in tick_indices])
            self.canvas.ax.set_xticklabels(
                [x_data[i] for i in tick_indices], ha="right", color="white"
            )

            # Configure Y-axis Tick Colors
            self.canvas.ax.tick_params(axis="y", colors="white")

            # Remove tight_layout to maintain fixed size
            self.canvas.figure.tight_layout()

            # Redraw Canvas
            self.canvas.draw()

            # Connect Click Event
            self.canvas.mpl_connect("pick_event", self.on_click)

        except Exception as e:
            QMessageBox.critical(
                self,
                "Plot Error",
                f"An error occurred while plotting the graph:\n{str(e)}",
            )

    # ---------------------------------
    # DataFrame Handling
    # ---------------------------------

    def display_dataframe(self, df):
        """
        Displays a Pandas DataFrame in the QTableWidget on the Dashboard tab.

        Args:
            df (pd.DataFrame): DataFrame to display.
        """
        if df.empty:
            QMessageBox.warning(self, "No Data", "The DataFrame is empty.")
            return

        self.df = df  # Store DataFrame for future use
        self.table_widget.clear()

        # Configure Table Dimensions
        self.table_widget.setRowCount(df.shape[0])
        self.table_widget.setColumnCount(df.shape[1])
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        # Populate Table with DataFrame Data
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                self.table_widget.setItem(i, j, item)

        # Resize Columns to Fit Contents
        self.table_widget.resizeColumnsToContents()

        # Populate X and Y Axis ComboBoxes
        self.x_axis_combo.clear()
        self.y_axis_combo.clear()

        for col in df.columns:
            self.x_axis_combo.addItem(col)
            self.y_axis_combo.addItem(col)

    def save_dataframe(self):
        """
        Opens a file dialog to save the displayed DataFrame as a CSV file.
        """
        if not hasattr(self, "df"):
            QMessageBox.warning(self, "Error", "No DataFrame to save.")
            return

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save DataFrame",
            "",
            "CSV Files (*.csv);;All Files (*)",
            options=options,
        )

        if file_name:
            try:
                self.df.to_csv(file_name, index=False)
                QMessageBox.information(
                    self, "Success", f"DataFrame saved successfully to {file_name}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to save DataFrame: {str(e)}"
                )

    # ---------------------------------
    # Script Execution Methods
    # ---------------------------------

    def run_all_daily_items(self):
        """
        Executes all daily scripts sequentially.
        """
        try:
            script_items = daily_items[1:]  # Exclude "Run All"
            self.pending_scripts = [
                (item, self.scripts_mapping[item])
                for item in script_items
                if item in self.scripts_mapping and self.scripts_mapping[item]
            ]

            self.script_results = {item: "pending" for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                self.script_buttons[item].set_status("pending")

            self.set_buttons_enabled(False)

            self.show_execution_indicators()
            self.log_text.clear()

            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.reset_execution_state()

    def run_all_weekly_items(self):
        """
        Executes all weekly scripts sequentially.
        """
        try:
            script_items = weekly_items[1:]  # Exclude "Run All"
            self.pending_scripts = [
                (item, self.scripts_mapping[item])
                for item in script_items
                if item in self.scripts_mapping and self.scripts_mapping[item]
            ]

            self.script_results = {item: "pending" for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                self.script_buttons[item].set_status("pending")

            self.set_buttons_enabled(False)

            self.show_execution_indicators()
            self.log_text.clear()

            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.reset_execution_state()

    def run_all_monthly_items(self):
        """
        Executes all monthly scripts sequentially.
        """
        try:
            script_items = monthly_items[1:]  # Exclude "Run All"
            self.pending_scripts = [
                (item, self.scripts_mapping[item])
                for item in script_items
                if item in self.scripts_mapping and self.scripts_mapping[item]
            ]

            self.script_results = {item: "pending" for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                if item in self.scripts_mapping and self.scripts_mapping[item]:
                    self.script_buttons[item].set_status("pending")

            self.set_buttons_enabled(False)

            self.show_execution_indicators()
            self.log_text.clear()

            self.run_next_script()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.reset_execution_state()

    def run_next_script(self):
        """
        Runs the next script in the pending_scripts list.
        """
        if self.pending_scripts:
            next_script_name, next_script_func = self.pending_scripts.pop(0)
            self.current_script_name = next_script_name
            self.run_script(next_script_func, next_script_name)
        else:
            self.log_text.append(
                "<span style='color: green;'>All scripts executed.</span>"
            )
            self.show_summary()
            self.reset_execution_state()

    def run_script(self, script_func, script_name):
        """
        Executes a single script function using QThread.

        Args:
            script_func (callable): The script's main function to execute.
            script_name (str): The name of the script.
        """
        if not callable(script_func):
            QMessageBox.critical(
                self,
                "Error",
                f"Script function for '{script_name}' is not callable.",
            )
            self.run_next_script()
            return

        # Create a QThread object
        thread = QThread()
        # Create a worker object
        worker = Worker()
        # Move worker to the thread
        worker.moveToThread(thread)
        # Connect signals and slots
        thread.started.connect(lambda: worker.run(script_func, script_name))
        worker.finished.connect(self.on_script_finished)
        worker.log_output.connect(self.on_log_output)
        worker.log_error.connect(self.on_log_error)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        # Start the thread
        thread.start()

        # Keep a reference to the thread to prevent garbage collection
        self.threads.append(thread)

    def run_single_script(self, script_func, script_name):
        """
        Executes a single script function and handles its execution state.

        Args:
            script_func (callable): The script's main function to execute.
            script_name (str): The name of the script.
        """
        try:
            if not callable(script_func):
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Script function for '{script_name}' is not callable.",
                )
                return

            # Reset Status Icons
            if script_name in self.script_buttons:
                self.script_buttons[script_name].set_status("pending")
            self.script_results = {script_name: "pending"}

            # Disable Sub-Category Buttons
            self.set_buttons_enabled(False)

            # Show Execution Indicators
            self.show_execution_indicators()
            self.log_text.clear()

            # Update Status to 'running'
            if script_name in self.script_buttons:
                self.script_buttons[script_name].set_status("running")
                self.highlight_active_button(script_name)
            self.current_script_name = script_name

            # Create a QThread object
            thread = QThread()
            # Create a worker object
            worker = Worker()
            # Move worker to the thread
            worker.moveToThread(thread)
            # Connect signals and slots
            thread.started.connect(lambda: worker.run(script_func, script_name))
            worker.finished.connect(self.on_script_finished)
            worker.log_output.connect(self.on_log_output)
            worker.log_error.connect(self.on_log_error)
            worker.finished.connect(thread.quit)
            worker.finished.connect(worker.deleteLater)
            thread.finished.connect(thread.deleteLater)
            # Start the thread
            thread.start()

            # Keep a reference to the thread to prevent garbage collection
            self.threads.append(thread)

        except Exception as e:
            QMessageBox.critical(
                self,
                "Exception",
                f"An error occurred while executing the script:\n{str(e)}",
            )
            self.reset_execution_state()

    def execute_script(self, script_func, script_name):
        """
        Placeholder for execute_script if needed outside QThread.

        Args:
            script_func (callable): The script's main function to execute.
            script_name (str): The name of the script.
        """
        pass  # Not used in refactored version

    def on_script_finished(self, script_name, result):
        """
        Handles the completion of a script.

        Args:
            script_name (str): The name of the script.
            result (str): 'success' or 'failed'.
        """
        self.script_results[script_name] = result
        if script_name in self.script_buttons:
            self.script_buttons[script_name].set_status(result)
            color = "green" if result == "success" else "red"
            self.log_text.append(
                f"<span style='color: {color};'>Script '{script_name}' {result}.</span>"
            )
        else:
            # For "Run All", indicate overall status
            color = "green" if result == "success" else "red"
            self.log_text.append(
                f"<span style='color: {color};'>Script '{script_name}' {result}.</span>"
            )
        self.highlight_active_button(None)
        self.run_next_script()

    @Slot(str)
    def on_log_output(self, message):
        """
        Appends standard output messages to the log area with white color.

        Args:
            message (str): The message to append.
        """
        self.log_text.append(
            f"<span style='color: white;'>{message.replace(chr(10), '<br>')}</span>"
        )

    @Slot(str)
    def on_log_error(self, message):
        """
        Appends standard error messages to the log area with red color.

        Args:
            message (str): The error message to append.
        """
        self.log_text.append(
            f"<span style='color: red;'>{message.replace(chr(10), '<br>')}</span>"
        )

    def show_summary(self):
        """
        Displays a summary of script execution results.
        """
        summary = "Script Execution Summary:\n\n"
        for script_name, result in self.script_results.items():
            summary += f"{script_name}: {result.capitalize()}\n"
        QMessageBox.information(self, "Summary", summary)

    def cancel_process(self):
        """
        Cancels the currently running script process.
        Note: Implementing cancellation requires additional handling in worker threads.
        Currently, this function just notifies the user.
        """
        QMessageBox.information(
            self,
            "Cancel",
            "Cancelling script execution is not supported in this refactored version.",
        )

    def handle_timeout(self):
        """
        Handles the event when a script execution times out.
        Note: Implementing timeouts requires additional handling in worker threads.
        Currently, this function is a placeholder.
        """
        QMessageBox.warning(
            self,
            "Timeout",
            "Script execution timed out.",
        )
        self.reset_execution_state()

    def show_execution_indicators(self):
        """
        Shows the loading animation, cancel button, and log area.
        """
        self.animation_label.show()
        self.animation_movie.start()
        self.cancel_button.show()
        self.log_text.show()
        self.log_label.show()

    def hide_execution_indicators(self):
        """
        Hides the loading animation, cancel button, and log area.
        """
        self.animation_movie.stop()
        self.animation_label.hide()
        self.cancel_button.hide()

    def reset_execution_state(self):
        """
        Resets the execution state by hiding indicators and re-enabling buttons.
        """
        self.set_buttons_enabled(True)
        self.hide_execution_indicators()
        self.log_text.hide()
        self.log_label.hide()

    def set_buttons_enabled(self, enabled):
        """
        Enables or disables all sub-category buttons.

        Args:
            enabled (bool): True to enable, False to disable.
        """
        for i in range(self.button_layout.count()):
            widget = self.button_layout.itemAt(i).widget()
            if isinstance(widget, ScriptButtonWidget):
                widget.button.setEnabled(enabled)
            elif isinstance(widget, QPushButton):
                widget.setEnabled(enabled)

    # ---------------------------------
    # Graph Functionality
    # ---------------------------------

    def on_click(self, event):
        """
        Handles click events on the plot to display data point values.

        Args:
            event: The matplotlib event.
        """
        if not self.df.empty and hasattr(self, "df"):
            if event.artist != self.canvas.ax.collections[0]:
                return

            ind = event.ind[0]
            x_column = self.x_axis_combo.currentText()
            y_column = self.y_axis_combo.currentText()

            try:
                x_value = self.df[x_column].iloc[ind]
                y_value = self.df[y_column].iloc[ind]
                QMessageBox.information(
                    self, "Data Point", f"X: {x_value}\nY: {y_value}"
                )
            except IndexError:
                QMessageBox.warning(
                    self, "Index Error", "Clicked point is out of range."
                )

    # ---------------------------------
    # Utility Methods
    # ---------------------------------

    def highlight_active_button(self, script_name):
        """
        Highlights the active script button and resets others.
        If script_name is None, reset all buttons.
        """
        for item in self.script_buttons:
            widget = self.script_buttons[item]
            if script_name is None:
                # Reset to default style and uncheck the button
                widget.button.setStyleSheet(
                    """
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
                        background-color: #207544;
                        color: white;
                    }
                """
                )
                widget.button.setChecked(False)
            elif item == script_name:
                # Highlight the active button
                widget.button.setStyleSheet(
                    """
                    QPushButton {
                        border: 1px solid #cccccc;
                        border-radius: 8px;
                        padding: 10px;
                        text-align: center;
                        color: white;
                    }
                    QPushButton:hover {
                        background-color: #89a4FF;
                    }
                """
                )

                if self.is_ability_mode:
                    widget.button.setStyleSheet(
                        """
                        QPushButton {
                            background-color: #7e07a0;
                        }
                        """
                    )
                else:
                    widget.button.setStyleSheet(
                        """
                        QPushButton {
                            background-color: #207544;
                        }
                        """
                    )

                widget.button.setChecked(True)
            else:
                # Reset other buttons to default style
                widget.button.setStyleSheet(
                    """
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
                        background-color: #207544;
                        color: white;
                    }
                """
                )
                widget.button.setChecked(False)

    def update_category_styles(self):
        """
        Updates the checked state of category buttons based on their expansion state.
        """
        self.daily_button.setChecked(self.expanded_categories["Daily"])
        self.weekly_button.setChecked(self.expanded_categories["Weekly"])
        self.monthly_button.setChecked(self.expanded_categories["Monthly"])

    def set_category_button_checked(self, category, checked):
        """
        Sets the checked state of a specific category button.

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

    # ---------------------------------
    # Script Execution Utilities
    # ---------------------------------

    def run_python_script(self, file_name):
        """
        Executes a Python script on a separate thread and displays its DataFrame output.

        Args:
            file_name (str): The name of the script file to execute.
        """
        # Placeholder implementation.
        # You need to define how to execute external Python scripts and retrieve DataFrame.

        def execute_external_script():
            try:
                # Execute the external script and capture its output
                # This example assumes the script returns a DataFrame stored in a variable named 'df'
                # Modify as per your actual script implementation

                # Import the script as a module
                script_path = self.get_resource_path(os.path.join("scripts", file_name))
                if not os.path.exists(script_path):
                    self.log_error.emit(f"Script file not found: {script_path}")
                    self.finished.emit(file_name, "failed")
                    return

                import importlib.util

                spec = importlib.util.spec_from_file_location(
                    "external_script", script_path
                )
                external_script = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(external_script)

                # Assume the script has a 'main' function that returns a DataFrame
                if hasattr(external_script, "main") and callable(external_script.main):
                    df = external_script.main()
                    if isinstance(df, pd.DataFrame):
                        self.display_dataframe(df)
                        self.log_output.emit(
                            f"Script '{file_name}' executed successfully."
                        )
                        self.finished.emit(file_name, "success")
                    else:
                        self.log_error.emit(
                            f"Script '{file_name}' did not return a DataFrame."
                        )
                        self.finished.emit(file_name, "failed")
                else:
                    self.log_error.emit(
                        f"Script '{file_name}' does not have a callable 'main' function."
                    )
                    self.finished.emit(file_name, "failed")

            except Exception as e:
                self.log_error.emit(str(e))
                self.finished.emit(file_name, "failed")

        # Create a QThread object
        thread = QThread()
        # Create a worker object
        worker = Worker()
        # Move worker to the thread
        worker.moveToThread(thread)
        # Connect signals and slots
        thread.started.connect(execute_external_script)
        worker.finished.connect(self.on_script_finished)
        worker.log_output.connect(self.on_log_output)
        worker.log_error.connect(self.on_log_error)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        # Start the thread
        thread.start()

        # Keep a reference to the thread to prevent garbage collection
        self.threads.append(thread)

    # ---------------------------------
    # Utility Methods (continued)
    # ---------------------------------

    def highlight_active_button(self, script_name):
        """
        Highlights the active script button and resets others.
        If script_name is None, reset all buttons.
        """
        for item in self.script_buttons:
            widget = self.script_buttons[item]
            if script_name is None:
                # Reset to default style and uncheck the button
                widget.button.setStyleSheet(
                    """
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
                        background-color: #207544;
                        color: white;
                    }
                """
                )
                widget.button.setChecked(False)
            elif item == script_name:
                # Highlight the active button
                widget.button.setStyleSheet(
                    """
                    QPushButton {
                        border: 1px solid #cccccc;
                        border-radius: 8px;
                        padding: 10px;
                        text-align: center;
                        color: white;
                    }
                    QPushButton:hover {
                        background-color: #89a4FF;
                    }
                """
                )

                if self.is_ability_mode:
                    widget.button.setStyleSheet(
                        """
                        QPushButton {
                            background-color: #7e07a0;
                        }
                        """
                    )
                else:
                    widget.button.setStyleSheet(
                        """
                        QPushButton {
                            background-color: #207544;
                        }
                        """
                    )

                widget.button.setChecked(True)
            else:
                # Reset other buttons to default style
                widget.button.setStyleSheet(
                    """
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
                        background-color: #207544;
                        color: white;
                    }
                """
                )
                widget.button.setChecked(False)

    def update_category_styles(self):
        """
        Updates the checked state of category buttons based on their expansion state.
        """
        self.daily_button.setChecked(self.expanded_categories["Daily"])
        self.weekly_button.setChecked(self.expanded_categories["Weekly"])
        self.monthly_button.setChecked(self.expanded_categories["Monthly"])

    def set_category_button_checked(self, category, checked):
        """
        Sets the checked state of a specific category button.

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

    # =============================================================================
    # Main Entry Point
    # =============================================================================

    def main_app_cleanup(self):
        """
        Cleans up all running threads before closing the application.
        """
        for thread in self.threads:
            if thread.isRunning():
                thread.quit()
                thread.wait()

    def closeEvent(self, event):
        """
        Overrides the close event to ensure proper thread cleanup.
        """
        self.main_app_cleanup()
        event.accept()


# =============================================================================
# Main Entry Point Function
# =============================================================================


def main():
    """
    The main entry point for the application.
    """
    app = QApplication(sys.argv)

    # Apply Global Button Style
    button_style = """
    QPushButton {
        background-color: #ffffff;
        border: 2px solid #cccccc;
        border-radius: 10px;
        padding: 10px;
        text-align: center;
        color: black;
        font-weight: bold;
        font-size: 14px;
    }
    QPushButton:hover {
        background-color: #dcdcdc;
    }
    QPushButton:checked {
        background-color: #207544;
        color: white;
    }
    """
    app.setStyleSheet(button_style)

    window = MainApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
