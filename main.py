# main.py

# =============================================================================
# Imports
# =============================================================================

# Standard Library Imports
import os
import sys
import subprocess
from datetime import datetime, timedelta
from io import StringIO

# Third-Party Imports
import win32com.client as win32  # type: ignore
import pandas as pd
import mplcursors  # type: ignore
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
)
from PySide6.QtCore import Qt, QProcess, Slot, QTimer  # type: ignore
from PySide6.QtGui import QMovie, QPixmap, QPainter, QColor, QGuiApplication  # type: ignore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# =============================================================================
# Constants
# =============================================================================

# Categorized Items
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
            self.status_label.setFixedSize(24, 24)  # Increased size for better visibility
            self.status_label.setAlignment(Qt.AlignCenter)
            self.status = 'pending'
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
        """
        Updates the status icon based on the current status.
        """
        if not self.has_status:
            return

        # Construct absolute path to the icon
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_mapping = {
            'pending': 'pending.png',
            'running': 'running.png',
            'success': 'success.png',
            'failed': 'failed.png'
        }
        icon_file = icon_mapping.get(self.status, '')
        icon_path = os.path.join(script_dir, 'resources', icon_file)

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
        """
        Sets the status and updates the status icon.

        Args:
            status (str): New status ('pending', 'running', 'success', 'failed').
        """
        self.status = status
        self.update_status_icon()


class MplCanvas(FigureCanvas):
    """
    Matplotlib Canvas to embed plots within the Qt application.
    """
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig, self.ax = plt.subplots(figsize=(width, height), dpi=dpi)
        super().__init__(self.fig)


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
                    row_content.append(item.text() if item else '')
                clipboard_content += "\t".join(row_content) + "\n"

        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_content.strip())

        QMessageBox.information(self, "Copied", "Selected data has been copied to the clipboard.")

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
        self.process = None
        self.selected_subcategory_button = None
        self.script_buttons = {}
        self.script_results = {}
        self.pending_scripts = []
        self.current_script_name = None

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

    def initialize_scripts_mapping(self):
        """
        Maps script names to their corresponding file paths.
        """
        script_dir = os.path.dirname(os.path.abspath(__file__))
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
            "Employee Attrition": os.path.join(script_dir, "monthly_tasks", "employee_attrition.py"),
            "Expired NOAs": os.path.join(script_dir, "monthly_tasks", "NOA_exp.py"),
            "Inventory Request": os.path.join(script_dir, "monthly_tasks", "inventory_request.py"),
            "Next Months Expired NOAs": os.path.join(script_dir, "monthly_tasks", "next_month_NOA_exp.py"),
            "Patient Attrition": os.path.join(script_dir, "monthly_tasks", "patient_attrition.py"),
        }

    def initialize_ui(self):
        """
        Initializes and sets up all UI components.
        """
        # Define the custom style sheet for tabs
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
                background: #A0C4FF; /* Same as the header color */
                font-weight: bold;
                color: #d9e6f2; /* Slightly darkened text color */
                color: white;
                margin-top: 0px;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            """

        # Create a QTabWidget for switching between tabs
        self.tab_widget = QTabWidget(self)

        # Apply the custom style sheet to the QTabWidget
        self.tab_widget.setStyleSheet(tab_style)

        # Setup Individual Tabs
        self.setup_scripts_tab()
        self.setup_dashboard_tab()
        self.setup_graph_tab()

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.tab_widget)
        self.setLayout(main_layout)

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
        categories_header.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #A0C4FF;
            }
        """)
        self.category_layout.addWidget(categories_header, alignment=Qt.AlignCenter)

        # Create Main Category Buttons as Toggles
        self.daily_button = self.create_category_button("Daily", daily_items, "daily")
        self.weekly_button = self.create_category_button("Weekly", weekly_items, "weekly")
        self.monthly_button = self.create_category_button("Monthly", monthly_items, "monthly")

        self.category_layout.addWidget(self.daily_button)
        self.category_layout.addWidget(self.weekly_button)
        self.category_layout.addWidget(self.monthly_button)

        # Add Stretch to Push Buttons to the Top
        self.category_layout.addStretch()  # Corrected: Add stretch to the layout, not the button

        # Container for Category Buttons
        self.category_container = QWidget()
        self.category_container.setLayout(self.category_layout)
        self.category_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        scripts_layout.addWidget(self.category_container)

        # Scripts Buttons Layout
        self.script_layout = QVBoxLayout()
        self.script_layout.setSpacing(10)
        self.script_layout.setAlignment(Qt.AlignTop)

        # "Scripts" Header
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

        # Add Components to Indicators Layout
        self.indicator_layout.addWidget(self.animation_label, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.cancel_button, alignment=Qt.AlignCenter)
        self.indicator_layout.addWidget(self.log_label, alignment=Qt.AlignLeft)
        self.indicator_layout.addWidget(self.log_text, alignment=Qt.AlignCenter)

        # Spacer to Push Indicators to the Top
        self.indicator_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Container for Indicators
        self.indicator_container = QWidget()
        self.indicator_container.setLayout(self.indicator_layout)
        self.indicator_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        scripts_layout.addWidget(self.indicator_container)

        # State to Track Expanded Categories
        self.expanded_categories = {
            "Daily": False,
            "Weekly": False,
            "Monthly": False
        }

        # Initialize Category Styles
        self.update_category_styles()

        # Dictionary to Map Script Names to Their Paths
        self.initialize_scripts_mapping()

        # Add Scripts Tab to QTabWidget
        self.tab_widget.addTab(self.scripts_tab, "Scripts")

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
                if self.current_category == "Weekly":
                    button.clicked.connect(self.run_all_weekly_items)
                elif self.current_category == "Daily":
                    button.clicked.connect(self.run_all_daily_items)
                elif self.current_category == "Monthly":
                    button.clicked.connect(self.run_all_monthly_items)
                else:
                    button.clicked.connect(self.show_message)
                script_widget = ScriptButtonWidget(item, button, has_status=False)
            else:
                # Configure Individual Script Button
                if item in self.scripts:
                    script_path = self.scripts[item]
                    button.clicked.connect(
                        lambda checked, path=script_path, name=item: self.run_single_script(path, name)
                    )
                else:
                    button.clicked.connect(self.show_message)
                script_widget = ScriptButtonWidget(item, button, has_status=True)
                self.script_buttons[item] = script_widget

            # Apply Stylesheet
            button.setStyleSheet("""
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
            button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            # Connect Selection Handler
            button.clicked.connect(lambda checked, btn=button: self.select_subcategory(btn))

            # Add to Layout
            self.button_layout.addWidget(script_widget, alignment=Qt.AlignLeft)

    def select_subcategory(self, button):
        """
        Handles the selection of a sub-category button, ensuring only one is checked.

        Args:
            button (QPushButton): The button that was clicked.
        """
        if self.selected_subcategory_button and self.selected_subcategory_button != button:
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
        custom_programs = {
            "Active Patients per Month": "Admission_by_Month.py",
            "View Patient Data": "patient_data_extractor.py",
            "Active Contractor per Month": "Caregiver_by_Month.py",
            "View Contractor Data": "caregiver_data_extractor.py",
            "Generate Report 4": "Report4.py",
            "Generate Report 5": "Report5.py",
            "Run Custom Report 6": "Report6.py"
        }

        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)

        for button_label, program in custom_programs.items():
            button = QPushButton(button_label)
            button.clicked.connect(lambda checked, prog=program: self.run_python_script(prog))
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
        Sets up the right-side layout for displaying the DataFrame.

        Args:
            splitter (QSplitter): The splitter to add the right-side widget to.
        """
        self.table_widget = CustomTableWidget()
        self.table_widget.setColumnCount(0)
        self.table_widget.setRowCount(0)
        self.table_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)  # Allow multiple selections
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectItems)  # Item selection behavior

        # Apply styles to show cell selection clearly
        self.table_widget.setStyleSheet("""
            QTableWidget {
                background-color: #f0f0f0;
                color: black;
            }
            QTableWidget::item:selected {
                background-color: #3498db; /* Custom color for selected cells */
                color: white; /* Text color for selected cells */
                border: 2px solid #2980b9; /* Optional border for selected cells */
            }
            QTableWidget::item {
                background-color: white; /* Normal cell background */
                color: black; /* Normal cell text color */
            }
        """)

        splitter.addWidget(self.table_widget)

    def setup_dashboard_bottom_buttons(self):
        """
        Sets up the bottom buttons (Save DataFrame, Copy Selected) for the Dashboard tab.
        """
        bottom_buttons_layout = QHBoxLayout()
        bottom_buttons_layout.setAlignment(Qt.AlignCenter)

        # Save DataFrame Button
        self.save_dataframe_button = QPushButton("Save DataFrame")
        self.save_dataframe_button.clicked.connect(self.save_dataframe)

        # Copy Selected Button
        self.copy_selected_button = QPushButton("Copy Selected")
        self.copy_selected_button.clicked.connect(self.table_widget.copy_selected_data)

        # Add Buttons to Layout
        bottom_buttons_layout.addWidget(self.save_dataframe_button)
        bottom_buttons_layout.addWidget(self.copy_selected_button)

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

        # X-axis Selection
        self.x_axis_label = QLabel("Select X-axis:")
        self.x_axis_label.setStyleSheet("""
            QLabel {
                font-size: 14pt;
                font-weight: bold;
                color: #A0C4FF;
            }
        """)
        self.x_axis_combo = QComboBox(self)
        self.x_axis_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                font-size: 12pt;
            }
        """)
        controls_layout.addWidget(self.x_axis_label)
        controls_layout.addWidget(self.x_axis_combo)

        # Y-axis Selection
        self.y_axis_label = QLabel("Select Y-axis:")
        self.y_axis_label.setStyleSheet("""
            QLabel {
                font-size: 14pt;
                font-weight: bold;
                color: #A0C4FF;
            }
        """)
        self.y_axis_combo = QComboBox(self)
        self.y_axis_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                font-size: 12pt;
            }
        """)
        controls_layout.addWidget(self.y_axis_label)
        controls_layout.addWidget(self.y_axis_combo)

        # Label SpinBox
        self.label_spinbox_label = QLabel("Number of X-axis labels to display:")
        self.label_spinbox_label.setStyleSheet("""
            QLabel {
                font-size: 14pt;
                font-weight: bold;
                color: #A0C4FF;
            }
        """)
        self.label_spinbox = QSpinBox(self)
        self.label_spinbox.setRange(1, 20)
        self.label_spinbox.setValue(6)
        self.label_spinbox.setSuffix(" labels")
        self.label_spinbox.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                font-size: 12pt;
            }
        """)
        controls_layout.addWidget(self.label_spinbox_label)
        controls_layout.addWidget(self.label_spinbox)

        # Spacer to Push Plot Button to Bottom
        controls_layout.addStretch()

        # Plot Graph Button
        self.plot_button = QPushButton("Plot Graph", self)
        self.plot_button.setFixedWidth(150)
        self.plot_button.setStyleSheet("""
            QPushButton {
                background-color: #A0C4FF;
                border: none;
                border-radius: 8px;
                padding: 10px;
                font-size: 14pt;
                color: white;
            }
            QPushButton:hover {
                background-color: #89A4FF;
            }
        """)
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

        # Matplotlib Canvas
        self.canvas = MplCanvas(self, width=5, height=4, dpi=100)
        self.canvas.setStyleSheet("""
            background-color: #2e2e2e;
        """)
        graph_area_layout.addWidget(self.canvas)

        # Add Graph Area Layout to Main Graph Layout
        graph_layout.addLayout(graph_area_layout, stretch=3)

        # Initialize an Empty Axes
        self.canvas.ax = self.canvas.figure.add_subplot(111)

    def apply_graph_styles(self):
        """
        Applies consistent styles to the graph area.
        """
        self.canvas.figure.patch.set_facecolor('#2e2e2e')  # Dark background
        self.canvas.ax.set_facecolor('#2e2e2e')  # Dark background for plot area
        self.canvas.ax.tick_params(axis='x', colors='white')  # White ticks
        self.canvas.ax.tick_params(axis='y', colors='white')
        self.canvas.ax.xaxis.label.set_color('white')  # White labels
        self.canvas.ax.yaxis.label.set_color('white')

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
        if not hasattr(self, 'df'):
            QMessageBox.warning(self, "Error", "No DataFrame to save.")
            return

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save DataFrame", "", "CSV Files (*.csv);;All Files (*)", options=options
        )

        if file_name:
            try:
                self.df.to_csv(file_name, index=False)
                QMessageBox.information(self, "Success", f"DataFrame saved successfully to {file_name}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save DataFrame: {str(e)}")

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
                (item, self.scripts[item]) for item in script_items if item in self.scripts
            ]

            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                self.script_buttons[item].set_status('pending')

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
                (item, self.scripts[item]) for item in script_items if item in self.scripts
            ]

            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                self.script_buttons[item].set_status('pending')

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
                (item, self.scripts[item]) for item in script_items if item in self.scripts
            ]

            self.script_results = {item: 'pending' for item, _ in self.pending_scripts}

            for item in self.script_buttons:
                if item in self.scripts:  # Ensure it's a monthly script
                    self.script_buttons[item].set_status('pending')

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
            next_script_name, next_script_path = self.pending_scripts.pop(0)
            self.current_script_name = next_script_name
            self.run_script(next_script_path, next_script_name)
        else:
            self.log_text.append("<span style='color: green;'>All scripts executed.</span>")
            self.show_summary()

    def run_script(self, script_path, script_name):
        """
        Executes a single script using QProcess.

        Args:
            script_path (str): The path to the script to execute.
            script_name (str): The name of the script.
        """
        try:
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                self.run_next_script()
                return

            self.script_results[script_name] = 'running'
            self.script_buttons[script_name].set_status('running')
            self.highlight_active_button(script_name)

            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            self.process.setArguments(['-u', script_path])

            # Connect Signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)

            # Start Timeout Timer
            self.timer.start()

            # Start Process
            self.process.start()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred:\n{str(e)}")
            self.run_next_script()

    def run_single_script(self, script_path, script_name):
        """
        Executes a single script and handles its execution state.

        Args:
            script_path (str): The path to the script to execute.
            script_name (str): The name of the script.
        """
        try:
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found at {script_path}")
                return

            # Reset Status Icons
            self.script_buttons[script_name].set_status('pending')
            self.script_results = {script_name: 'pending'}

            # Disable Sub-Category Buttons
            self.set_buttons_enabled(False)

            # Show Execution Indicators
            self.show_execution_indicators()
            self.log_text.clear()

            # Update Status to 'running'
            self.script_buttons[script_name].set_status('running')
            self.highlight_active_button(script_name)
            self.current_script_name = script_name

            # Initialize QProcess
            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            self.process.setArguments(['-u', script_path])

            # Connect Signals
            self.process.readyReadStandardOutput.connect(self.handle_stdout)
            self.process.readyReadStandardError.connect(self.handle_stderr)
            self.process.finished.connect(self.process_finished)

            # Start Timeout Timer
            self.timer.start()

            # Start Process
            self.process.start()

        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred while executing the script:\n{str(e)}")
            self.reset_execution_state()

    def handle_stdout(self):
        """
        Handles standard output from the running process and appends it to the log area.
        """
        if self.process:
            data = bytes(self.process.readAllStandardOutput()).decode('utf-8')
            formatted_data = data.replace('\n', '<br>')
            self.log_text.append(f"<span style='color: white;'>{formatted_data}</span>")

    def handle_stderr(self):
        """
        Handles standard error from the running process and appends it to the log area.
        """
        if self.process:
            data = bytes(self.process.readAllStandardError()).decode('utf-8')
            formatted_data = data.replace('\n', '<br>')
            self.log_text.append(f"<span style='color: red;'>{formatted_data}</span>")

    @Slot(int, QProcess.ExitStatus)
    def process_finished(self, exitCode, exitStatus):
        """
        Handles the completion of the script process.

        Args:
            exitCode (int): The exit code of the process.
            exitStatus (QProcess.ExitStatus): The exit status.
        """
        self.timer.stop()
        if exitCode == 0:
            self.log_text.append(f"<span style='color: green;'>Script '{self.current_script_name}' executed successfully.</span>")
            self.script_results[self.current_script_name] = 'success'
            self.script_buttons[self.current_script_name].set_status('success')
        else:
            self.log_text.append(f"<span style='color: red;'>Script '{self.current_script_name}' failed with exit code {exitCode}.</span>")
            self.script_results[self.current_script_name] = 'failed'
            self.script_buttons[self.current_script_name].set_status('failed')

        # Reset Button Highlights
        self.highlight_active_button(None)

        # Check if there are more scripts to run
        if self.pending_scripts:
            self.run_next_script()
        else:
            # All scripts have been executed
            self.hide_execution_indicators()
            self.log_text.append("<span style='color: green;'>All scripts executed.</span>")
            self.show_summary()

        # Re-enable Buttons only if no more scripts are running
        if not self.pending_scripts:
            self.set_buttons_enabled(True)

        # Uncheck the previously selected button
        if self.selected_subcategory_button:
            self.selected_subcategory_button.setChecked(False)
            self.selected_subcategory_button = None

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
        """
        if self.process and self.process.state() == QProcess.Running:
            self.process.kill()
            self.process = None
            self.timer.stop()
            self.pending_scripts = []
            self.log_text.append("<span style='color: orange;'>Script execution has been cancelled.</span>")

            # Update Status to 'failed'
            if self.current_script_name:
                self.script_results[self.current_script_name] = 'failed'
                self.script_buttons[self.current_script_name].set_status('failed')

            # Hide Execution Indicators
            self.hide_execution_indicators()

            # Re-enable Buttons
            self.set_buttons_enabled(True)

            # Show Summary
            self.show_summary()

    def handle_timeout(self):
        """
        Handles the event when a script execution times out.
        """
        if self.process and self.process.state() == QProcess.Running:
            self.process.kill()
            self.process = None
            self.timer.stop()
            self.pending_scripts = []
            self.log_text.append(f"<span style='color: orange;'>Script '{self.current_script_name}' timed out.</span>")

            # Update Status to 'failed'
            self.script_results[self.current_script_name] = 'failed'
            self.script_buttons[self.current_script_name].set_status('failed')

            # Hide Execution Indicators
            self.hide_execution_indicators()

            # Re-enable Buttons
            self.set_buttons_enabled(True)

            if self.pending_scripts:
                self.run_next_script()
            else:
                self.show_summary()

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

    def plot_graph(self):
        """
        Plots the graph with the selected X and Y axes.
        """
        if not hasattr(self, 'df') or self.df.empty:
            QMessageBox.warning(self, "No Data", "No data available to plot. Please run a script to load data.")
            return

        # Clear Previous Plot
        self.canvas.ax.clear()

        # Get Selected Columns
        x_column = self.x_axis_combo.currentText()
        y_column = self.y_axis_combo.currentText()

        if not x_column or not y_column:
            QMessageBox.warning(self, "Selection Error", "Please select both X and Y axes.")
            return

        # Get Number of Labels to Display
        max_labels = self.label_spinbox.value()

        try:
            x_data = self.df[x_column]
            y_data = self.df[y_column]

            # Scatter Plot
            scatter_plot = self.canvas.ax.scatter(x_data, y_data, picker=True, color='#A0C4FF')

            # Set Axis Labels
            self.canvas.ax.set_xlabel(x_column)
            self.canvas.ax.set_ylabel(y_column)

            # Configure X-axis Ticks and Labels
            num_points = len(x_data)
            step = max(1, num_points // max_labels)
            tick_indices = range(0, num_points, step)
            self.canvas.ax.set_xticks([x_data[i] for i in tick_indices])
            self.canvas.ax.set_xticklabels([x_data[i] for i in tick_indices], ha="right", color='white')

            # Configure Y-axis Tick Colors
            self.canvas.ax.tick_params(axis='y', colors='white')

            self.canvas.figure.tight_layout()

            # Redraw Canvas
            self.canvas.draw()

            # Connect Click Event
            self.canvas.mpl_connect("pick_event", self.on_click)

        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"An error occurred while plotting the graph:\n{str(e)}")

    def on_click(self, event):
        """
        Handles click events on the plot to display data point values.

        Args:
            event: The matplotlib event.
        """
        if event.artist != self.canvas.ax.collections[0]:
            return

        ind = event.ind[0]
        x_column = self.x_axis_combo.currentText()
        y_column = self.y_axis_combo.currentText()

        try:
            x_value = self.df[x_column].iloc[ind]
            y_value = self.df[y_column].iloc[ind]
            QMessageBox.information(self, "Data Point", f"X: {x_value}\nY: {y_value}")
        except IndexError:
            QMessageBox.warning(self, "Index Error", "Clicked point is out of range.")

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
                widget.button.setChecked(False)
            elif item == script_name:
                # Highlight the active button
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
                        background-color: #89a4FF;
                    }
                """)
                widget.button.setChecked(True)
            else:
                # Reset other buttons to default style
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

    def show_execution_indicators(self):
        """
        Displays the loading animation, cancel button, and log area.
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
        self.log_text.hide()
        self.log_label.hide()

    def reset_execution_state(self):
        """
        Resets the execution state by hiding indicators and re-enabling buttons.
        """
        self.set_buttons_enabled(True)
        self.hide_execution_indicators()
        self.log_text.hide()
        self.log_label.hide()

    # ---------------------------------
    # Script Running Method
    # ---------------------------------

    def run_python_script(self, file_name):
        """
        Executes a Python script and displays its DataFrame output.

        Args:
            file_name (str): The name of the script file to execute.
        """
        try:
            # Construct the full path to the script
            script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_extraction")
            script_path = os.path.join(script_dir, file_name)

            # Check if the script exists
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "Error", f"Script not found: {file_name}")
                return

            # Execute the script
            process = subprocess.Popen(
                [sys.executable, script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True  # Ensures the output is returned as a string
            )
            stdout, stderr = process.communicate()

            if process.returncode == 0:
                output = stdout.strip()
                print(f"Raw Output from {file_name}:\n{output}")

                try:
                    # Attempt to parse the output as a CSV formatted DataFrame
                    # Assuming the script output is CSV formatted (comma-separated)
                    df = pd.read_csv(StringIO(output))
                    self.display_dataframe(df)  # Assuming you have a method to display the DataFrame
                except pd.errors.ParserError as pe:
                    QMessageBox.critical(self, "Error", f"DataFrame Parsing error: {str(pe)}")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to parse DataFrame output: {str(e)}")
            else:
                error_message = stderr.strip()
                print(f"Error Output from {file_name}:\n{error_message}")
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Script '{file_name}' failed to execute.\nError: {error_message}"
                )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while executing the script: {str(e)}")


# =============================================================================
# Main Entry Point
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
        background-color: #a0c4ff;
        border: 2px solid #89a4ff;
        color: black;
    }
    """
    app.setStyleSheet(button_style)

    window = MainApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
