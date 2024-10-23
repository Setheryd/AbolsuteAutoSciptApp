# Third-Party Imports
import os
import win32com.client as win32  # type: ignore
import pandas as pd
# import mplcursors  # type: ignore
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
from PySide6.QtCore import Qt, QProcess, Slot, QTimer, QPropertyAnimation  # type: ignore
from PySide6.QtGui import QMovie, QPixmap, QPainter, QColor, QGuiApplication  # type: ignore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


# Categorized Items
ability_daily_items = sorted(
    [
        "Employee Birthday Email",
    ]
)
ability_daily_items.insert(0, "Run All")  # Ensure "Run All" is at the top

ability_weekly_items = sorted(
    [
        "Weekly Example",
    ]
)
ability_weekly_items.insert(0, "Run All")  # Ensure "Run All" is at the top

ability_monthly_items = sorted(
    [
        "Monthly Example",
    ]
)
ability_monthly_items.insert(0, "Run All")  # Ensure "Run All" is at the top

def setup_ability_mode_tabs(self):
    """
    Sets up the UI for the Ability mode with tabs on the right side.
    """

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
                background: #7e07a0;
                font-weight: bold;
                color: #d9e6f2;
                color: white;
                margin-top: 0px;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            
            QPushButton:checked {
                background-color: #7e07a0;
                color: white;
            }
        """

    ability_setup_scripts_tab(self)
    self.setup_dashboard_tab()
    self.setup_graph_tab()

    # Apply the custom style sheet to the QTabWidget
    self.tab_widget.setStyleSheet(tab_style)

    # Add more tabs or widgets specific to Ability mode

def ability_setup_scripts_tab(self):
    """
    Sets up the Scripts tab in Ability mode.
    """
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
            color: #7e07a0;
        }
    """)
    self.category_layout.addWidget(categories_header, alignment=Qt.AlignCenter)

    # Create Main Category Buttons as Toggles
    self.daily_button = self.create_category_button(
        "Daily", ability_daily_items, "daily"
    )
    self.weekly_button = self.create_category_button(
        "Weekly", ability_weekly_items, "weekly"
    )
    self.monthly_button = self.create_category_button(
        "Monthly", ability_monthly_items, "monthly"
    )

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
            color: #7e07a0; 
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
    loading_gif_path = os.path.join(script_dir, "../resources", "loading.gif")
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
    ability_update_category_styles(self)

    # Dictionary to Map Script Names to Their Paths
    ability_initialize_scripts_mapping(self)

    # Add Scripts Tab to QTabWidget
    self.tab_widget.addTab(self.scripts_tab, "Scripts")



def ability_setup_graph_tab(self):
    """
    Sets up the Graph tab in Ability mode.
    """
    graph_tab = QWidget()
    graph_layout = QVBoxLayout()

    graph_label = QLabel("Graph Tab Content")
    graph_layout.addWidget(graph_label)

    graph_tab.setLayout(graph_layout)
    self.tab_widget.addTab(graph_tab, "Graph")


def ability_initialize_scripts_mapping(self):
    """
    Maps script names to their corresponding file paths.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    self.scripts = {
        # Daily Scripts
        "Employee Birthday Email": os.path.join(
            script_dir, "../daily_tasks", "birthday.py"
        ),
        # Weekly Scripts
        # Monthly Scripts
    }


def ability_update_category_styles(self):
    """
    Updates the checked state of category buttons based on their expansion state.
    """
    self.daily_button.setChecked(self.expanded_categories["Daily"])
    self.weekly_button.setChecked(self.expanded_categories["Weekly"])
    self.monthly_button.setChecked(self.expanded_categories["Monthly"])

