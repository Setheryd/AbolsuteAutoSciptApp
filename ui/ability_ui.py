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
    QProgressBar,
    QGraphicsBlurEffect,
)
from PySide6.QtCore import Qt, QProcess, Slot, QTimer, QPropertyAnimation  # type: ignore
from PySide6.QtGui import QMovie, QPixmap, QPainter, QColor, QGuiApplication  # type: ignore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

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
                background: purple;
                font-weight: bold;
                color: #d9e6f2;
                color: white;
                margin-top: 0px;
            }

            QTabBar::tab:!selected {
                margin-top: 2px;
            }
        """
   
    

    ability_setup_scripts_tab(self)
    ability_setup_dashboard_tab(self)
    ability_setup_graph_tab(self)

    # Apply the custom style sheet to the QTabWidget
    self.tab_widget.setStyleSheet(tab_style)

    # Add more tabs or widgets specific to Ability mode

def ability_setup_scripts_tab(self):
    """
    Sets up the Scripts tab in Ability mode.
    """
    scripts_tab = QWidget()
    scripts_layout = QVBoxLayout()

    scripts_label = QLabel("Scripts Tab Content")
    scripts_layout.addWidget(scripts_label)

    scripts_tab.setLayout(scripts_layout)
    self.tab_widget.addTab(scripts_tab, "Scripts")


def ability_setup_dashboard_tab(self):
    """
    Sets up the Dashboard tab in Ability mode.
    """
    dashboard_tab = QWidget()
    dashboard_layout = QVBoxLayout()

    dashboard_label = QLabel("Dashboard Tab Content")
    dashboard_layout.addWidget(dashboard_label)

    dashboard_tab.setLayout(dashboard_layout)
    self.tab_widget.addTab(dashboard_tab, "Dashboard")


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
