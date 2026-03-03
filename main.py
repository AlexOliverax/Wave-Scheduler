import sys
import json
import logging
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from app.ui import MainWindow
from app.utils import load_config

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("waves_scheduler.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Define icon paths relative to the executable / script location
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(BASE_DIR, "icon")
APP_ICON = os.path.join(ICON_DIR, "WavesS.ico")
LOGO_IMAGE = os.path.join(ICON_DIR, "WaveSpp.PNG")


def main():
    """
    Main entry point for the Waves Scheduler application.
    Loads configuration, initializes and displays the UI.
    """
    try:
        # Load configuration
        config = load_config()

        # Override icon / logo paths with paths relative to this executable
        config["app_icon"] = APP_ICON
        # Only set logo if it actually exists (avoids Qt warnings)
        if os.path.exists(LOGO_IMAGE):
            config["logo_image"] = LOGO_IMAGE
        elif "logo_image" in config:
            # Remove stale hardcoded paths that no longer exist
            del config["logo_image"]

        # Initialize application
        app = QApplication(sys.argv)
        app.setWindowIcon(QIcon(APP_ICON))
        app.setApplicationName("Waves Scheduler")
        app.setApplicationDisplayName("Waves Scheduler")

        window = MainWindow(config)
        window.show()

        # Execute application
        sys.exit(app.exec_())
    except Exception as e:
        logger.error(f"Application failed to start: {str(e)}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()