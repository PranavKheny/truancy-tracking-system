import sys
from PyQt6.QtWidgets import QApplication
from constructor import Student
from gui import TruancyApp

def main():
    #Init. Qt App
    app = QApplication(sys.argv)
    #create main window and display
    window = TruancyApp()
    window.show()
    #start Qt event loop and exit when done
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
