import sys
import os
sys.path.append(os.path.dirname(__file__))

from app import MessManagementApp

if __name__ == "__main__":
    app = MessManagementApp()
    app.mainloop()