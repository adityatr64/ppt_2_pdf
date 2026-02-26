import tkinter as tk
from views import MainView, setup_windows_taskbar
from controllers import AppController


def main():
    setup_windows_taskbar()
    
    root = tk.Tk()
    
    view = MainView(root)
    controller = AppController(view)
    
    root.mainloop()


if __name__ == "__main__":
    main()
