import customtkinter as ctk
from views import MainView, setup_windows_taskbar
from controllers import AppController


def main():
    setup_windows_taskbar()
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    
    root = ctk.CTk()
    
    view = MainView(root)
    controller = AppController(view)
    
    root.mainloop()


if __name__ == "__main__":
    main()
