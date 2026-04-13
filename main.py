import customtkinter as ctk
from app import NutritionApp

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

if __name__ == "__main__":
    root = ctk.CTk()
    app = NutritionApp(root)
    root.mainloop()
