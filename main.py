import tkinter as tk
import sys
import GUI
import Data_Store

app = tk.Tk()

gui_setup = GUI.GUI_setup(app)
gui_setup.setup_styles()
Data_Store.central_data_store.update_data('gui_setup', gui_setup)

data_1_instances = GUI.intro_class(gui_setup.intro_tab)

app.mainloop()
sys.exit()
