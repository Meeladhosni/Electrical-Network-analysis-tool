import ast
import os
from tkinter import filedialog


class CentralDataStore:
    def __init__(self):
        self.data = {}

    def update_data(self, key, value):
        self.data[key] = value

    def get_data(self, key):
        return self.data.get(key)


# Global instance of the data store
central_data_store = CentralDataStore()


class DataLdr:
    def __init__(self, app):
        self.app = app
        self.file_path = None
        self.data = {}

    @staticmethod
    def process_line(line):
        try:
            key, value = line.split('=', 1)
            return key.strip(), ast.literal_eval(value.strip())
        except (ValueError, SyntaxError):
            return None

    def choose_file(self):
        # Set the initial directory path
        current_script_path = os.path.dirname(os.path.abspath(__file__))
        initial_dir = os.path.join(current_script_path, 'Saved data')

        file_path = filedialog.askopenfilename(title="Select a file",
                                               initialdir=initial_dir,
                                               filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        if file_path:
            self.file_path = file_path

    def read_file(self):
        if not self.file_path:
            return "Please choose a file first."

        with open(self.file_path, 'r') as file:
            next(file)
            for line in file:
                result = self.process_line(line)
                if result:
                    key, value = result
                    self.data[key] = value
            return self.data

class place_font_reader:
    place_font_info_file_path = os.path.join('config', 'place_font_ref_info.txt')

    def __init__(self):
        self.file_path = self.place_font_info_file_path
        self.ref_values = []

    def ref_info(self):
        with open(self.file_path, "r") as file:
            for line in file:
                parts = line.strip().split('=')
                if len(parts) == 2:
                    value = parts[1].strip()
                    self.ref_values.append(value)
        return self.ref_values


place_font_info_reader = place_font_reader()
ref_values = place_font_info_reader.ref_info()
ref_x = int(ref_values[0])
ref_y = int(ref_values[1])
ref_f = int(ref_values[2])

central_data_store.update_data('saved_x_value', ref_x)
central_data_store.update_data('saved_y_value', ref_y)
central_data_store.update_data('saved_f_value', ref_f)
