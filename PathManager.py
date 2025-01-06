import os
import sys

class locationManager:
    def __init__(self, current_path = ''):
        self.current_loc = self.get_current_loc() if not current_path else current_path
        self.input_path = os.path.join(self.current_loc, 'Input')
        self.output_path = os.path.join(self.current_loc,'Output')
        self.reference_path = os.path.join(self.current_loc, 'MasterReference.xlsx')
        self.archive_path = os.path.join(self.current_loc, 'Archive')
    
    def get_current_loc(self):
        """Returns filepath of this program."""
        return os.path.dirname(os.path.abspath(__file__))#os.path.dirname(sys.executable)
        

    def get_input_path(self):
        return self.input_path

    def get_output_path(self):
        return self.output_path

    def get_reference_path(self):
        return self.reference_path

    def get_archive_path(self):
        return self.archive_path