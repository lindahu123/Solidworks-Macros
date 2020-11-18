import win32com.client
import pythoncom

swYearLastDigit = 9
sw = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(swYearLastDigit-2)))  

class PartModel():
    def __init__(self, input):
        self._file = input
        self.thickness = 1

    def open_sldworks(self):
        swYearLastDigit = 9
        sw = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(swYearLastDigit-2)))  

    def load_sldprtfile_or_create_sldprtfile(self):
        pass


class PartView():
    def __init__(self):
        pass

    def display_part(self):
        pass


class PartController():
    def __init__(self, length, height, bore):
        pass

    def modify_dims(self):
        pass

    def find_study(self):
        pass

    def modify_study(self):
        pass

    def run_study(self):
        pass

    def optimize_thickness(self):
        pass


class PartRun():
    def __init__(self):
        pass

    def main(self):
        # steps:
        # 1) open app and file
        # 2) modify dimensions
        # 3) set up correct study (might have different options depending on what direction
        # the force will be and how the load gets distributed)
        # 4) run study and analyze results
        # 5) rerun study with new thicknesses until a certain target metric is met
        # 6) display final part? or display throughout the while loop?
        pass
