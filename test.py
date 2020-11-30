import win32com.client
import pythoncom

class PartModel():
    def __init__(self, input):
        model = sw.ActiveDoc
        modelExt = model.Extension
        selMgr = model.SelectionManager
        featureMgr = model.FeatureManager
        sketchMgr = model.SketchManager
        eqMgr = model.GetEquationMgr
        ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
        length = 1/39.3701
        bore_size = 0.5/39.3701
        bore_placement = 1.5/39.3701
        height = bore_size + bore_placement
        thickness = 0.2/39.3701
        arc_rad = 1.25/39.3701
        L_bracket_length = 1/39.3701

    def connect_sldworks(self):
        swYearLastDigit = 9
        sw = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(swYearLastDigit-2)))  

    def new_part_doc(self):
        pass

    def create_simple_tab(self):
        pass

    def eqmgr_simple_tab(self):
        pass
    
    def create_l_bracket(self):
        pass

    def eqmgr_l_bracket(self):
        pass

    def insert_plane(self):
        pass
    
    def insert_sketch(self):
        pass
    
    def select_by_id(self):
        pass

    def select_by_id2(self):
        pass

    def create_corner_rectangle(self):
        pass

    def create_circle(self):
        pass

    def create_circle_by_perimeter(self):
        pass

    def extrude_cut(self):
        pass

    def boss_extrude(self):
        pass

    def clear_selection(self):
        pass



class PartView():
    def __init__(self,part):
        pass

    def display_template(self):
        pass


class PartController():
    def __init__(self):
        pass

    def modify_dims(self):
        pass
    
    def save_part(self):
        pass

    def check_connection(self):
        sw.SendMsgToUser("Hello world! SOLIDWORKS API!")

    def get_input(self):
        pass

def main(self):
    model = PartModel()
    view = View(model)
    controller = Controller()

    # steps:
    # 1) connect with pythoncom and open new part file
    # 2) prompt for type of bracket
    # 3) create bracket normally
    # 4) use equation manager and global variables to definte the part parametrically
    # 5) prompt user for custom dimensions
    # 6) continue until user is done
    # 7) allow user to quit and save part file
    pass
