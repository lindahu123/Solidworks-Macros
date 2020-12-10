import win32com.client
import pythoncom
from PIL import Image

swYearLastDigit = 9
sw = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(swYearLastDigit-2))) 

sw.newpart
model = sw.ActiveDoc
modelExt = model.Extension
selMgr = model.SelectionManager
featureMgr = model.FeatureManager
sketchMgr = model.SketchManager
eqMgr = model.GetEquationMgr
ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


class PartModel():
    def __init__(self,length,bore_size,bore_placement,width,arc_rad,L_bracket_length= None):
        #devision to convert from inches to meters
        self.length = length/39.3701
        self.bore_size = 2* bore_size/39.3701
        self.bore_placement = bore_placement/39.3701
        self.height = self.bore_size + self.bore_placement
        self.width = width/39.3701
        self.arc_rad = arc_rad/39.3701
        if L_bracket_length == None:
            self.L_bracket_length = 0
        else:
            self.L_bracket_length = L_bracket_length/39.3701
    
    def add_all_global_variables(self):
        eqMgr.Add2(-1,self.create_global_eq("length",self.length), True) 
        eqMgr.Add2(-1,self.create_global_eq("width",self.width), True) 
        eqMgr.Add2(-1,self.create_global_eq("bore size",self.bore_size), True)
        eqMgr.Add2(-1,self.create_global_eq("bore placement",self.bore_placement), True) 
        eqMgr.Add2(1, "\"height\" = \"bore size\"+ \"bore placement\"", True)
        eqMgr.Add2(-1,self.create_global_eq("arc rad",self.arc_rad), True) 
        eqMgr.Add2(-1,self.create_global_eq("L bracket length",self.L_bracket_length), True) 

    def create_simple_tab(self):
        #make a new sketch on the Front Plane
        self.select_by_id2("Front Plane","PLANE")
        self.insert_sketch()
        #sketch and dimension main tab shape
        self.create_corner_rectangle(self.length,self.height)
        self.create_circle(self.length,self.bore_placement,self.bore_size)
        self.select_by_id2("Line1", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Line2", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line1", "SKETCHSEGMENT",True)
        model.AddVerticalDimension2(0, 0, 0)
        self.clear_selection()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line2", "SKETCHSEGMENT",True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        modelExt.AddDimension(0, 0.001, 0, 0)
        self.select_by_id2("Sketch1", "SKETCH")
        self.boss_extrude(self.width)
        self.select_by_id2("Front Plane","PLANE")
        self.insert_sketch()
        self.create_circle_by_perimeter(self.length, self.arc_rad)
        self.clear_selection()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        modelExt.AddDimension(0, 0.001, 0, 0)
        self.select_by_id2("Top Plane", "PLANE")
        self.select_by_id2("Arc1", "SKETCHSEGMENT", True)
        model.AddVerticalDimension2(0, 0, 0)
        self.select_by_id2("Right Plane", "PLANE")
        self.select_by_id2("Arc1", "SKETCHSEGMENT", True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.extrude_cut()

    def simple_tab_equations(self):
        eqMgr.Add2(-1, self.create_eq("D1@Sketch1","length"), True) 
        eqMgr.Add2(-1, self.create_eq("D2@Sketch1","height"), True) 
        eqMgr.Add2(-1, self.create_eq("D3@Sketch1","bore placement"), True)
        eqMgr.Add2(-1, self.create_eq("D5@Sketch1","bore size"), True)
        eqMgr.Add2(-1, self.create_eq("D4@Sketch1","length","/2"), True)

    def create_l_bracket(self):
        self.select_by_id2("Front Plane","PLANE")
        self.insert_sketch()
        self.create_corner_rectangle(self.length,self.height)
        self.create_circle(self.length,self.bore_placement,self.bore_size)
        self.select_by_id2("Sketch1", "SKETCH")
        self.boss_extrude(self.width)
        self.clear_selection()
        self.select_by_id2("Front Plane","PLANE")
        self.insert_plane(self.thickness)
        self.create_corner_rectangle(self.length, self.width)
        self.boss_extrude(self.L_bracket_length)

    def insert_plane(self,width):
        model.CreatePlaneAtOffset(width, 0)
    
    def insert_sketch(self):
        sketchMgr.InsertSketch(True)
        
    def select_by_id2(self,plane,obj_type,append = False):
        modelExt.SelectByID2(plane, obj_type, 0, 0, 0, append, 0, ARG_NULL, 0)

    def create_corner_rectangle(self,length,height):
        sketchMgr.CreateCornerRectangle(0, 0, 0, length, height, 0)

    def create_circle(self,length,bore_placement,bore_size):
        sketchMgr.CreateCircle(length/2, bore_placement,0,length/2,bore_placement+bore_size/2, 0)

    def create_circle_by_perimeter(self,length, arc_rad):
        sketchMgr.PerimeterCircle(0, 0, length/2, arc_rad - length, length, 0)

    def extrude_cut(self):
        featureMgr.FeatureCut3(False, False, False, 1, 0, 100, 100, False, False, False, False, 0, 0, False, False, False, False, False, True, True, False, False, False, 0, 0, False)

    def boss_extrude(self,width):
        featureMgr.FeatureExtrusion2(True,False,False,0,0,width,0.001,False,False,False,False,0,0,False,False,False,False,True,True,True,0,0, False)

    def clear_selection(self):
        model.ClearSelection2(True)

    def add_dimension(self):
        modelExt.AddDimension(0, 0, 0, 0)

    def create_global_eq(self, variable_name,value):
        return "\"{0}\" = {1}".format(variable_name, value)

    def create_eq(self,dimension,variable_name,math = ""):
        return "\"{0}\" =  \"{1}\"{2}".format(dimension, variable_name,math)


class PartView():
    def __init__(self):
        pass

    def display_template(self,file):  
        display(Image.open(file))


class PartController():
    def __init__(self):
        pass
        
    def modify_dims(self):
        #steps
        # 1) find index of equation(s)
        # 2) delete equations
        # 3) find index of global variable
        # 4) delete variable
        # 5) re enter new variable value
        # 6) re enter equations
        # 7) rebuild document
        model.EditRebuild3
    
    def save_part(self):
        pass

    def check_connection(self):
        sw.SendMsgToUser("Hello world! SOLIDWORKS API!")

    def start_input(self):
        type_input = input(
            "Enter tab type (S for simple tab, L for L-bracket, c to check connection): ")
        stripped_input = type_input.strip()
        return stripped_input

    def initial_dimension_input(self):
        dimension_list = ["length","bore_size","bore_placement","width","notched_radius"]
        dimension_inputs = []
        for dimension in dimension_list:
            type_input = input(dimension + ":")
            dimension_inputs.append(float(type_input))
        return dimension_inputs

    def quit(self):
        # sys.exit(0)
        pass
    
    def equation_index_dictionary(self):
        i = 0
        equation_dictionary = {}
        while eqMgr.Equation(i) != '':
            dict1[eqMgr.Equation(i)] = i
            i +=1

    def dimension_change_input(self):
        type_input = input(
            "Dimension to change, q to quit, or s to save")
        stripped_input = type_input.strip()
        if stripped_input == "q":
            self.quit()
        if stripped_input == "s":
            pass

def main():
    #steps
    # 1) connect with pythoncom and open new part file
    part_controller = PartController()
    part_view = PartView()
    # 2) prompt for type of bracket and dimensions
    stripped_input = part_controller.start_input()
    if stripped_input == "S":
        part_view.display_template('simpletab.JPG')
        dimension_inputs = part_controller.initial_dimension_input()
        # 3) create bracket normally
        part_model = PartModel(dimension_inputs[0],dimension_inputs[1],dimension_inputs[2],dimension_inputs[3],dimension_inputs[4])
        part_model.create_simple_tab()
        # 4) use equation manager and global variables to definte the part parametrically
        part_model.add_all_global_variables()
        part_model.simple_tab_equations()
    elif stripped_input == "L":
        part_view.display_template('Ltab.JPG')
        part_model = PartModel(length,bore_size,bore_placement,thickness,arc_rad,L_bracket_length =  None)
        part_model.create_l_bracket()
    elif stripped_input == "c":
        part_controller.check_connection()
        part_controller.start_input()
    # 5) prompt user for custom dimensions
    part_controller.dimension_change_input()
    
    # 6) allow user to escape or save part file

main()