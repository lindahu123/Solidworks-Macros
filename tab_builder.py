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
    def __init__(self,length,bore_size,bore_placement,width, height, arc_rad, L_bracket_length):
        #devision to convert from inches to meters
        self.length = length/39.3701
        self.bore_size = bore_size/39.3701
        self.bore_placement = bore_placement/39.3701
        self.height = height/39.3701
        self.width = width/39.3701
        if arc_rad is not None:
            self.arc_rad = arc_rad/39.3701
        if L_bracket_length is not None:
            self.L_bracket_length = L_bracket_length/39.3701
 
    def add_all_global_variables(self):
        eqMgr.Add2(-1,self.create_global_eq("length",self.length), True) 
        eqMgr.Add2(-1,self.create_global_eq("width",self.width), True) 
        eqMgr.Add2(-1,self.create_global_eq("bore size",self.bore_size), True)
        eqMgr.Add2(-1,self.create_global_eq("bore placement",self.bore_placement), True) 
        eqMgr.Add2(-1,self.create_global_eq("height",self.height), True)
        if hasattr(self, 'arc_rad'):
            eqMgr.Add2(-1,self.create_global_eq("arc rad",self.arc_rad), True) 
        if hasattr(self, 'L_bracket_length'):
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
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude1","width"), True)
        eqMgr.Add2(-1, self.create_eq("D3@Sketch2","length","/2"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Sketch2","arc rad"), True)
        eqMgr.Add2(-1, self.create_eq("D2@Sketch2","arc rad","- \"length\""), True)

    def create_l_bracket(self):
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
        self.clear_selection()
        self.select_by_id2("Front Plane", "PLANE")
        model.CreatePlaneAtOffset(self.width, 0)
        self.insert_sketch()
        self.create_corner_rectangle(self.length, self.width)
        self.select_by_id2("Line2", "SKETCHSEGMENT")
        self.select_by_id2("Line4", "SKETCHSEGMENT",True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.select_by_id2("Line1", "SKETCHSEGMENT")
        self.select_by_id2("Line3", "SKETCHSEGMENT",True)
        model.AddVerticalDimension2(0, 0, 0)
        self.boss_extrude(self.L_bracket_length)

    def l_bracket_equations(self):
        eqMgr.Add2(-1, self.create_eq("D1@Sketch1","length"), True) 
        eqMgr.Add2(-1, self.create_eq("D2@Sketch1","height"), True) 
        eqMgr.Add2(-1, self.create_eq("D3@Sketch1","bore placement"), True)
        eqMgr.Add2(-1, self.create_eq("D5@Sketch1","bore size"), True)
        eqMgr.Add2(-1, self.create_eq("D4@Sketch1","length","/2"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude1","width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Plane1","width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Sketch2","length"), True) 
        eqMgr.Add2(-1, self.create_eq("D2@Sketch2","width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude2","L bracket length"), True)

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
    
    def save_part(self):
        pass

    def check_connection(self):
        sw.SendMsgToUser("Hello world! SOLIDWORKS API!")

    def start_input(self):
        type_input = input(
            "Enter tab type (S for simple tab, L for L-bracket, c to check connection): ")
        stripped_input = type_input.strip()
        return stripped_input

    def initial_dimension_input(self,tab_type):
        if tab_type == "S":
            dimension_list = ["length","bore_size","bore_placement","width","height","notched_radius"]
        if tab_type == "L":
            dimension_list = ["length","bore_size","bore_placement","width","height","L_bracket_length"]
        dimension_inputs = []
        for dimension in dimension_list:
            type_input = input(dimension + ":")
            dimension_inputs.append(float(type_input))
        return dimension_inputs

    def main_mod_dims(self):
        var_input = self.dimension_change_input()
        if type(var_input) != list:
            if var_input.strip() == "q":
                sw.CloseDoc("")
                return
            if var_input.strip() == "q":
                return
        variable = var_input[0].strip()
        value = var_input[1].strip()
        self.modify_dims(variable,value)
        self.main_mod_dims()

    def modify_dims(self,variable,value):
        #steps
        #get var from prompted input
        equations_to_modify = self.modify_equation_list(variable)
        # # 1) find index of equation(s)
        # # 2) delete equations
        self.delele_list_of_equations(equations_to_modify)
        global_var = self.global_var_finder(equations_to_modify,variable)
        equations_to_modify.remove(global_var)
        # 3) re enter new variable value
        value = float(value)
        value = value/39.3701
        eqMgr.Add2(-1,self.create_global_eq(variable,value), True) 
        # 4) re enter equations
        for eq in equations_to_modify:
            eqMgr.Add2(-1, eq, True) 
        # 5) rebuild document
        model.EditRebuild3

    def modify_equation_list(self,var):
        list_of_equations = self.equation_list_maker()
        equations_to_modify = []
        for i in list_of_equations:
            if var in i:
                equations_to_modify.append(i)
        return equations_to_modify

    def global_var_finder(self,list_of_eq,var):
        for i in list_of_eq:
            if var in i:
                if "@" not in i:
                    return i

    def delele_list_of_equations(self,list_equations):
        for equation in list_equations:
            dictionary = self.equation_index_dictionary()
            index = dictionary[equation]
            eqMgr.Delete(index)

    def equation_index_dictionary(self):
        i = 0
        equation_dictionary = {}
        while eqMgr.Equation(i) != '':
            equation_dictionary[eqMgr.Equation(i)] = i
            i +=1
        return equation_dictionary

    def equation_list_maker(self):
        i = 0
        list_of_equations = []
        while eqMgr.Equation(i) != '':
            list_of_equations.append(eqMgr.Equation(i))
            i +=1
        return list_of_equations

    def dimension_change_input(self):
        type_input = input(
            "m to modify dimension, q to quit, or s to save: ")
        stripped_input = type_input.strip()
        if stripped_input == "q":
            return "q"
        if stripped_input == "s":
            return "s"
        if stripped_input == "m":
            type_input = input(
            "enter new dimension equation from this list [height,bore size,bore placement,width,arc rad, L bracket length]: ")
            split = type_input.split("=")
            return split

    def create_global_eq(self, variable_name,value):
        return "\"{0}\" = {1}".format(variable_name, value)


def main():
    #steps
    # 1) connect with pythoncom and open new part file
    part_controller = PartController()
    part_view = PartView()
    # 2) prompt for type of bracket and dimensions
    stripped_input = part_controller.start_input()
    if stripped_input == "S":
        part_view.display_template('simpletab.JPG')
        dimension_inputs = part_controller.initial_dimension_input(stripped_input)
        # 3) create bracket normally
        part_model = PartModel(dimension_inputs[0],dimension_inputs[1],dimension_inputs[2],dimension_inputs[3],dimension_inputs[4],dimension_inputs[5],None)
        part_model.create_simple_tab()
        # 4) use equation manager and global variables to definte the part parametrically
        part_model.add_all_global_variables()
        part_model.simple_tab_equations()
    elif stripped_input == "L":
        part_view.display_template('Ltab.JPG')
        dimension_inputs = part_controller.initial_dimension_input(stripped_input)
        part_model = PartModel(dimension_inputs[0],dimension_inputs[1],dimension_inputs[2],dimension_inputs[3],dimension_inputs[4],None,dimension_inputs[5])
        part_model.create_l_bracket()
        part_model.add_all_global_variables()
        part_model.l_bracket_equations()
    elif stripped_input == "c":
        part_controller.check_connection()
        part_controller.start_input()
    # 5) prompt user for custom dimensions
    part_controller.main_mod_dims()
    # 6) allow user to escape or save part file

main()
