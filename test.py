import win32com.client
import pythoncom

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
    def __init__(self):
        self.length = 1/39.3701
        self.bore_size = 0.5/39.3701
        self.bore_placement = 1.5/39.3701
        self.height = self.bore_size + self.bore_placement
        self.thickness = 0.2/39.3701
        self.arc_rad = 1.25/39.3701
        self.L_bracket_length = 1/39.3701
    
    def add_all_global_variables(self):
        eqMgr.Add2(-1,self.create_global_eq("length",self.length), True) 
        eqMgr.Add2(-1,self.create_global_eq("thickness",self.thickness), True) 
        eqMgr.Add2(-1,self.create_global_eq("bore size",self.bore_size), True)
        eqMgr.Add2(-1,self.create_global_eq("bore placement",self.bore_placement), True) 
        eqMgr.Add2(-1,self.create_global_eq("height",self.height), True)
        eqMgr.Add2(-1,self.create_global_eq("arc rad",self.arc_rad), True) 
        eqMgr.Add2(-1,self.create_global_eq("L bracket length",self.L_bracket_length), True) 

    def create_simple_tab(self):
        self.select_by_id2("Front Plane","PLANE")
        self.insert_sketch()
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
        self.boss_extrude(self.thickness)
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
        self.boss_extrude(self.thickness)
        self.clear_selection()
        self.select_by_id2("Front Plane","PLANE")
        self.insert_plane(self.thickness)
        self.create_corner_rectangle(self.length, self.thickness)
        self.boss_extrude(self.L_bracket_length)

    def insert_plane(self,thickness):
        model.CreatePlaneAtOffset(thickness, 0)
    
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

    def boss_extrude(self,thickness):
        featureMgr.FeatureExtrusion2(True,False,False,0,0,thickness,0.001,False,False,False,False,0,0,False,False,False,False,True,True,True,0,0, False)

    def clear_selection(self):
        model.ClearSelection2(True)

    def add_dimension(self):
        modelExt.AddDimension(0, 0, 0, 0)

    def create_global_eq(self, variable_name,value):
        return "\"{0}\" = {1}".format(variable_name, value)

    def create_eq(self,dimension,variable_name,math = ""):
        return "\"{0}\" =  \"{1}\"{2}".format(dimension, variable_name,math)


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
    
    def equation_index_dictionary(self):
        i = 0
        equation_dictionary = {}
        while eqMgr.Equation(i) != '':
            dict1[eqMgr.Equation(i)] = i
            i +=1

def main(self):
    #steps
    # 1) connect with pythoncom and open new part file
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

    partmodel = PartModel()
    
    # 2) prompt for type of bracket
    # 3) create bracket normally
    partmodel.create_simple_tab()
    # 4) use equation manager and global variables to definte the part parametrically
    partmodel.add_all_global_variables()
    partmodel.simple_tab_equations()
    # 5) prompt user for custom dimensions
    # 6) continue until user is done
    # 7) allow user to quit and save part file
    