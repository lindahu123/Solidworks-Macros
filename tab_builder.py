import win32com.client
import pythoncom
from PIL import Image

swYearLastDigit = 9
sw = win32com.client.Dispatch(
    "SldWorks.Application.%d" % (20+(swYearLastDigit-2)))

sw.newpart
model = sw.ActiveDoc
modelExt = model.Extension
selMgr = model.SelectionManager
featureMgr = model.FeatureManager
sketchMgr = model.SketchManager
eqMgr = model.GetEquationMgr
ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


class PartModel():
    """
    Create the part in cad, add in equations to be modified later
    """

    def __init__(self, length, bore_size, bore_placement, width,
                 height, arc_rad, L_bracket_length):
        # division to convert from inches to meters
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
        """
        Add in all global variables in Solidworks
        """
        eqMgr.Add2(-1, self.create_global_eq("length", self.length), True)
        eqMgr.Add2(-1, self.create_global_eq("width", self.width), True)
        eqMgr.Add2(-1, self.create_global_eq("bore size",
                                             self.bore_size), True)
        eqMgr.Add2(-1, self.create_global_eq("bore placement",
                                             self.bore_placement), True)
        eqMgr.Add2(-1, self.create_global_eq("height", self.height), True)
        if hasattr(self, 'arc_rad'):
            eqMgr.Add2(-1, self.create_global_eq("arc rad",
                                                 self.arc_rad), True)
        if hasattr(self, 'L_bracket_length'):
            eqMgr.Add2(-1, self.create_global_eq("L bracket length",
                                                 self.L_bracket_length), True)

    def create_simple_tab(self):
        """
        Create a simple tab in SolidWorks
        """
        # make a new sketch on the Front Plane
        self.select_by_id2("Front Plane", "PLANE")
        self.insert_sketch()
        # sketch and dimension main tab shape
        self.create_corner_rectangle(self.length, self.height)
        self.create_circle(self.length, self.bore_placement, self.bore_size)
        self.select_by_id2("Line1", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Line2", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line1", "SKETCHSEGMENT", True)
        model.AddVerticalDimension2(0, 0, 0)
        self.clear_selection()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line2", "SKETCHSEGMENT", True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        modelExt.AddDimension(0, 0.001, 0, 0)
        self.select_by_id2("Sketch1", "SKETCH")
        # extrude sketch
        self.boss_extrude(self.width)
        self.select_by_id2("Front Plane", "PLANE")
        # sketch and dimension sketch for cut
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
        # extrude cut
        self.extrude_cut()

    def simple_tab_equations(self):
        """
        Add in equations for simple tab in Solidworks
        """
        eqMgr.Add2(-1, self.create_eq("D1@Sketch1", "length"), True)
        eqMgr.Add2(-1, self.create_eq("D2@Sketch1", "height"), True)
        eqMgr.Add2(-1, self.create_eq("D3@Sketch1", "bore placement"), True)
        eqMgr.Add2(-1, self.create_eq("D5@Sketch1", "bore size"), True)
        eqMgr.Add2(-1, self.create_eq("D4@Sketch1", "length", "/2"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude1", "width"), True)
        eqMgr.Add2(-1, self.create_eq("D3@Sketch2", "length", "/2"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Sketch2", "arc rad"), True)
        eqMgr.Add2(-1, self.create_eq("D2@Sketch2",
                                      "arc rad", "- \"length\""), True)

    def create_l_bracket(self):
        """
        Create a L bracket in SolidWorks
        """
        # make a new sketch on the Front Plane
        self.select_by_id2("Front Plane", "PLANE")
        self.insert_sketch()
        # sketch and dimension main tab shape
        self.create_corner_rectangle(self.length, self.height)
        self.create_circle(self.length, self.bore_placement, self.bore_size)
        self.select_by_id2("Line1", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Line2", "SKETCHSEGMENT")
        self.add_dimension()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line1", "SKETCHSEGMENT", True)
        model.AddVerticalDimension2(0, 0, 0)
        self.clear_selection()
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        self.select_by_id2("Line2", "SKETCHSEGMENT", True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.select_by_id2("Arc1", "SKETCHSEGMENT")
        modelExt.AddDimension(0, 0.001, 0, 0)
        self.select_by_id2("Sketch1", "SKETCH")
        # extrude sketch
        self.boss_extrude(self.width)
        self.clear_selection()
        self.select_by_id2("Front Plane", "PLANE")
        # create new plane
        model.CreatePlaneAtOffset(self.width, 0)
        self.insert_sketch()
        # add rectangular sketch on new plane
        self.create_corner_rectangle(self.length, self.width)
        self.select_by_id2("Sketch2", "SKETCH")
        self.select_by_id2("Line2", "SKETCHSEGMENT", True)
        self.select_by_id2("Line4", "SKETCHSEGMENT", True)
        model.AddHorizontalDimension2(0, 0, 0)
        self.select_by_id2("Sketch2", "SKETCH")
        self.select_by_id2("Line1", "SKETCHSEGMENT", True)
        self.select_by_id2("Line3", "SKETCHSEGMENT", True)
        model.AddVerticalDimension2(0, 0, 0)
        # extrude sketch
        self.boss_extrude(self.L_bracket_length)

    def l_bracket_equations(self):
        """
        Add in L bracket equations in SolidWorks
        """
        eqMgr.Add2(-1, self.create_eq("D1@Sketch1", "length"), True)
        eqMgr.Add2(-1, self.create_eq("D2@Sketch1", "height"), True)
        eqMgr.Add2(-1, self.create_eq("D3@Sketch1", "bore placement"), True)
        eqMgr.Add2(-1, self.create_eq("D5@Sketch1", "bore size"), True)
        eqMgr.Add2(-1, self.create_eq("D4@Sketch1", "length", "/2"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude1", "width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Plane1", "width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Sketch2", "length"), True)
        eqMgr.Add2(-1, self.create_eq("D2@Sketch2", "width"), True)
        eqMgr.Add2(-1, self.create_eq("D1@Boss-Extrude2",
                                      "L bracket length"), True)

    def insert_plane(self, width):
        """
        insert plane at an offset from selected plane
        Args:
            width: distance of offset
        """
        model.CreatePlaneAtOffset(width, 0)

    def insert_sketch(self):
        """
        insert sketch at plane
        """
        sketchMgr.InsertSketch(True)

    def select_by_id2(self, obj, obj_type, append=False):
        """
        Select different objects in order to preform operations on after
        Args:
            obj: Object name
            obj_type: Object type (ex. PLANE,SKETCH,SKECTCHSEGMENT)
            append: True to keep old selection, False to only keep newest
        """
        modelExt.SelectByID2(obj, obj_type, 0, 0, 0, append, 0, ARG_NULL, 0)

    def create_corner_rectangle(self, length, height):
        """
        Create a corner rectangle
        Args:
            length: length of rectangle
            height: height of rectangle
        """
        sketchMgr.CreateCornerRectangle(0, 0, 0, length, height, 0)

    def create_circle(self, length, bore_placement, bore_size):
        """
        Create a center circle
        Args:
            length: 1/2 of the x coordinate of the center point
            bore_placement: the height of the center of the circle
            bore_size: diameter of circle
        """
        sketchMgr.CreateCircle(length/2, bore_placement,
                               0, length/2, bore_placement+bore_size/2, 0)

    def create_circle_by_perimeter(self, length, arc_rad):
        """
        Create a Perimeter circle
        Args:
            length: distance between intersection points
            arc_rad: diameter of circle
        """
        sketchMgr.PerimeterCircle(0, 0, length/2, arc_rad - length, length, 0)

    def extrude_cut(self):
        """
        Cuts for 100 m in both directions
        """
        featureMgr.FeatureCut3(False, False, False, 1, 0, 100, 100, False,
                               False, False, False, 0, 0, False, False,
                               False, False, False, True, True, False,
                               False, False, 0, 0, False)

    def boss_extrude(self, width):
        """
        Extrudes a sketch
        Args:
            width: the distance to extrude
        """
        featureMgr.FeatureExtrusion2(True, False, False, 0, 0, width, 0.001,
                                     False, False, False, False, 0, 0, False,
                                     False, False, False, True, True, True,
                                     0, 0, False)

    def clear_selection(self):
        """
        clears selected entities
        """
        model.ClearSelection2(True)

    def add_dimension(self):
        """
        Add dimension to a sketch segment
        """
        modelExt.AddDimension(0, 0, 0, 0)

    def create_global_eq(self, variable_name, value):
        """
        Create a global variable
        Args:
            variable_name: the name of the new global variable
            value: the value of that variable

        Return: global variable equation in the correct format
        """
        return "\"{0}\" = {1}".format(variable_name, value)

    def create_eq(self, dimension, variable_name, math=""):
        """
        Create a new equation
        Args:
            dimension: dimension name (ex. D1@Sketch1)
            variable_name: name of global variable to equate to
            math: divide or multiply the global variable
        Return: equation in the correct format
        """
        return "\"{0}\" =  \"{1}\"{2}".format(dimension, variable_name, math)


class PartView():
    """
    Display image of either simple tab or L bracket
    """

    def __init__(self):
        pass

    def display_template(self, file):
        """
        Display image
        Args:
            file: path to image
        """
        display(Image.open(file))


class PartController():
    """
    Prompt user for inputs and modify solidworks part
    """

    def __init__(self):
        pass

    def check_connection(self):
        """
        Send message to solidworks
        """
        sw.SendMsgToUser("Hello world! SOLIDWORKS API!")

    def start_input(self):
        """
        Get tab type input from user to start
        Return: tab type input
        """
        type_input = input(
            "Enter tab type (S for simple tab, L for L-bracket,\
c to check connection): ")
        stripped_input = type_input.strip()
        return stripped_input

    def initial_dimension_input(self, tab_type):
        """
        Get starting dimensions from user
        Args:
            tab_type: S for simple tab, L for L bracket
        Returns: dimension inputs as a list
        """
        if tab_type == "S":
            dimension_list = ["length", "bore_size",
                              "bore_placement", "width", "height",
                              "notched_radius"]
        if tab_type == "L":
            dimension_list = [
                "length", "bore_size", "bore_placement", "width", "height",
                "L_bracket_length"]
        dimension_inputs = []
        for dimension in dimension_list:
            type_input = input(dimension + ":")
            dimension_inputs.append(float(type_input))
        return dimension_inputs

    def main_mod_dims(self):
        """
        Main function to modify dimensions
        """
        # get input of which dimension needs to be changed
        var_input = self.dimension_change_input()
        if len(var_input) == 1:
            if "q" in var_input[0]:
                sw.CloseDoc("")
                return
            if "s" in var_input[0]:
                return
        # strip white spaces
        variable = var_input[0].strip()
        value = var_input[1].strip()
        # modify the variable and set to new value
        self.modify_dims(variable, value)
        # prompt user again until user wants to quit or save
        self.main_mod_dims()

    def modify_dims(self, variable, value):
        """
        Modify dimensions
        Args:
            variable: global variable being changed
            value: new value to assign to variable
        """
        # get variable from prompted input
        equations_to_modify = self.modify_equation_list(variable)
        # find index of equation(s)
        self.delete_list_of_equations(equations_to_modify)
        global_var = self.global_var_finder(equations_to_modify, variable)
        # delete equations
        equations_to_modify.remove(global_var)
        # re enter new variable value
        value = float(value)
        value = value/39.3701
        eqMgr.Add2(-1, self.create_global_eq(variable, value), True)
        # re enter equations
        for eq in equations_to_modify:
            eqMgr.Add2(-1, eq, True)
        # rebuild document
        model.EditRebuild3

    def modify_equation_list(self, var):
        """
        Get list of equations that need to be modified
        Args:
            var: global variable being changed
        Returns: list of equations affected by global variable
        """
        list_of_equations = self.equation_list_maker()
        equations_to_modify = []
        for i in list_of_equations:
            if var in i:
                equations_to_modify.append(i)
        return equations_to_modify

    def global_var_finder(self, list_of_eq, var):
        """
        Isolate global variable definition equation from a list of equations
        Args:
            var: global variable being changed
            list_of_eq: list to be parsed through
        Returns: equation that defines the global variable
        """
        for i in list_of_eq:
            if var in i:
                if "@" not in i:
                    return i

    def delete_list_of_equations(self, list_equations):
        """
        Delete all listed equations
        Args:
            list_equations: equations to be deleted
        """
        for equation in list_equations:
            dictionary = self.equation_index_dictionary()
            index = dictionary[equation]
            eqMgr.Delete(index)

    def equation_index_dictionary(self):
        """
        Make a dictionary that maps equation to the index of
        the equation in Solidworks

        Return: dictionary
        """
        i = 0
        equation_dictionary = {}
        while eqMgr.Equation(i) != '':
            equation_dictionary[eqMgr.Equation(i)] = i
            i += 1
        return equation_dictionary

    def equation_list_maker(self):
        """
        Make a list of all equations in Solidworks
        Return: the list
        """
        i = 0
        list_of_equations = []
        while eqMgr.Equation(i) != '':
            list_of_equations.append(eqMgr.Equation(i))
            i += 1
        return list_of_equations

    def dimension_change_input(self):
        """
        Prompt user for what dimension they need to modify
        Return: list [variable,value]
        """
        type_input = input(
            "m to modify dimension, q to quit, or s to save: ")
        stripped_input = type_input.strip()
        if stripped_input == "q":
            return ["q"]
        if stripped_input == "s":
            return ["s"]
        if stripped_input == "m":
            type_input = input("enter new dimension equation from this list\
[height, bore size, bore placement, width,\
arc rad, L bracket length]: ")
            split = type_input.split("=")
            return split

    def create_global_eq(self, variable_name, value):
        return "\"{0}\" = {1}".format(variable_name, value)


def main():
    part_controller = PartController()
    part_view = PartView()
    # prompt for type of bracket and dimensions
    stripped_input = part_controller.start_input()
    if stripped_input == "S":
        part_view.display_template('simpletab.JPG')
        dimension_inputs = part_controller.initial_dimension_input(
            stripped_input)
        # create bracket normally
        part_model = PartModel(dimension_inputs[0], dimension_inputs[1],
                               dimension_inputs[2], dimension_inputs[3],
                               dimension_inputs[4], dimension_inputs[5], None)
        part_model.create_simple_tab()
        # use equations to define the part parametrically
        part_model.add_all_global_variables()
        part_model.simple_tab_equations()
    elif stripped_input == "L":
        part_view.display_template('Ltab.JPG')
        dimension_inputs = part_controller.initial_dimension_input(
            stripped_input)
        part_model = PartModel(dimension_inputs[0], dimension_inputs[1],
                               dimension_inputs[2], dimension_inputs[3],
                               dimension_inputs[4], None, dimension_inputs[5])
        part_model.create_l_bracket()
        part_model.add_all_global_variables()
        part_model.l_bracket_equations()
    elif stripped_input == "c":
        part_controller.check_connection()
        part_controller.start_input()
    # prompt user for custom dimensions
    part_controller.main_mod_dims()


if __name__ == "__main__":
    main()
