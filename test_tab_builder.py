import tab_builder
import pytest

eq1 = ['"length" = 2', '"D4@Sketch1" =  "length"/2',
       '"D3@Sketch2" =  "length"/2', '"D2@Sketch2" =  "arc rad"- "length"']
var1 = '"length" = 2'
var2 = '"width" = 1'
eq2 = ['"D1@Boss-Extrude1" =  "width"', '"width" = 1']
var3 = '"bore size" = 1'
eq3 = ['"D5@Sketch1" =  "bore size"', '"bore size" = 1']
var4 = '"bore placement" = 1'
eq4 = ['"D3@Sketch1" =  "bore placement"',
       '"bore placement" = 1']
var5 = '"height" = 1'
eq5 = ['"D2@Sketch1" =  "height"', '"height" = 1']
var6 = '"arc rad" = 1'
eq6 = ['"D1@Sketch2" =  "arc rad"', '"arc rad" = 1']

test_cases = [(eq1, var1), (eq2, var2), (eq3, var3),
              (eq4, var4), (eq5, var5), (eq6, var6)]


@ pytest.mark.parametrize("equation_list, var", test_cases)
def test_global_var_finder(equation_list, var):
    part_controller = tab_builder.PartController()
    input_var = var.split("=")
    input_var = input_var[0].strip()
    global_var = part_controller.global_var_finder(equation_list, input_var)
    assert global_var == var
