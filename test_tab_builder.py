import tab_builder


# check modify equation list
# check global variable finder
# check equation_list_maker

equation_list = ['"length" = 0.025399986284007407', '"width" = 0.006349996571001852', '"bore size" = 0.012699993142003704', '"bore placement" = 0.03809997942601111', '"height" = 0.050799972568014815', '"arc rad" = 0.03174998285500926', '"D1@Sketch1" =  "length"', '"D2@Sketch1" =  "height"', '"D3@Sketch1" =  "bore placement"', '"D5@Sketch1" =  "bore size"', '"D4@Sketch1" =  "length"/2', '"D1@Boss-Extrude1" =  "width"', '"D3@Sketch2" =  "length"/2', '"D1@Sketch2" =  "arc rad"', '"D2@Sketch2" =  "arc rad"- "length"']


@pytest.mark.timeout(test_global_vars)
def test_modify_equation_list():
    length = 1
    bore_size = 0.25
    bore_placement = 1.5
    width = 0.1
    height = 2
    arc_rad = 1.25
    L_bracket_length = None

    part_model = PartModel(length,bore_size,bore_placement,width, height, arc_rad, L_bracket_length)
    part_model.create_simple_tab()
    part_model.add_all_global_variables
    assert data == test_years1 