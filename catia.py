'''
MasseNacelle python module.
Python / CATIA interaction module for automation of basic operations with *.CATPart models.
For all operations python wrappers of CATIA COM interface functions are used. 
'''

__author__ = 'Vadim NOVIKOV / Smartec JSC'
__status__ = 'TEST'
__date__ = '15 janvier 2018'


from itertools import chain
import win32api
import win32com.client.dynamic
import pythoncom
from pywintypes import com_error

## Dictionary which store rgb codes of colors
rgb_colors = {
		'black'  : [  0,  0,  0], 
		'white'  : [255,255,255], 
		'red'    : [255,  0,  0], 
		'lime'   : [  0,255,  0], 
		'blue'   : [  0,  0,255], 
		'yellow' : [255,255,  0], 
		'cyan'   : [0  ,255,255], 
		'magenta': [255,0  ,255], 
		'silver' : [192,192,192], 
		'gray'   : [128,128,128], 
		'maroon' : [128,  0,  0], 
		'olive'  : [128,128,  0], 
		'green'  : [  0,128,  0], 
		'purple' : [128,  0,128], 
		'teal'   : [  0,128,128], 
		'navy'   : [  0,  0,128]
		}

# Constants for catia: constraint type
catia_constant_reference = 0
catia_constant_distance = 1
catia_constant_on = 2
catia_constant_concentricity = 3
catia_constant_tangency = 4
catia_constant_length = 5
catia_constant_angle = 6
catia_constant_planar_angle = 7
catia_constant_parallelism = 8
catia_constant_axis_parallelism = 9
catia_constant_horizontality = 10
catia_constant_perpendicularity = 11
catia_constant_axis_perpendicularity = 12
catia_constant_verticality = 13
catia_constant_radius = 14
catia_constant_symmetry = 15
catia_constant_midpoint = 16
catia_constant_equidistance = 17
catia_constant_major_radius = 18
catia_constant_minor_radius = 19
catia_constant_surf_contact = 20
catia_constant_lin_contact = 21
catia_constant_ponc_contact = 22
catia_constant_chamfer = 23
catia_constant_chamfer_perpend = 24
catia_constant_annul_contact = 25
catia_constant_cylinder_radius = 26
catia_constant_st_continuity = 27
catia_constant_st_distance = 28
catia_constant_sd_continuity = 29
catia_constant_sd_shape = 30
# Constants for catia: constraint mode
catia_constant_driving = 0
catia_constant_driven = 1


## Class for storing catia application COM-object and general actions with documents."""
class CATIA():

	def __init__(self, visible=True):
		pythoncom.CoInitialize()
		self.app = win32com.client.Dispatch('CATIA.Application')
		self.app.Visible = visible

	def __init_part_objects(self):
		self.part = self.app.ActiveDocument.Part
		self.hybrid_bodies = self.part.HybridBodies
		self.shape_factory = self.part.HybridShapeFactory
		self.parameteres = self.part.Parameters

	def open(self, file_path):
		self.app.Documents.Open(file_path)
		self.__init_part_objects()

	def new(self, part_name):
		self.app.Documents.Add(part_name)
		self.__init_part_objects()

	def save_as(self, file_path):
		self.app.ActiveDocument.SaveAs(file_path)

	def save(self):
		self.app.ActiveDocument.Save()

	def close(self):
		self.app.ActiveDocument.Close()

	def quit(self):
		self.app.Quit()

## Running of CATIA application.
def start_catia(catia_path, visible=True):
	global cur_catia
	cur_catia = CATIA(visible)
	cur_catia.open(catia_path)

## Saving of active open document.
def catia_active_document_save():
	global cur_catia
	cur_catia.save()

## Hide specified geometry element.
def hide(*elements):
	to_hide = cur_catia.app.ActiveDocument.Selection
	to_hide.Clear()
	for geometry in elements:
		to_hide.Add(geometry)
	to_hide.VisProperties.SetShow(1)
	to_hide.Clear()
	cur_catia.part.Update()

## Creating of empty geometrical set.
def create_hybrid_body(name):
	hybrid_body = cur_catia.hybrid_bodies.Add()
	hybrid_body.Name = name
	cur_catia.part.Update()
	cur_catia.current_hybrid_body = hybrid_body
	return hybrid_body

## Indicates geometrical set selecting by name as active set for adding any further created geometry objects.
def activate_hybrid_body(hybrid_body_name):
	cur_catia.current_hybrid_body = cur_catia.hybrid_bodies.Item(hybrid_body_name)

## Creates a point by coordinates and appends result to active geometrical set.
def create_point_coord(name, coords, ref_axis_system=False):
	point = cur_catia.shape_factory.AddNewPointCoord(*coords)
	if ref_axis_system:
		point.RefAxisSystem = get_reference(ref_axis_system)
	point.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(point)
	cur_catia.part.Update()
	return point

## Creates a new datum of point within the current body and appends result to active geometrical set.
def create_point_datum(name, point):
	point_ref = get_reference(point)
	datum = cur_catia.shape_factory.AddNewPointDatum(point_ref)
	cur_catia.current_hybrid_body.AppendHybridShape(datum)
	datum.Name = name
	cur_catia.part.Update()
	return datum

## Creates an axis system.
def create_axis_system(name, origin_point, axis_vector_1, axis_vector_2, axis_vector_3):
	axis_system = cur_catia.part.AxisSystems.Add()
	axis_system.OriginPoint = origin_point
	axis_system.PutXAxis(axis_vector_1)
	axis_system.PutYAxis(axis_vector_2)
	axis_system.PutZAxis(axis_vector_3)
	axis_system.IsCurrent = True
	axis_system.Name = name
	return axis_system

## Creates a new offset trough point plane within the current body and appends result to active geometrical set.
def create_plane_offset_pt(name, reference_plane, point):
	plane = cur_catia.shape_factory.AddNewPlaneOffsetPt(reference_plane, point)
	plane.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(plane)
	cur_catia.part.Update()
	return plane

## Creates a new direction specified by an element within the current body and appends result to active geometrical set.
def create_direction(reference_geometry):
	return cur_catia.shape_factory.AddNewDirection(reference_geometry)

## Creates a new point-direction line within the current body and appends result to active geometrical set.
def create_line_pt_dir(name, point, direction, limit_1, limit_2, orientation):
	line = cur_catia.shape_factory.AddNewLinePtDir(point, direction, limit_1, limit_2, bool(orientation))
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a new point-direction line within the current body and appends result to active geometrical set.
def create_line_dir_on_support(name, point, direction, plane, limit_1, limit_2, orientation):
	line = cur_catia.shape_factory.AddNewLinePtDirOnSupport(point, direction, plane, limit_1, limit_2, bool(orientation))
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a new bitangent line within the current body and appends result to active geometrical set.
def create_line_bitang(name, curve_1, curve_2, support=None):
	line = cur_catia.shape_factory.AddNewLineBiTangent(curve_1, curve_2, support)
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a new Split within the current body and appends result to active geometrical set.
def create_hybrid_split(name, surface, splitting_geometry, orientation):
	split = cur_catia.shape_factory.AddNewHybridSplit(surface, splitting_geometry, orientation)
	split.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(split)
	cur_catia.part.Update()
	return split

## Creates a new Translate within the current body and appends result to active geometrical set.
def create_translate(name, geometry, direction, distance):
	translate = cur_catia.shape_factory.AddNewEmptyTranslate()
	translate.Name = name
	translate.ElemToTranslate = geometry
	translate.VectorType = 0
	translate.Direction = direction
	translate.DistanceValue = distance
	translate.VolumeResult = False
	cur_catia.current_hybrid_body.AppendHybridShape(translate)
	cur_catia.part.Update()
	return translate

## Creates a new Intersection within the current body and appends result to active geometrical set.
def create_intersection(name, geometry_1, geometry_2):
	intersection = cur_catia.shape_factory.AddNewIntersection(geometry_1, geometry_2)
	intersection.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(intersection)
	cur_catia.part.Update()
	return intersection

## Creates a new offset trough point plane within the current body and appends result to active geometrical set.
def create_plane_offset(name, reference_plane, offset, orientation):
	plane = cur_catia.shape_factory.AddNewPlaneOffset(reference_plane, offset, bool(orientation))
	plane.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(plane)
	cur_catia.part.Update()
	return plane

## Creates a new angle line within the current body and appends result to active geometrical set.
def create_line_angle(name, reference_line, plane, point, limit_1, limit_2, angle):
	line = cur_catia.shape_factory.AddNewLineAngle(reference_line, plane, point, False, limit_1, limit_2, angle, False)
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a Boundary within the current body and appends result to active geometrical set.
def create_boundary_of_surfaces(name, reference_surface):
	boundary = cur_catia.shape_factory.AddNewBoundaryOfSurface(reference_surface)
	cur_catia.current_hybrid_body.AppendHybridShape(boundary)
	boundary.Name = name
	cur_catia.part.Update()
	return boundary


## Creates a new Extremum within the current body and appends result to active geometrical set.
def create_extremum(name, geometry, direction, orientation, direction_2=None, orientation_2=None, direction_3=None, orientation_3=None):
	extremum = cur_catia.shape_factory.AddNewExtremum(geometry, direction, orientation)
	if direction_2:
		extremum.Direction2 = direction_2
		extremum.ExtremumType2 = orientation_2
	if direction_3:
		extremum.Direction3 = direction_3
		extremum.ExtremumType3 = orientation_3
	cur_catia.current_hybrid_body.AppendHybridShape(extremum)
	extremum.Name = name
	cur_catia.part.Update()
	return extremum

## Creates a new point-point line with extensions within the current body and appends result to active geometrical set.
def create_line_pt_pt(name, point_1, point_2):
	line = cur_catia.shape_factory.AddNewLinePtPt(point_1, point_2)
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a new point-point line with support within the current body and appends result to active geometrical set.
def create_line_pt_pt_on_support(name, point_1, point_2, plane):
	line = cur_catia.shape_factory.AddNewLinePtPtOnSupport(point_1, point_2, plane)
	line.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(line)
	cur_catia.part.Update()
	return line

## Creates a new normal plane within the current body and appends result to active geometrical set.
def create_plane_normal(name, line, point):
	plane = cur_catia.shape_factory.AddNewPlaneNormal(line, point)
	plane.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(plane)
	cur_catia.part.Update()
	return plane

## Creates a new whole circle defined by its center, a passing point within the current body and appends result to active geometrical set.
def create_circle_ctr_pt(name, center_point, radius_point, plane):
	circle = cur_catia.shape_factory.AddNewCircleCtrPt(center_point, radius_point, plane, True)
	circle.Name = name
	circle.SetLimitation(1)
	cur_catia.current_hybrid_body.AppendHybridShape(circle)
	cur_catia.part.Update()
	return circle

## Creates a new extrude within the current body and appends result to active geometrical set.
def create_extrude(name, line, limit_1, limit_2, direction):
	extrude = cur_catia.shape_factory.AddNewExtrude(line, limit_1, limit_2, direction)
	extrude.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(extrude)
	cur_catia.part.Update()
	return extrude

## Creates a new Join within the current body and appends result to active geometrical set.
def create_join(name, set_connex, *elements):
	join = cur_catia.shape_factory.AddNewJoin(elements[0], elements[1])
	join.Name = name
	if len(elements) > 2:
		for i in range(2, len(elements)):
			join.AddElement(elements[i])
	join.SetConnex(set_connex)
	join.SetManifold(0)
	join.SetSimplify(0)
	join.SetSuppressMode(0)
	join.SetDeviation(0.001)
	join.SetAngularToleranceMode(0)
	join.SetAngularTolerance(0.5)
	join.SetFederationPropagation(0)
	cur_catia.current_hybrid_body.AppendHybridShape(join)
	cur_catia.part.Update()
	return join

## Creates a new angle plane within the current body and appends result to active geometrical set.
def create_plane_angle(name, reference_plane, reference_line, angle, orientation):
	plane = cur_catia.shape_factory.AddNewPlaneAngle(reference_plane, reference_line, angle, bool(orientation))
	plane.ProjectionMode = False
	plane.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(plane)
	cur_catia.part.Update()
	return plane

## Creates a new empty Rotate within the current body and appends result to active geometrical set.
def create_empty_rotate(name, geometry, line_axis, angle):
	empty_rotate = cur_catia.shape_factory.AddNewEmptyRotate()
	empty_rotate.ElemToRotate = get_reference(geometry)
	empty_rotate.VolumeResult = False
	empty_rotate.RotationType = 0
	empty_rotate.Axis = get_reference(line_axis)
	empty_rotate.AngleValue = angle
	empty_rotate.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(empty_rotate)
	cur_catia.part.Update()
	return empty_rotate

## Creates boolean parameter in selected parameter set or root parameter set.
def create_boolean(name, value=False, param_set=None):
	if param_set:
		param_set_ref = cur_catia.parameteres.RootParameterSet.ParameterSets.Item(param_set)
		boolean = param_set_ref.DirectParameters.CreateBoolean(name, value)
	else:
		boolean = cur_catia.parameteres.CreateBoolean(name, value)
	return boolean

## Creates string parameter in selected parameter set or root parameter set.
def create_string(name, value='', param_set=None):
	if param_set:
		param_set_ref = cur_catia.parameteres.RootParameterSet.ParameterSets.Item(param_set)
		string = param_set_ref.DirectParameters.CreateString('', value)
	else:
		string = cur_catia.parameteres.CreateString('', value)
	string.Rename(name)
	return string

## Creates real parameter in selected parameter set or root parameter set.
def create_real(name, value=0.0, param_set=None):
	if param_set:
		param_set_ref = cur_catia.parameteres.RootParameterSet.ParameterSets.Item(param_set)
		real = param_set_ref.DirectParameters.CreateReal('', value)
	else:
		real = cur_catia.parameteres.CreateReal('', value)
	real.Rename(name)
	return real

## Creates dimension parameter in selected parameter set or root parameter set.
def create_dimension(name, dimension_type, value, param_set=None):
	if param_set:
		param_set_ref = cur_catia.parameteres.RootParameterSet.ParameterSets.Item(param_set)
		dimension = param_set_ref.DirectParameters.CreateDimension('', dimension_type, value)
	else:
		dimension = cur_catia.parameteres.CreateDimension('', dimension_type, value)
	dimension.Rename(name)
	dimension.Value = value
	return dimension

## Creates a new point on a curve from a ratio of distance to an extremity within the current body and appends result to active geometrical set.
def create_point_on_curve_from_percent(name, line, perc, orientation):
	point = cur_catia.shape_factory.AddNewPointOnCurveFromPercent(line, perc, orientation)
	point.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(point)
	cur_catia.part.Update()
	return point

## Creates a new point on a curve extremum appropriated a max coord selected (max x, min x, max y etc.) 
# within the current body and appends result to active geometrical set.
def create_point_on_curve_extr(name, line, dir_coord):
	point_1 = cur_catia.shape_factory.AddNewPointOnCurveFromPercent(line, 1.0, True)
	point_2 = cur_catia.shape_factory.AddNewPointOnCurveFromPercent(line, 1.0, False)
	cur_catia.current_hybrid_body.AppendHybridShape(point_1)
	cur_catia.current_hybrid_body.AppendHybridShape(point_2)
	cur_catia.part.Update()
	# getting coordinates
	point_1_X = create_dimension('point_1_X', 'LENGTH', 0.0)
	point_2_X = create_dimension('point_2_X', 'LENGTH', 0.0)
	point_1_Y = create_dimension('point_1_Y', 'LENGTH', 0.0)
	point_2_Y = create_dimension('point_2_Y', 'LENGTH', 0.0)
	point_1_Z = create_dimension('point_1_Z', 'LENGTH', 0.0)
	point_2_Z = create_dimension('point_2_Z', 'LENGTH', 0.0)
	# formulas . . . 
	point_1_X_formula = create_formula('point_1_X', '', point_1_X, '{0}\\{1}.coord(1)'.format(cur_catia.current_hybrid_body.Name, point_1.Name))
	point_2_X_formula = create_formula('point_2_X', '', point_2_X, '{0}\\{1}.coord(1)'.format(cur_catia.current_hybrid_body.Name, point_2.Name))
	point_1_Y_formula = create_formula('point_1_Y', '', point_1_Y, '{0}\\{1}.coord(2)'.format(cur_catia.current_hybrid_body.Name, point_1.Name))
	point_2_Y_formula = create_formula('point_2_Y', '', point_2_Y, '{0}\\{1}.coord(2)'.format(cur_catia.current_hybrid_body.Name, point_2.Name))
	point_1_Z_formula = create_formula('point_1_Z', '', point_1_Z, '{0}\\{1}.coord(3)'.format(cur_catia.current_hybrid_body.Name, point_1.Name))
	point_2_Z_formula = create_formula('point_2_Z', '', point_2_Z, '{0}\\{1}.coord(3)'.format(cur_catia.current_hybrid_body.Name, point_2.Name))
	# temporary elements to delete from tree
	temp_elements = [point_1_X_formula, point_1_Y_formula, point_1_Z_formula, 
					 point_2_X_formula, point_2_Y_formula, point_2_Z_formula, 
					 point_1_X        , point_1_Y        , point_1_Z        , 
					 point_2_X        , point_2_Y        , point_2_Z        ]
	# Selecting poins to delete / return
	if dir_coord == '+X':
		if point_1_X.Value > point_2_X.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	elif dir_coord == '-X':
		if point_1_X.Value < point_2_X.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	elif dir_coord == '+Y':
		if point_1_Y.Value > point_2_Y.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	elif dir_coord == '-Y':
		if point_1_Y.Value < point_2_Y.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	elif dir_coord == '+Z':
		if point_1_Z.Value > point_2_Z.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	elif dir_coord == '-Z':
		if point_1_Z.Value < point_2_Z.Value:
			point_to_del = point_2
			point = point_1
		else:
			point_to_del = point_1
			point = point_2
	point.Name = name
	sel = cur_catia.app.ActiveDocument.Selection
	sel.Clear()
	for element in temp_elements:
		sel.Add(element)
	sel.Delete()
	sel.Clear()
	cur_catia.shape_factory.DeleteObjectForDatum(point_to_del)
	cur_catia.part.Update()
	return point


## Creates a new circle tangent to 2 curves and passing through one point within the current body and appends result to active geometrical set.
def create_circle_bitang_point(name, line_1, line_2, point, support, orientation_1, orientation_2):
	try:
		print(0)
		circle = cur_catia.shape_factory.AddNewCircleBitangentPoint(line_1, line_2, point, support, orientation_1, orientation_2)
		cur_catia.current_hybrid_body.AppendHybridShape(circle)
		cur_catia.part.Update()
	except:
		try:
			print(1)
			cur_catia.shape_factory.DeleteObjectForDatum(circle)
			circle = cur_catia.shape_factory.AddNewCircleBitangentPoint(line_1, line_2, point, support, -orientation_1, orientation_2)
			cur_catia.current_hybrid_body.AppendHybridShape(circle)
			cur_catia.part.Update()
		except:
			try:
				print(2)
				cur_catia.shape_factory.DeleteObjectForDatum(circle)
				circle = cur_catia.shape_factory.AddNewCircleBitangentPoint(line_1, line_2, point, support, orientation_1, -orientation_2)
				cur_catia.current_hybrid_body.AppendHybridShape(circle)
				cur_catia.part.Update()
			except:
				try:
					print(3)
					cur_catia.shape_factory.DeleteObjectForDatum(circle)
					circle = cur_catia.shape_factory.AddNewCircleBitangentPoint(line_1, line_2, point, support, -orientation_1, -orientation_2)
					cur_catia.current_hybrid_body.AppendHybridShape(circle)
					cur_catia.part.Update()
				except com_error as e:
					raise e
	circle.SetLimitation(1)
	circle.Name = name
	return circle

## Creates a new revolution within the current body and appends result to active geometrical set.
def create_revol(name, geometry, angle_1, angle_2, line_axis):
	revol = cur_catia.shape_factory.AddNewRevol(geometry, angle_1, angle_2, line_axis)
	revol.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(revol)
	cur_catia.part.Update()
	return revol

## Creates a formula relation and adds it to the part's collection of relations.
def create_formula(name, comment, dimension, text_definition):
	formula = cur_catia.part.Relations.CreateFormula(name, comment, dimension, text_definition)
	cur_catia.part.Update()
	return formula

# Sketching
## Creates sketch in current geometrical set.
def create_sketch(name, reference_plane, origin, axis_h, axis_v):
	sketch = cur_catia.current_hybrid_body.HybridSketches.Add(reference_plane)
	sketch.Name = name
	sketch.SetAbsoluteAxisData(list(chain(origin, axis_h, axis_v)))
	cur_catia.part.Update()
	return sketch

## Open sketch in edit mode.
def open_sketch(sketch):
	cur_catia.factory2D = sketch.OpenEdition()
	cur_catia.current_sketch = sketch

## Creates point in opened sketch.
def sketch_create_point(name, coords):
	point = cur_catia.factory2D.CreatePoint(*coords)
	point.Name = name
	cur_catia.part.Update()
	return point

## Creates line in opened sketch through 2 point and its coordinates.
def sketch_create_line(name, p1_coords, p2_coords, p1, p2):
	line = cur_catia.factory2D.CreateLine(*(p1_coords + p2_coords))
	line.Name = name
	line.StartPoint = p1
	line.EndPoint = p2
	cur_catia.part.Update()
	return line

## Creates a new constraint applying to two geometric elements and adds it to the Constraints collection.
def sketch_create_constraint(name, con_type, geometry_1, geometry_2, val=None, ang_sector=None, mode=0):
	ref_1 = get_reference(geometry_1)
	ref_2 = get_reference(geometry_2)
	con = cur_catia.current_sketch.Constraints.AddBiEltCst(con_type, ref_1, ref_2)
	if val:
		con.Dimension.Value = val
	if ang_sector:
		con.AngleSector = ang_sector
	if name:
		con.Name = name
	con.Mode = mode
	return con

## Creates and returns the projection of an object on the opened sketch.
def sketch_create_projection(name, geometry):
	projection = cur_catia.current_sketch.factory2D.CreateProjection(geometry)
	projection.Name = name
	cur_catia.part.Update()
	return projection

## Returns reference to constraint object in sketch by constraint name.
def sketch_get_dimention(sketch, constraint_name):
	return sketch.Constraints.Item(constraint_name).Dimension

## Creates and returns the possible intersections of an object with the sketch.
def sketch_create_intersection(geometry):
	intersection = cur_catia.current_sketch.factory2D.CreateIntersections(geometry)
	cur_catia.part.Update()
	return intersection

## Close sketch.
def close_sketch():
	cur_catia.current_sketch.CloseEdition()
	cur_catia.part.Update()

## Creates a new CurvePar within the current body.
def create_curve_par(name, curve, support, distance, invert_direction, geodesic=False):
	curve_par = cur_catia.shape_factory.AddNewCurvePar(curve, support, distance, invert_direction, geodesic)
	curve_par.SmoothingType = 0
	curve_par.Name = name
	cur_catia.current_hybrid_body.AppendHybridShape(curve_par)
	cur_catia.part.Update()
	return curve_par

## Creates a new CurvePar by defining offset from curve in support within the current body.
def create_curve_par_dir_safe(name, curve, support, distance, dir_coord, geodesic=False):
	curve_par_1 = cur_catia.shape_factory.AddNewCurvePar(curve, support, distance, True , geodesic)
	curve_par_2 = cur_catia.shape_factory.AddNewCurvePar(curve, support, distance, False, geodesic)
	# Creating 2 versions of curve by offset: straightforward direct and inverse
	cur_catia.current_hybrid_body.AppendHybridShape(curve_par_1)
	cur_catia.current_hybrid_body.AppendHybridShape(curve_par_2)
	# Creating midpoints on curves to determine by them which curve satisfies direction needed (dir_coord)
	mid_point_1 = cur_catia.shape_factory.AddNewPointOnCurveFromPercent(curve_par_1, 0.5, True)
	mid_point_2 = cur_catia.shape_factory.AddNewPointOnCurveFromPercent(curve_par_2, 0.5, True)
	# Updating . . . 
	cur_catia.current_hybrid_body.AppendHybridShape(mid_point_1)
	cur_catia.current_hybrid_body.AppendHybridShape(mid_point_2)
	cur_catia.part.Update()
	cur_catia.part.Update()
	# getting coordinates
	mid_point_1_X = create_dimension('mid_point_1_X', 'LENGTH', 0.0)
	mid_point_2_X = create_dimension('mid_point_2_X', 'LENGTH', 0.0)
	mid_point_1_Y = create_dimension('mid_point_1_Y', 'LENGTH', 0.0)
	mid_point_2_Y = create_dimension('mid_point_2_Y', 'LENGTH', 0.0)
	mid_point_1_Z = create_dimension('mid_point_1_Z', 'LENGTH', 0.0)
	mid_point_2_Z = create_dimension('mid_point_2_Z', 'LENGTH', 0.0)
	# formulas . . . 
	mid_point_1_X_formula = create_formula('mid_point_1_X', '', mid_point_1_X, '{0}\\{1}.coord(1)'.format(cur_catia.current_hybrid_body.Name, mid_point_1.Name))
	mid_point_2_X_formula = create_formula('mid_point_2_X', '', mid_point_2_X, '{0}\\{1}.coord(1)'.format(cur_catia.current_hybrid_body.Name, mid_point_2.Name))
	mid_point_1_Y_formula = create_formula('mid_point_1_Y', '', mid_point_1_Y, '{0}\\{1}.coord(2)'.format(cur_catia.current_hybrid_body.Name, mid_point_1.Name))
	mid_point_2_Y_formula = create_formula('mid_point_2_Y', '', mid_point_2_Y, '{0}\\{1}.coord(2)'.format(cur_catia.current_hybrid_body.Name, mid_point_2.Name))
	mid_point_1_Z_formula = create_formula('mid_point_1_Z', '', mid_point_1_Z, '{0}\\{1}.coord(3)'.format(cur_catia.current_hybrid_body.Name, mid_point_1.Name))
	mid_point_2_Z_formula = create_formula('mid_point_2_Z', '', mid_point_2_Z, '{0}\\{1}.coord(3)'.format(cur_catia.current_hybrid_body.Name, mid_point_2.Name))
	# temporary elements to delete from tree
	temp_elements = [mid_point_1_X_formula, mid_point_1_Y_formula, mid_point_1_Z_formula, 
					 mid_point_2_X_formula, mid_point_2_Y_formula, mid_point_2_Z_formula, 
					 mid_point_1_X        , mid_point_1_Y        , mid_point_1_Z        , 
					 mid_point_2_X        , mid_point_2_Y        , mid_point_2_Z        ]
	# Selecting poins to delete / return
	if dir_coord == '+X':
		if mid_point_1_X.Value > mid_point_2_X.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	elif dir_coord == '-X':
		if mid_point_1_X.Value < mid_point_2_X.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	elif dir_coord == '+Y':
		if mid_point_1_Y.Value > mid_point_2_Y.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	elif dir_coord == '-Y':
		if mid_point_1_Y.Value < mid_point_2_Y.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	elif dir_coord == '+Z':
		if mid_point_1_Z.Value > mid_point_2_Z.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	elif dir_coord == '-Z':
		if mid_point_1_Z.Value < mid_point_2_Z.Value:
			curve_to_del = curve_par_2
			curve_par = curve_par_1
		else:
			curve_to_del = curve_par_1
			curve_par = curve_par_2
	# Define properties needed
	curve_par.Name = name
	curve_par.SmoothingType = 0
	# Deleting temporary elements
	sel = cur_catia.app.ActiveDocument.Selection
	sel.Clear()
	for element in temp_elements:
		sel.Add(element)
	sel.Delete()
	sel.Clear()
	cur_catia.shape_factory.DeleteObjectForDatum(mid_point_1)
	cur_catia.shape_factory.DeleteObjectForDatum(mid_point_2)
	cur_catia.shape_factory.DeleteObjectForDatum(curve_to_del)
	# Updating . . . 
	cur_catia.part.Update()
	return curve_par

## Creates a reference from a operator.
def get_reference(obj):
	return cur_catia.part.CreateReferenceFromObject(obj)

## Returns a reference to item in specified geometrical set. Default - active geomtrical set.
def get_item(item_name, hybrid_body_name=''):
	if not hybrid_body_name:
		return cur_catia.current_hybrid_body.HybridShapes.Item(item_name)
	else:
		hb = get_hybrid_body(hybrid_body_name)
		return hb.HybridShapes.Item(item_name)

## Returns a reference to parameter specified by name.
def get_parametre_ref(parametre_name):
	return get_reference(cur_catia.parameteres.Item(parametre_name))

## Returns value of the parameter specified by name.
def get_parametre_val(parametre_name):
	return cur_catia.parameteres.Item(parametre_name).Value

## Returns reference to AxisSystem specified by its name.
def get_axis_system(name):
	return cur_catia.part.AxisSystems.Item(name)

## Returns reference to one of origin plane (xy, yz or zx).
def get_origin_plane(plane):
	if plane == 'xy':
		return cur_catia.part.OriginElements.PlaneXY
	elif plane == 'yz':
		return cur_catia.part.OriginElements.PlaneYZ
	elif plane == 'zx':
		return cur_catia.part.OriginElements.PlaneZX

## Returns reference to geometrical set by its name.
def get_hybrid_body(hybrid_body_name):
	return cur_catia.hybrid_bodies.Item(hybrid_body_name)

## Check if parameter exists in tree.
def parametre_exists(parametre_name):
	try:
		get_parametre_ref(parametre_name)
		return True
	except com_error:
		return False

## Set color of geometrical object.
def set_color(geometry, color, heritance=0):
	args = rgb_colors[color].copy()
	args.append(heritance)
	sel = cur_catia.app.ActiveDocument.Selection
	sel.Clear()
	sel.Add(geometry)
	sel.VisProperties.SetRealColor(*args)
	sel.Clear()
	cur_catia.part.Update()

## Creates parametres set.
def create_parametere_set(name):
	param_set = cur_catia.parameteres.CreateSetOfParameters(cur_catia.parameteres.RootParameterSet)
	param_sets = cur_catia.parameteres.RootParameterSet.ParameterSets
	param_set_ref = get_reference(param_sets.Item(param_sets.Count))
	cur_catia.shape_factory.ChangeFeatureName(param_set_ref, name)
	return param_set_ref


