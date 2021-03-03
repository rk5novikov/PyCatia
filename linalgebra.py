from math import cos, sin, pi, sqrt, atan2
from bisect import bisect_left

vector = (lambda p1, p2: [coo2 - coo1 for coo2, coo1 in zip(p2, p1)])
length = (lambda v: sqrt(sum([i**2 for i in v])))
unit = (lambda v: [coo/length(v) for coo in v])
cross_product = (lambda v1, v2:[v1[1]*v2[2] - v1[2]*v2[1], v1[2]*v2[0] - v1[0]*v2[2], v1[0]*v2[1] - v1[1]*v2[0]])
dot_product = (lambda v1, v2: sum([coo1 * coo2 for coo1, coo2 in zip(v1, v2)]))
summ = (lambda v1, v2: [coo1 + coo2 for coo1, coo2 in zip(v1, v2)])
subtract = (lambda v1, v2: [coo1 - coo2 for coo1, coo2 in zip(v1, v2)])
multiply = (lambda const, v: [const * coo for coo in v])
point_proj_on_axis = (lambda point, p1, p2: summ(p1, [dot_product(vector(p1, point), unit(vector(p1, p2))) * coo for coo in unit(vector(p1, p2))]))
distance_2points = (lambda p1, p2: sqrt(sum([(coo1 - coo2)**2 for coo1, coo2 in zip(p1, p2)])))
distance_point_axis = (lambda point, p1, p2: distance_2points(point, point_proj_on_axis(point, p1, p2)))
point_rotate = (lambda point, p1,p2, teta: summ(p1, vector_rotate(vector(p1, point), vector(p1, p2), teta)))

def vector_rotate(v, axis, teta):
    t = teta * pi / 180
    x, y, z = unit(axis)
    return [(cos(t)+(1-cos(t))*x**2)*v[0] + ((1-cos(t))*x*y-sin(t)*z)*v[1] + ((1-cos(t))*x*z+sin(t)*y)*v[2],
                    ((1-cos(t))*y*x+sin(t)*z)*v[0] + (cos(t)+(1-cos(t))*y**2)*v[1] + ((1-cos(t))*y*z-sin(t)*x)*v[2],
                    ((1-cos(t))*z*x-sin(t)*y)*v[0] + ((1-cos(t))*z*y+sin(t)*x)*v[1] + (cos(t)+(1-cos(t))*z**2)*v[2]]

# OK
def lin_interp(x_, ps):
	if x_ >= ps[-1][0]:
		return ps[-1][1]
	elif x_ <= ps[0][0]:
		return ps[0][1]
	else:
		i = bisect_left([p[0] for p in ps], x_)
		return ps[i-1][1] + (ps[i][1] - ps[i-1][1]) * (x_ - ps[i-1][0]) / (ps[i][0] - ps[i-1][0])

# OK
def read_list(path_to_file):
	res = []
	if not path_to_file:
		return []
	with open(path_to_file) as f:
		for line in f:
			if not line.isspace():
				res.append(float(line.strip()))
	return res

# OK
def read_courbe(path_to_file):
	res = []
	if not path_to_file:
		return []
	with open(path_to_file, 'r') as f:
		for line in f:
			if not line.isspace():
				data = line.split()
				res.append(( float(data[0]), float(data[1]) ))
	return sorted(res, key=lambda p: p[0])

# OK
def read_courbes(path_to_file):
	assert path_to_file, 'Empty path to file given'
	res = {}
	indata = 0
	with open(path_to_file, 'r') as f:
		for line in f:
			data = line.split()
			if data:
				if data[0] == r'*':
					cur_temp = float(data[1])
					res[cur_temp] = []
					indata = 1
				elif indata:
					res[cur_temp].append((float(data[0]), float(data[1])))
	for k in res.keys():
		res[k] = sorted(res[k], key=lambda p: p[0])
	return res

# OK
def check_num(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def chunks(l, n):
    for i in range(0, len(l), n):
        yield l[i:i + n] 