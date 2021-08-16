import win32com.client
from array import *
catia = win32com.client.Dispatch("catia.application")
partdocument1 = catia.documents.add("Part")
part1 = partdocument1.part
hybridbodies1 = part1.hybridbodies
hybridbody1 = hybridbodies1.add()
part1.update()


hybridShapeFactory1 = part1.HybridShapeFactory
hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, 0)

hybridshapepointcoord_list = []
controlpoints_data = [(152.550690, 40.539612,0),(56.189056, -47.647209,0),(-73.518417, 52.912094,0),(-158.852295, -55.544548,0)]
control_point_num = len(controlpoints_data)
for i in range(control_point_num):
    hybridshapepointcoord_list.append(hybridShapeFactory1.addnewpointcoord(controlpoints_data[i][0],controlpoints_data[i][1],controlpoints_data[i][2]))

part1.update()

hybridShapeSpline1 = hybridShapeFactory1.AddNewSpline()
hybridShapeSpline1.SetSplineType(0)
hybridShapeSpline1.SetClosing(0)

point_list = []
for point in hybridshapepointcoord_list:
    # reference_list.append(part1.CreateReferenceFromObject(point))
    reference = part1.CreateReferenceFromObject(point)
    point_list.append(reference)
    hybridShapeSpline1.AddPointWithConstraintExplicit(reference, None, -1, 1, None, 0)

hybridbody1.AppendHybridShape(hybridShapeSpline1) # 添加样条曲线到几何图形集中

""" hybridShapeSpline1.RemoveCurvatureRadiusDirection(4) """
""" print(hybridShapeSpline1.GetNbControlPoint()) """

part1.update()