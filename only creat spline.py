import win32com.client
catia = win32com.client.Dispatch("catia.application")
partdocument1 = catia.documents.add("Part")
part1 = partdocument1.part
bodies1 = part1.bodies
body1 = bodies1.item("零件几何体")


# 将yz设置为参考平面
rf = part1.OriginElements.planeyz
sketch1 = body1.sketches.add(rf)
# 设置sketch1的主坐标
arrayOfVariantOfDouble1 = [
    0.000000,
    0.000000,
    0.000000,
    0.000000,
    1.000000,
    0.000000,
    0.000000,
    0.000000,
    1.000000
]
sketch1.SetAbsoluteAxisData(arrayOfVariantOfDouble1)

factory2D1 = sketch1.openedition()
geometricElements1 = sketch1.GeometricElements
axis2D1 = geometricElements1.Item("绝对轴")
line2D1 = axis2D1.GetItem("横向")
line2D2 = axis2D1.GetItem("纵向")

# 批量实例化控制点,并将其添加到controlpoint2D列表中
controlpoint2D = []
# 列表中为控制点参数
controlpoints_data = [(152.550690, 40.539612),(56.189056, -47.647209),(-73.518417, 52.912094),(-158.852295, -55.544548)]
# 一条曲线中控制点个数
control_point_num = len(controlpoints_data)
for i in range(control_point_num-1,-1,-1):
    controlpoint2D.append(factory2D1.createcontrolpoint(controlpoints_data[i][0],controlpoints_data[i][1]))
#使用控制点列表作为参数实例化样条曲线
spline2D1 = factory2D1.createspline(controlpoint2D)

#获得控制点个数
""" num = spline2D1.getnumberofcontrolpoints()
print(num) """

#获得曲线结束点
""" endpoint = spline2D1.endpoint
print(endpoint) """

#判断曲线是否有周期性
# print(spline2D1.isperiodic())

#获取控制点的切线方向
""" tangent = spline2D1.gettangent(1,[0,1])
print(tangent) """

#获取控制点的曲率
""" cur = spline2D1.getcurvature(0,[0,0,0])
print(cur) """

#获取曲线长度
""" length = spline2D1.getlengthatparam(1,4)
print(length) """

#获取连续性
""" conti = spline2D1.continuity
print(conti) """

sketch1.closeedition()

part1.update()