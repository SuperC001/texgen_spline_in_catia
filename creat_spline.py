import win32com.client
catia = win32com.client.Dispatch("catia.application")
documents1 = catia.documents
partdocument1 = documents1.add("Part")
part1 = partdocument1.part
bodies1 = part1.bodies
body1 = bodies1.add()
sketches1 = body1.sketches

# 将yz设置为参考平面
rf = part1.OriginElements.planeyz
sketch1 = sketches1.add(rf)
factory2D1 = sketch1.openedition()

# 批量实例化控制点,并将其添加到controlpoint2D列表中
controlpoint2D = []

# 列表中为控制点参数
controlpoints_data = [(185.371460, 34.484993),(77.456879, -47.910454),(-76.406647, 43.435303),(-166.204132, 34.484993)]
# 一条曲线中控制点个数
control_point_num = len(controlpoints_data)
for i in range(control_point_num):
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