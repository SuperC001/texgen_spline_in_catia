import win32com.client
catia = win32com.client.Dispatch("catia.application")
documents1 = catia.documents
partdocument1 = documents1.add("Part")
part1 = partdocument1.part
bodies1 = part1.bodies
body1 = bodies1.add()
sketches1 = body1.sketches
rf = part1.OriginElements.planeyz
sketch1 = sketches1.add(rf)
factory2D1 = sketch1.openedition()

# 创建样条曲线的控制点
controlpoint2D1 = factory2D1.createcontrolpoint(185.371460, 34.484993)
controlpoint2D2 = factory2D1.createcontrolpoint(77.456879, -47.910454)
controlpoint2D3 = factory2D1.createcontrolpoint(-76.406647, 43.435303)
controlpoint2D4 = factory2D1.createcontrolpoint(-166.204132, 34.484993)

# 创建控制点列表，为创建样条线提供参数
controlpoints = [controlpoint2D1, controlpoint2D2, controlpoint2D3, controlpoint2D4]

spline2D1 = factory2D1.createspline(controlpoints)

sketch1.closeedition()

part1.update()