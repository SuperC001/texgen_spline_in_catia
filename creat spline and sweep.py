import win32com.client


# 创建样条曲线并扫掠
def create_spline_and_sweep(controlpoint_list,ellipse_x,ellipse_y):
    body1 = part1.bodies.add()
    body1.name = "零件1"
    sketches1 = body1.sketches

    # 将yz设置为参考平面
    reference1 = part1.OriginElements.planeyz
    sketch1 = sketches1.add(reference1)
    # 设置sketch1的主坐标
    arrayOfVariantOfDouble1 = [0,0,0,0,1,0,0,0,1]
    sketch1.SetAbsoluteAxisData(arrayOfVariantOfDouble1)

    factory2D1 = sketch1.openedition()

    # 批量实例化控制点,并将其添加到controlpoint2D列表中
    arrayOfObject1 = []
    # 一条曲线中控制点个数
    control_point_num = len(controlpoint_list)
    for i in range(control_point_num-1,-1,-1):
        arrayOfObject1.append(factory2D1.createcontrolpoint(controlpoint_list[i][0],controlpoint_list[i][1]))

    #使用控制点列表作为参数实例化样条曲线
    spline2D1 = factory2D1.createspline(arrayOfObject1)
    sketch1.closeedition()
    part1.update()

    #####  沿以上创建的曲线进行包络体扫掠

    # 在起始控制点创建草图平面
    hybridShapeFactory1 = part1.HybridShapeFactory
    reference2 = part1.CreateReferenceFromObject(sketch1)
    reference3 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;3);None:(Limits1:();Limits2:();-1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)
    hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference2, reference3)
    body1.InsertHybridShape(hybridShapePlaneNormal1)
    part1.update()


    ## 用rf_plane作为参考平面,画草图
    hybridShapes1 = body1.HybridShapes
    reference4 = hybridShapes1.Item("平面.1")
    sketch2 = sketches1.Add(reference4)
    factory2D2 = sketch2.OpenEdition()

    # 创建截面形状,此例用圆
    circle2D1 = factory2D2.CreateClosedEllipse(0,0,1,0,ellipse_x,ellipse_y)
    sketch2.closeedition()

    ## 创建扫掠体
    reference6 = part1.CreateReferenceFromObject(sketch2)
    reference7 = part1.CreateReferenceFromObject(sketch1)
    hybridShapeSweepExplicit1 = hybridShapeFactory1.AddNewSweepExplicit(reference6, reference7)
    hybridShapeSweepExplicit1.SubType = 1
    hybridShapeSweepExplicit1.SetAngleRef(1, 0.000000)
    hybridShapeSweepExplicit1.SolutionNo = 0
    hybridShapeSweepExplicit1.SmoothActivity = False
    hybridShapeSweepExplicit1.GuideDeviationActivity = False
    hybridShapeSweepExplicit1.Context = 1
    hybridShapeSweepExplicit1.SetbackValue = 0.020000
    hybridShapeSweepExplicit1.FillTwistedAreas = 1
    orderedGeometricalSets1 = body1.OrderedGeometricalSets
    orderedGeometricalSet1 = orderedGeometricalSets1.add()
    orderedGeometricalSet1.InsertHybridShape(hybridShapeSweepExplicit1)
    part1.update()



catia = win32com.client.Dispatch("catia.application")
partdocument1 = catia.documents.add("Part")#Part类型的文档,也就是零件文档
part1 = partdocument1.part

#control_list为控制点列表
controlpoint_list = [(0, 0),(10, 10),(20, 0),(30, 10),(40,0),(50,10)]
create_spline_and_sweep(controlpoint_list,ellipse_x=2,ellipse_y=1)