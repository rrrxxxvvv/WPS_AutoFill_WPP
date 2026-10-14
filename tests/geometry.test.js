const assert = require('assert')
const engine = require('../js/fill-engine.js')
const g = engine._geometry

function ellipse() {
  return { cx: 0, cy: 0, rx: 100, ry: 60 }
}

function line(x1, y1, x2, y2) {
  return { p1: { x: x1, y: y1 }, p2: { x: x2, y: y2 } }
}

const e = ellipse()
let regions = g.buildRegions(e, [line(-150, 0, 150, 0)], 160)
assert.strictEqual(regions.length, 2, '一条直线应分成两个区域')

regions = g.buildRegions(e, [line(-150, 0, 150, 0), line(0, -100, 0, 100)], 160)
assert.strictEqual(regions.length, 4, '两条相交直线应分成四个区域')

regions = g.buildRegions(e, [
  line(-150, 0, 150, 0),
  line(0, -100, 0, 100),
  line(-150, -50, 150, 50)
], 160)
assert.strictEqual(regions.length, 6, '三条直线中有共点情形时应得到六个区域')

const totalArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
const ellipseArea = Math.PI * e.rx * e.ry
assert.ok(Math.abs(totalArea - ellipseArea) / ellipseArea < 0.002, '区域总面积应接近椭圆面积')

const rectangle = [
  { x: 0, y: 0 },
  { x: 200, y: 0 },
  { x: 200, y: 120 },
  { x: 0, y: 120 }
]
regions = g.buildRegions(rectangle, [
  line(-20, 100, 100, -20),
  line(-20, 40, 100, -20)
])
assert.strictEqual(regions.length, 3, '交点位于矩形外部的两条直线应形成三个区域')
const rectangleArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
assert.ok(Math.abs(rectangleArea - 24000) < 0.01, '矩形切分后不应产生整页或外框之外的面积')

const concave = [
  { x: 0, y: 0 },
  { x: 100, y: 0 },
  { x: 100, y: 30 },
  { x: 30, y: 30 },
  { x: 30, y: 100 },
  { x: 0, y: 100 }
]
regions = g.buildRegions(concave, [line(15, -20, 15, 120)])
assert.strictEqual(regions.length, 2, '凹自由形状也应能按直线切分')
const concaveArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
assert.ok(Math.abs(concaveArea - 5100) < 0.01, '凹形切分后的区域总面积应等于原外框面积')

const concaveInterior = g.polygonInteriorPoint(concave)
assert.ok(g.pointInPolygon(concaveInterior.point, concave, 1e-6), '凹形区域的点选定位点必须位于真实轮廓内部')
assert.ok(concaveInterior.clearance > 0, '凹形区域的点选定位点应与边界保持正距离')
assert.ok(!(concaveInterior.point.x > 30 && concaveInterior.point.y > 30), '定位点不能落入凹形外接矩形中不属于图形的缺口')

const overlappingPaths = [
  { closed: true, points: [
    { x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 100 }, { x: 0, y: 100 }
  ] },
  { closed: true, points: [
    { x: 50, y: 0 }, { x: 150, y: 0 }, { x: 150, y: 100 }, { x: 50, y: 100 }
  ] }
]
regions = g.buildGenericRegions(overlappingPaths)
assert.strictEqual(regions.length, 3, '两个重叠多边形应按交点拆成三个独立闭合区域')
const overlappingArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
assert.ok(Math.abs(overlappingArea - 15000) < 0.01, '通用线稿区域总面积应等于重叠多边形的并集面积')

regions = g.buildGenericRegions([
  overlappingPaths[0],
  { closed: false, points: [{ x: 50, y: -20 }, { x: 50, y: 120 }] }
])
assert.strictEqual(regions.length, 2, '穿过闭合多边形的开放路径应将其切成两个区域')
const openDividerArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
assert.ok(Math.abs(openDividerArea - 10000) < 0.01, '开放路径不应在闭合轮廓外生成区域')

const curveShape = {
  Name: '曲线 1',
  Vertices: {
    Count: 4,
    Item(index) {
      const point = [[0, 0], [30, 100], [70, -100], [100, 0]][index - 1]
      return {
        Count: 2,
        Item(coordinateIndex) { return point[coordinateIndex - 1] }
      }
    }
  },
  Nodes: {
    Count: 2,
    Item(index) {
      return { SegmentType: 1, Points: index === 1 ? [[0, 0]] : [[100, 0]] }
    }
  }
}
const sampledCurve = g.sampleFreeformPath(curveShape, 96)
assert.strictEqual(sampledCurve.closed, false)
assert.strictEqual(sampledCurve.points.length, 97, '快速档应按 96 段采样贝塞尔曲线，而不是受旧的 48 段上限限制')
assert.ok(sampledCurve.points.some(point => Math.abs(point.y) > 20), '曲线采样结果应保留实际弯曲')

const graphVertices = [
  { x: 0, y: 0 },
  { x: 100, y: 0 },
  { x: 100, y: 100 },
  { x: 0, y: 100 },
  { x: 50, y: 50 }
]
const graphEdges = [
  { u: 0, v: 1 }, { u: 1, v: 2 }, { u: 2, v: 3 }, { u: 3, v: 0 },
  { u: 0, v: 4 }, { u: 1, v: 4 }, { u: 2, v: 4 }, { u: 3, v: 4 }
]
regions = g.enumerateGraphFaces(graphVertices, graphEdges)
assert.strictEqual(regions.length, 4, '椭圆节点线网的四个内部闭合面都应被识别')
const networkArea = regions.reduce((sum, p) => sum + Math.abs(g.polygonArea(p)), 0)
assert.ok(Math.abs(networkArea - 10000) < 0.01, '线网识别必须排除最外层无限区域')

const curvedNodes = [
  { cx: 0, cy: 0, rx: 10, ry: 10 },
  { cx: 100, cy: 0, rx: 10, ry: 10 },
  { cx: 50, cy: 100, rx: 10, ry: 10 }
]
const curvedFace = g.buildCurvedNetworkFace([0, 1, 2], curvedNodes, 160)
assert.ok(curvedFace.length > 12, '节点处应使用高密度椭圆弧，而不是只保留三个圆心')
assert.ok(Math.abs(g.polygonArea(curvedFace)) < 5000, '曲边面应扣除三个椭圆节点占据的角部')
curvedFace.forEach(point => {
  const distanceToNearestCenter = Math.min(...curvedNodes.map(node => Math.hypot(point.x - node.cx, point.y - node.cy)))
  assert.ok(distanceToNearestCenter >= 9.999, '曲边区域不能穿过椭圆内部或圆心')
})

const halfFrameOutline = g.adjustedNormalizedPolygon({
  AutoShapeType: 159,
  Adjustments: { Item() { throw new Error('使用默认调节点') } }
}, { width: 240, height: 170 })
const halfFrameInset = (170 / 3) / 240
assert.ok(Math.abs(halfFrameOutline[2][0] - (2 / 3)) < 1e-6, '半闭框右侧斜边必须按纵向调节值换算 x2')
assert.ok(Math.abs(halfFrameOutline[4][1] - (1 - halfFrameInset)) < 1e-6, '半闭框内侧竖边必须按横向调节值换算 y2')
const halfFrameSentinelAdjustment = g.adjustedNormalizedPolygon({
  AutoShapeType: 159,
  Adjustments: { Item() { return 1 } }
}, { width: 240, height: 170 })
assert.deepStrictEqual(halfFrameSentinelAdjustment, halfFrameOutline,
  'WPS 12.1 对不同预设统一返回的调整值 1 必须视为无效包装值并采用标准默认值')
const realWpsHalfFrame = g.adjustedNormalizedPolygon({
  AutoShapeType: 159,
  Adjustments: { Item() { return 1 } }
}, { width: 221.10000610351562, height: 138.77102661132812 })
assert.ok(Math.abs(realWpsHalfFrame[2][0] - (2 / 3)) < 1e-6,
  '真实 WPS 半闭框的右上内点 x2 应由纵向默认调节值决定')
assert.ok(Math.abs(realWpsHalfFrame[4][1] - (1 - (138.77102661132812 / 3) / 221.10000610351562)) < 1e-6,
  '真实 WPS 半闭框的左下内点 y2 应按宽高比换算')

const realWpsSnip2Same = g.adjustedNormalizedPolygon({
  AutoShapeType: 156,
  Adjustments: { Item() { return 1 } }
}, { width: 215.4499969482422, height: 113.4000015258789 })
assert.strictEqual(realWpsSnip2Same.length, 8, '剪去同侧角矩形应生成标准八节点路径')
assert.ok(Math.abs(realWpsSnip2Same[0][0] - (113.4000015258789 / 6) / 215.4499969482422) < 1e-6,
  '剪去同侧角矩形的顶部切角必须按短边计算')
assert.deepStrictEqual(realWpsSnip2Same[4], [1, 1], '默认第二调节点为 0，底部不应凭空切角')

const snipPath = [
  { x: 0, y: 0 }, { x: 170, y: 0 }, { x: 200, y: 30 }, { x: 200, y: 140 }, { x: 0, y: 140 }
]
const halfFramePath = halfFrameOutline.map(point => ({ x: 120 + point[0] * 240, y: 50 + point[1] * 170 }))
const connectorCurve = []
for (let i = 0; i <= 48; i += 1) {
  const t = i / 48
  const mt = 1 - t
  connectorCurve.push({
    x: 80 * mt * mt * mt + 3 * 80 * mt * mt * t + 3 * 260 * mt * t * t + 260 * t * t * t,
    y: -30 * mt * mt * mt + 3 * 110 * mt * mt * t + 3 * 110 * mt * t * t + 250 * t * t * t
  })
}
const mixedRegions = g.buildGenericRegions([
  { closed: true, points: snipPath },
  { closed: true, points: halfFramePath },
  { closed: false, points: connectorCurve, curved: true }
])
assert.ok(mixedRegions.length >= 6, '直边自动形状与曲线连接符应共同形成完整闭合区域')
assert.ok(mixedRegions.some(region => region.length >= 12), '曲线切开的填色块必须保留足够多的曲线采样节点')

const localVertexConnector = {
  Name: '局部坐标曲线连接符',
  Type: 9,
  Left: 100,
  Top: 50,
  Width: 200,
  Height: 100,
  Rotation: 0,
  HorizontalFlip: 0,
  VerticalFlip: 0,
  Vertices: [[0, 0], [0, 50], [200, 50], [200, 100]]
}
const normalizedConnector = g.sampleCurvedConnectorPath(localVertexConnector, 96)
assert.ok(Math.hypot(normalizedConnector.points[0].x - 100, normalizedConnector.points[0].y - 50) < 1e-6,
  'WPS 返回局部坐标控制点时必须转换到幻灯片坐标')
assert.ok(Math.hypot(normalizedConnector.points.at(-1).x - 300, normalizedConnector.points.at(-1).y - 150) < 1e-6,
  '曲线终点转换后必须与连接符边界终点一致')

const realWpsConnector = g.sampleCurvedConnectorPath({
  Name: '曲线连接符 1136',
  Type: 1,
  AutoShapeType: -2,
  Connector: -1,
  ConnectorFormat: { Type: 3 },
  Left: 302.70001220703125,
  Top: 118.4000015258789,
  Width: 244.75,
  Height: 196.3000030517578,
  Rotation: 90,
  HorizontalFlip: 0,
  VerticalFlip: -1,
  Vertices: [],
  Adjustments: { Item() { return 1 } }
}, 96)
assert.strictEqual(realWpsConnector.curveSource, 'preset:curvedConnector3')
assert.strictEqual(realWpsConnector.points.length, 97, 'curvedConnector3 应保留两段连续三次贝塞尔曲线')
assert.ok(Math.hypot(realWpsConnector.points[0].x - 326.92501068115234,
  realWpsConnector.points[0].y - 94.17500305175781) < 1e-4,
'预设曲线起点必须正确应用垂直翻转和 90 度旋转')
assert.ok(Math.hypot(realWpsConnector.points.at(-1).x - 523.2250137329102,
  realWpsConnector.points.at(-1).y - 338.9250030517578) < 1e-4,
'预设曲线终点必须正确应用垂直翻转和 90 度旋转')
const curveSplitRegions = g.buildGenericRegions([
  { closed: true, points: [
    { x: 100, y: 50 }, { x: 300, y: 50 }, { x: 300, y: 150 }, { x: 100, y: 150 }
  ] },
  normalizedConnector
])
assert.strictEqual(curveSplitRegions.length, 2, '曲线从闭合框一角贯穿到另一角时应准确拆成两个区域')
curveSplitRegions.forEach(region => {
  const commands = g.freeformCommands(region)
  assert.ok(commands && commands.commands.some(command => command.kind === 'curve'),
    '区域拓扑必须保留曲线来源，不能退化成无语义折线点')
  assert.strictEqual(commands.commands.filter(command => command.kind === 'curve').length, 1,
    '同一段贝塞尔曲线的连续采样边应合并回一条真实曲线命令')
})

const screenshotSnip = [
  { x: 66, y: 149 }, { x: 447, y: 149 }, { x: 498, y: 200 },
  { x: 498, y: 456 }, { x: 66, y: 456 }
]
const screenshotHalfFrame = [
  { x: 333, y: 240 }, { x: 775, y: 240 }, { x: 682.333, y: 332.667 },
  { x: 425.667, y: 332.667 }, { x: 425.667, y: 425.333 }, { x: 333, y: 518 }
]
const screenshotCurve = g.sampleCurvedConnectorPath({
  Name: '曲线连接符 1136',
  Type: 9,
  Left: 248,
  Top: 47,
  Width: 393,
  Height: 489,
  Rotation: 0,
  HorizontalFlip: 0,
  VerticalFlip: 0
}, 160)
const screenshotRegions = g.buildGenericRegions([
  { closed: true, points: screenshotSnip },
  { closed: true, points: screenshotHalfFrame },
  screenshotCurve
])
assert.strictEqual(screenshotRegions.length, 8, '截图中的两个直边轮廓和 S 曲线应形成 8 个互不重叠区域')
assert.strictEqual(screenshotRegions.filter(region => g.freeformCommands(region)).length, 6,
  '截图中与曲线相邻的 6 个区域都必须保留真实贝塞尔边')
assert.ok(screenshotRegions.every(region => Math.abs(g.polygonArea(region)) > 1),
  '区域追踪不能产生跨越轮廓的退化或错误对角面')

const closeIntersectionRegions = g.buildGenericRegions([
  { closed: true, points: [
    { x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 100 }, { x: 0, y: 100 }
  ] },
  { closed: false, points: [{ x: 50, y: -10 }, { x: 50, y: 110 }] },
  { closed: false, points: [{ x: 50.35, y: -10 }, { x: 50.35, y: 110 }] }
])
assert.strictEqual(closeIntersectionRegions.length, 3,
  '0.35pt 间距的真实分割线不能被视觉容差合并：狭窄区域必须保留')
assert.ok(closeIntersectionRegions.some(region => Math.abs(Math.abs(g.polygonArea(region)) - 35) < 1e-6))

console.log('geometry tests passed:', {
  ellipseRegions: 6,
  ellipseRelativeAreaError: Math.abs(totalArea - ellipseArea) / ellipseArea,
  rectangleRegions: 3,
  concaveRegions: 2,
  genericOverlapRegions: 3,
  genericOpenDividerRegions: 2,
  sampledCurveVertices: sampledCurve.points.length,
  networkRegions: 4,
  curvedFaceVertices: curvedFace.length
})
