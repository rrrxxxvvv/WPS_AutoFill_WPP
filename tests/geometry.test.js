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
assert.ok(sampledCurve.points.length > 20, '贝塞尔曲线应按控制点采样，而不是退化成端点直线')
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
