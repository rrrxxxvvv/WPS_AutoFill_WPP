const assert = require('assert')
const engine = require('../js/fill-engine.js')

function makeShape(properties, collection) {
  const shape = Object.assign({
    Name: '',
    Type: 1,
    AutoShapeType: 1,
    Left: 0,
    Top: 0,
    Width: 10,
    Height: 10,
    Rotation: 0,
    HorizontalFlip: 0,
    VerticalFlip: 0,
    Fill: {
      Visible: -1,
      Transparency: 0,
      ForeColor: { RGB: 0 },
      Solid() {}
    },
    Line: { Visible: -1 },
    ZOrder() {},
    Delete() {
      const index = collection.items.indexOf(shape)
      if (index >= 0) collection.items.splice(index, 1)
    }
  }, properties)
  return shape
}

function makeNodeCollection(nodes) {
  return {
    get Count() { return nodes.length },
    Item(index) { return nodes[index - 1] },
    SetSegmentType(index, segmentType) { nodes[index - 1].SegmentType = segmentType },
    SetEditingType(index, editingType) { nodes[index - 1].EditingType = editingType }
  }
}

const shapes = {
  items: [],
  addPolylineCalls: 0,
  addCurveCalls: 0,
  buildFreeformCalls: 0,
  misplaceBuildCount: 0,
  flattenCurveBuildCount: 0,
  get Count() { return this.items.length },
  Item(index) { return this.items[index - 1] },
  AddPolyline(pointRows) {
    this.addPolylineCalls += 1
    const points = pointRows.map(point => ({ x: point[0], y: point[1] }))
    const xs = points.map(point => point.x)
    const ys = points.map(point => point.y)
    const left = Math.min(...xs)
    const top = Math.min(...ys)
    const right = Math.max(...xs)
    const bottom = Math.max(...ys)
    const nodes = points.map(point => ({ SegmentType: 0, EditingType: 1, Points: [[point.x, point.y]] }))
    const shape = makeShape({
      Type: 5,
      AutoShapeType: -2,
      Left: left,
      Top: top,
      Width: right - left,
      Height: bottom - top,
      _points: points,
      _createdWith: 'AddPolyline',
      Nodes: makeNodeCollection(nodes)
    }, shapes)
    shapes.items.push(shape)
    return shape
  },
  AddCurve(pointRows) {
    this.addCurveCalls += 1
    const points = pointRows.map(point => ({ x: point[0], y: point[1] }))
    assert.ok(points.length >= 4 && (points.length - 1) % 3 === 0,
      'AddCurve 点列必须是起点加若干组三次贝塞尔控制点/终点')
    const sampled = [points[0]]
    for (let i = 1; i < points.length; i += 3) {
      const p0 = points[i - 1]
      const p1 = points[i]
      const p2 = points[i + 1]
      const p3 = points[i + 2]
      for (let step = 1; step <= 32; step += 1) {
        const t = step / 32
        const mt = 1 - t
        sampled.push({
          x: mt * mt * mt * p0.x + 3 * mt * mt * t * p1.x + 3 * mt * t * t * p2.x + t * t * t * p3.x,
          y: mt * mt * mt * p0.y + 3 * mt * mt * t * p1.y + 3 * mt * t * t * p2.y + t * t * t * p3.y
        })
      }
    }
    const xs = sampled.map(point => point.x)
    const ys = sampled.map(point => point.y)
    const left = Math.min(...xs)
    const top = Math.min(...ys)
    const right = Math.max(...xs)
    const bottom = Math.max(...ys)
    const shape = makeShape({
      Type: 5,
      AutoShapeType: -2,
      Left: left,
      Top: top,
      Width: right - left,
      Height: bottom - top,
      _points: points,
      _createdWith: 'AddCurve'
    }, shapes)
    shapes.items.push(shape)
    return shape
  },
  AddShape(autoShapeType, left, top, width, height) {
    const shape = makeShape({
      Type: 1,
      AutoShapeType: autoShapeType,
      Left: left,
      Top: top,
      Width: width,
      Height: height,
      _createdWith: 'AddShape'
    }, shapes)
    shapes.items.push(shape)
    return shape
  },
  AddLine(x1, y1, x2, y2) {
    const shape = makeShape({
      Type: 9,
      AutoShapeType: -2,
      Left: Math.min(x1, x2),
      Top: Math.min(y1, y2),
      Width: Math.abs(x2 - x1),
      Height: Math.abs(y2 - y1),
      _points: [{ x: x1, y: y1 }, { x: x2, y: y2 }],
      _createdWith: 'AddLine'
    }, shapes)
    shapes.items.push(shape)
    return shape
  },
  BuildFreeform(editingType, x, y) {
    this.buildFreeformCalls += 1
    const misplaced = this.misplaceBuildCount > 0
    if (misplaced) this.misplaceBuildCount -= 1
    const flattenCurves = this.flattenCurveBuildCount > 0
    if (flattenCurves) this.flattenCurveBuildCount -= 1
    const coordinateOffset = misplaced ? 5000 : 0
    const points = [{ x: x + coordinateOffset, y: y + coordinateOffset }]
    const nodes = []
    return {
      AddNodes(segmentType, nodeEditingType, x1, y1, x2, y2, x3, y3) {
        // WPS 官方 JS API：直线段只传终点 X1/Y1；显式三次曲线传
        // 两个控制点和终点 X3/Y3。
        const endX = segmentType === 1 && Number.isFinite(x3) ? x3 : x1
        const endY = segmentType === 1 && Number.isFinite(y3) ? y3 : y1
        points.push({
          x: endX + coordinateOffset,
          y: endY + coordinateOffset
        })
        const effectiveSegmentType = flattenCurves && segmentType === 1 ? 0 : segmentType
        nodes.push({
          SegmentType: effectiveSegmentType,
          EditingType: nodeEditingType,
          Points: [[endX + coordinateOffset, endY + coordinateOffset]],
          segmentType, x1, y1, x2, y2, x3, y3
        })
      },
      ConvertToShape() {
        const xs = points.map(point => point.x)
        const ys = points.map(point => point.y)
        const left = Math.min(...xs)
        const top = Math.min(...ys)
        const right = Math.max(...xs)
        const bottom = Math.max(...ys)
        const shape = makeShape({
          Type: 5,
          AutoShapeType: -2,
          Left: left,
          Top: top,
          Width: right - left,
          Height: bottom - top,
          _points: points,
          _nodes: nodes,
          _createdWith: 'BuildFreeform',
          Nodes: makeNodeCollection(nodes)
        }, shapes)
        shapes.items.push(shape)
        return shape
      }
    }
  }
}

const boundary = makeShape({
  Name: '矩形 1',
  Type: 1,
  AutoShapeType: 1,
  Left: 0,
  Top: 0,
  Width: 200,
  Height: 120
}, shapes)
const divider1 = makeShape({
  Name: '直线 1',
  Type: 9,
  Left: -20,
  Top: -20,
  Width: 120,
  Height: 120,
  HorizontalFlip: -1
}, shapes)
const divider2 = makeShape({
  Name: '直线 2',
  Type: 9,
  Left: -20,
  Top: -20,
  Width: 120,
  Height: 60,
  HorizontalFlip: -1
}, shapes)
shapes.items.push(boundary, divider1, divider2)

const slide = { Shapes: shapes }
// 普通外框流程只选中外框，插件应自动读取当前页分割线，无需 Ctrl 多选。
let selected = [boundary]
let nullShapeRange = false
let unselectCount = 0
const selection = {
  get ShapeRange() {
    if (nullShapeRange) return null
    return {
      Count: selected.length,
      Item(index) { return selected[index - 1] }
    }
  },
  Unselect() {
    selected = []
    unselectCount += 1
  }
}
global.Application = {
  ActivePresentation: {},
  ActiveWindow: {
    Selection: selection,
    View: { Slide: slide }
  }
}

const analysis = engine.analyze({ segments: 96 })
assert.strictEqual(analysis.boundaryKind, '多边形')
assert.strictEqual(analysis.regionCount, 3)

const prepared = engine.prepareManualSelection({
  color: '#4f7cff',
  segments: 96,
  disableBoundaryFill: true
})
assert.strictEqual(prepared.count, 3)
assert.strictEqual(unselectCount, 1, '批量创建候选区域后应清除 WPS 自动累积的选择状态')
assert.strictEqual(shapes.addPolylineCalls, 3, '应优先使用 WPS 中坐标更稳定的 AddPolyline')
assert.strictEqual(boundary.Fill.Visible, 0, '准备手动选择时应关闭原外框填充')
const previews = shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX))
const markers = shapes.items.filter(shape => shape.Name.startsWith(engine.MARKER_PREFIX))
assert.strictEqual(previews.length, 3)
assert.strictEqual(markers.length, 3, '每个候选区域都应创建一个独立的小型点选标记')
markers.forEach(marker => {
  assert.strictEqual(marker._createdWith, 'AddShape')
  assert.ok(marker.Width <= 14 && marker.Height <= 14, '点选标记的 WPS 外接框必须保持很小')
  const key = marker.Name.slice(engine.MARKER_PREFIX.length)
  const preview = previews.find(shape => shape.Name === engine.PREVIEW_PREFIX + key)
  assert.ok(preview, '每个点选标记都必须对应一个候选区域')
  const center = { x: marker.Left + marker.Width / 2, y: marker.Top + marker.Height / 2 }
  assert.ok(engine._geometry.pointInPolygon(center, preview._points, 1e-6), '点选标记中心必须位于对应候选区域内部')
})
previews.forEach(shape => {
  assert.strictEqual(shape.Line.Visible, 0, '候选区域不应有边框')
  assert.ok(Math.abs(shape.Fill.Transparency - 0.82) < 1e-9, '候选区域应以可见浅色显示')
  shape._points.forEach(point => {
    assert.ok(point.x >= -0.001 && point.x <= 200.001, '自由形状节点不应被 WPS 默认参数拉到区域之外')
    assert.ok(point.y >= -0.001 && point.y <= 120.001, '自由形状节点不应被 WPS 默认参数拉到区域之外')
  })
})

selected = [previews[0]]
let captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.changed, false, '直接点中复杂候选形状时不应据此选择区域')
assert.ok(captured.hint.includes('圆形定位点'))
assert.strictEqual(unselectCount, 2, '误点候选形状后应立即清除其大型外接矩形')

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
selected = [markers[0]]
captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.changed, true)
assert.strictEqual(captured.count, 1)
assert.ok(previews[0].Name.startsWith(engine.PICKED_PREVIEW_PREFIX))
assert.ok(markers[0].Name.startsWith(engine.PICKED_MARKER_PREFIX))
assert.ok(Math.abs(previews[0].Fill.Transparency - 0.54) < 1e-9, '已点选候选区域应加深显示')
assert.strictEqual(unselectCount, 3, '点选小型定位点后应立即清除其控制框')

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
selected = [markers[1]]
captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.count, 2, '无需 Ctrl 即可连续记录多个候选区域')
assert.strictEqual(unselectCount, 4)

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
selected = [markers[0]]
captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.count, 1, '再次点击已点选区域应取消选择')
assert.ok(previews[0].Name.startsWith(engine.PREVIEW_PREFIX))
assert.ok(!previews[0].Name.startsWith(engine.PICKED_PREVIEW_PREFIX))
assert.ok(!markers[0].Name.startsWith(engine.PICKED_MARKER_PREFIX))
assert.strictEqual(unselectCount, 5)

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
const generated = engine.generateSelected({ color: '#ff6600', opacity: 72 })
assert.strictEqual(generated.count, 1)
assert.strictEqual(unselectCount, 6, '提交点选区域后应清除控制柄')
assert.strictEqual(shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX)).length, 0)
assert.strictEqual(shapes.items.filter(shape => shape.Name.startsWith(engine.MARKER_PREFIX)).length, 0)
const committed = shapes.items.filter(shape => shape.Name.startsWith(engine.PREFIX))
assert.strictEqual(committed.length, 1)
assert.strictEqual(committed[0].Line.Visible, 0)
assert.ok(Math.abs(committed[0].Fill.Transparency - 0.28) < 1e-9)

assert.strictEqual(engine.deleteGenerated(), 1)
assert.strictEqual(shapes.items.filter(shape => shape.Name.startsWith(engine.PREFIX)).length, 0)

nullShapeRange = true
const networkNodes = [
  [500, 0], [600, 0], [600, 100], [500, 100]
].map((point, index) => makeShape({
  Name: '椭圆节点 ' + (index + 1),
  Type: 1,
  AutoShapeType: 9,
  Left: point[0] - 10,
  Top: point[1] - 10,
  Width: 20,
  Height: 20
}, shapes))
const networkLines = [
  { name: '上边', left: 510, top: 0, width: 80, height: 0 },
  { name: '右边', left: 560, top: 50, width: 80, height: 0, rotation: 90 },
  { name: '下边', left: 510, top: 100, width: 80, height: 0 },
  { name: '左边', left: 500, top: 10, width: 0, height: 80 }
].map(item => makeShape({
  Name: item.name,
  Type: 9,
  Left: item.left,
  Top: item.top,
  Width: item.width,
  Height: item.height,
  Rotation: item.rotation || 0
}, shapes))
shapes.items.push(...networkNodes, ...networkLines)

const networkAnalysis = engine.analyze({ segments: 96 })
assert.strictEqual(networkAnalysis.mode, 'network')
assert.strictEqual(networkAnalysis.scannedWholeSlide, true)
assert.strictEqual(networkAnalysis.nodeCount, 4)
assert.strictEqual(networkAnalysis.faceCount, 1)
assert.strictEqual(networkAnalysis.nodeRegionCount, 4)
assert.strictEqual(networkAnalysis.regionCount, 5)
assert.strictEqual(networkAnalysis.ignoredLineCount, 2, '不连接椭圆节点的页面线段应被忽略')

const networkPrepared = engine.prepareManualSelection({
  color: '#4f7cff',
  segments: 96,
  includeNodeInteriors: true,
  disableBoundaryFill: true
})
assert.strictEqual(networkPrepared.count, 5)
assert.strictEqual(unselectCount, 7)
const networkPreviews = shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX))
assert.strictEqual(networkPreviews.length, 5)
assert.ok(networkPreviews.every(shape => shape._points.length > 4), '线网面和椭圆内部都不应退化成圆心三角形')
assert.ok(networkNodes.every(shape => shape.Fill.Visible === 0), '生成节点内部区域时应关闭原椭圆填充并保留轮廓')

const genericA = makeShape({
  Name: '矩形 A',
  Type: 1,
  AutoShapeType: 1,
  Left: 700,
  Top: 0,
  Width: 100,
  Height: 100
}, shapes)
const genericB = makeShape({
  Name: '矩形 B',
  Type: 1,
  AutoShapeType: 1,
  Left: 750,
  Top: 0,
  Width: 100,
  Height: 100
}, shapes)
const genericCurve = makeShape({
  Name: '曲线 1',
  Type: 5,
  AutoShapeType: -2,
  Left: 775,
  Top: -20,
  Width: 0,
  Height: 140,
  Vertices: [[775, -20], [775, 30], [775, 70], [775, 120]],
  Nodes: {
    Count: 2,
    Item(index) {
      return { SegmentType: 1, Points: index === 1 ? [[775, -20]] : [[775, 120]] }
    }
  }
}, shapes)
const genericGroup = makeShape({
  Name: '组合线稿 1',
  Type: 6,
  AutoShapeType: -2,
  GroupItems: {
    Count: 3,
    Item(index) { return [genericA, genericB, genericCurve][index - 1] }
  }
}, shapes)
shapes.items.push(genericGroup)
nullShapeRange = false
selected = [genericGroup]

const genericPrepared = engine.prepareManualSelection({
  color: '#4f7cff',
  segments: 96,
  disableBoundaryFill: true
})
assert.strictEqual(genericPrepared.mode, 'generic')
assert.strictEqual(genericPrepared.count, 4, '组合中的重叠多边形和自由曲线应递归展开并拆分闭合面')
assert.strictEqual(genericA.Fill.Visible, 0)
assert.strictEqual(genericB.Fill.Visible, 0)
const genericPreviews = shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX))
assert.strictEqual(genericPreviews.length, 4)
assert.ok(genericPreviews.every(shape => shape.Line.Visible === 0))

const snippedRectangle = makeShape({
  Name: '剪去单角的矩形 1134',
  Type: 1,
  AutoShapeType: 155,
  Left: 900,
  Top: 0,
  Width: 200,
  Height: 140
}, shapes)
const halfFrame = makeShape({
  Name: '半闭框 1135',
  Type: 1,
  AutoShapeType: 159,
  Left: 1020,
  Top: 50,
  Width: 240,
  Height: 170
}, shapes)
const curvedConnector = makeShape({
  Name: '曲线连接符 1136',
  Type: 9,
  AutoShapeType: -2,
  Connector: -1,
  ConnectorFormat: { Type: 3 },
  Left: 980,
  Top: -30,
  Width: 180,
  Height: 280
}, shapes)
selected = [snippedRectangle, halfFrame]
const exactObjectsWithoutCurve = engine.analyze({ segments: 96 })
selected = [snippedRectangle, halfFrame, curvedConnector]
const exactObjectAnalysis = engine.analyze({ segments: 96 })
assert.strictEqual(exactObjectAnalysis.mode, 'generic')
assert.strictEqual(exactObjectAnalysis.lineCount, 3, '三个用户实际对象都必须成功读取为几何路径')
assert.ok(exactObjectAnalysis.regionCount >= 2, '两种闭合自动形状至少应各自产生一个候选区域')
assert.ok(
  exactObjectAnalysis.regionCount > exactObjectsWithoutCurve.regionCount,
  '曲线连接符应按视觉曲线切开闭合区域，不依赖 WPS 连接点'
)

const addCurveCallsBeforeCurvedRegions = shapes.addCurveCalls
const exactObjectPrepared = engine.prepareManualSelection({
  color: '#4f7cff',
  segments: 96,
  disableBoundaryFill: true
})
assert.strictEqual(exactObjectPrepared.mode, 'generic')
assert.ok(exactObjectPrepared.count >= 2)
assert.strictEqual(snippedRectangle.Fill.Visible, 0)
assert.strictEqual(halfFrame.Fill.Visible, 0)
assert.ok(shapes.addCurveCalls > addCurveCallsBeforeCurvedRegions,
  '包含真实曲边的候选区域必须优先调用 AddCurve 写入完整贝塞尔点列')
const exactCurvePreviews = shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX))
assert.ok(exactCurvePreviews.some(shape => shape._createdWith === 'AddCurve'),
  '至少一个被曲线切分的候选区域应通过 AddCurve 保留真实曲线')
assert.ok(exactCurvePreviews
  .filter(shape => shape._createdWith === 'AddCurve')
  .every(shape => shape._points.length >= 4 && (shape._points.length - 1) % 3 === 0),
  'AddCurve 候选区域必须包含完整三次贝塞尔点列')
selected = [snippedRectangle, halfFrame, curvedConnector]
const debugResult = engine.generateGeometryDebug({ segments: 96 })
assert.ok(debugResult.intersectionCount >= 1, '几何调试应在轮廓与曲线交汇处生成交点')
assert.ok(debugResult.lineSegmentCount >= 1, '几何调试应生成交点之间的独立直线段')
assert.ok(debugResult.curveSegmentCount >= 1, '几何调试应生成交点之间的独立曲线段')
assert.strictEqual(debugResult.curveRejected, 0, '参数采样调试曲线不应触发曲线节点拒绝')
assert.ok(shapes.items.some(shape => shape.Name.startsWith(engine.DEBUG_PREFIX + 'Point_')))
assert.ok(shapes.items.some(shape => shape.Name.startsWith(engine.DEBUG_PREFIX + 'Line_')))
assert.ok(shapes.items.some(shape => shape.Name.startsWith(engine.DEBUG_PREFIX + 'Curve_')))
const sampledDebugCurves = shapes.items.filter(shape => shape.Name.startsWith(engine.DEBUG_PREFIX + 'Curve_'))
assert.ok(sampledDebugCurves.every(shape => shape._createdWith === 'AddPolyline' && shape._points.length >= 7),
  '调试曲线必须沿原曲线参数区间采样截取，不能重新用 BuildFreeform 猜测控制点')
selected = [curvedConnector]
const diagnostic = JSON.parse(engine.inspectSelectedGeometry())
assert.strictEqual(diagnostic.shapes.length, 1)
assert.deepStrictEqual(diagnostic.curveCreationApi, {
  addCurve: true,
  buildFreeform: true,
  addPolyline: true
})
assert.strictEqual(diagnostic.shapes[0].connectorFormat.type, 3)
assert.strictEqual(diagnostic.shapes[0].pluginCurveInterpretation.source, 'preset:curvedConnector3',
  'Vertices/Nodes 不可读时应明确使用标准 curvedConnector3 预设几何')

const simpleClosedShape = makeShape({
  Id: 9001,
  Name: '剪去同侧角的矩形',
  Type: 1,
  AutoShapeType: 156
}, shapes)
const compoundClosedShape = makeShape({
  Id: 9002,
  Name: '禁止符',
  Type: 1,
  AutoShapeType: 19
}, shapes)
const closedPath = points => ({
  points,
  closed: true,
  curved: false,
  edgeMeta: points.map(() => ({ kind: 'line' })),
  stroke: true
})
engine.setGeometrySnapshot({
  geometryForShape(shape) {
    if (shape.Id === 9001) {
      return [closedPath([
        { x: 0, y: 0 }, { x: 120, y: 0 }, { x: 120, y: 100 }, { x: 0, y: 100 }
      ])]
    }
    if (shape.Id === 9002) {
      return [
        closedPath([{ x: 50, y: -20 }, { x: 170, y: -20 }, { x: 170, y: 120 }, { x: 50, y: 120 }]),
        closedPath([{ x: 75, y: 10 }, { x: 145, y: 10 }, { x: 145, y: 90 }, { x: 75, y: 90 }]),
        closedPath([{ x: 65, y: 5 }, { x: 155, y: 95 }, { x: 145, y: 105 }, { x: 55, y: 15 }])
      ]
    }
    return []
  }
})
selected = [simpleClosedShape, compoundClosedShape]
const compoundOverlapAnalysis = engine.analyze({ segments: 96 })
assert.strictEqual(compoundOverlapAnalysis.mode, 'generic',
  '一个普通闭合形状与一个复合闭合形状同时选中时不得误判为外框模式')
assert.ok(compoundOverlapAnalysis.regionCount > 1,
  '剪角矩形与禁止符重叠时必须识别多个候选区域')
engine.setGeometrySnapshot(null)

selected = [1, 2, 3].map(index => makeShape({
  Name: '不支持对象 ' + index,
  Type: 13,
  AutoShapeType: -2
}, shapes))
assert.throws(
  () => engine.analyze({ segments: 96 }),
  error => {
    assert.ok(error.message.includes('扫描了 3 个对象：成功读取 0 条路径'))
    assert.ok(error.message.includes('未读取对象：'))
    assert.ok(!error.message.includes('。；'), '错误提示不应出现重复或冲突标点')
    return true
  }
)

delete global.Application
console.log('WPS mock workflow tests passed')
