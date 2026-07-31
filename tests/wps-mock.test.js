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

const shapes = {
  items: [],
  addPolylineCalls: 0,
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
    const shape = makeShape({
      Type: 5,
      AutoShapeType: -2,
      Left: left,
      Top: top,
      Width: right - left,
      Height: bottom - top,
      _points: points,
      _createdWith: 'AddPolyline'
    }, shapes)
    shapes.items.push(shape)
    return shape
  },
  BuildFreeform(editingType, x, y) {
    const points = [{ x, y }]
    return {
      AddNodes(segmentType, nodeEditingType, x1, y1, x2, y2, x3, y3) {
        // WPS 实际使用中会让省略的尾部坐标落到 0；这里按 X3/Y3
        // 作为最终节点来模拟，以捕获只传 X1/Y1 的回归。
        points.push({
          x: Number.isFinite(x3) ? x3 : 0,
          y: Number.isFinite(y3) ? y3 : 0
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
          _points: points
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
assert.strictEqual(previews.length, 3)
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
assert.strictEqual(captured.changed, true)
assert.strictEqual(captured.count, 1)
assert.ok(previews[0].Name.startsWith(engine.PICKED_PREVIEW_PREFIX))
assert.ok(Math.abs(previews[0].Fill.Transparency - 0.54) < 1e-9, '已点选候选区域应加深显示')
assert.strictEqual(unselectCount, 2, '点选候选区域后应立即清除 WPS 外接矩形控制框')

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
selected = [previews[1]]
captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.count, 2, '无需 Ctrl 即可连续记录多个候选区域')
assert.strictEqual(unselectCount, 3)

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
selected = [previews[0]]
captured = engine.captureManualSelection({ color: '#4f7cff' })
assert.strictEqual(captured.count, 1, '再次点击已点选区域应取消选择')
assert.ok(previews[0].Name.startsWith(engine.PREVIEW_PREFIX))
assert.ok(!previews[0].Name.startsWith(engine.PICKED_PREVIEW_PREFIX))
assert.strictEqual(unselectCount, 4)

selected = []
engine.captureManualSelection({ color: '#4f7cff' })
const generated = engine.generateSelected({ color: '#ff6600', opacity: 72 })
assert.strictEqual(generated.count, 1)
assert.strictEqual(unselectCount, 5, '提交点选区域后应清除控制柄')
assert.strictEqual(shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX)).length, 0)
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
assert.strictEqual(unselectCount, 6)
const networkPreviews = shapes.items.filter(shape => shape.Name.startsWith(engine.PREVIEW_PREFIX))
assert.strictEqual(networkPreviews.length, 5)
assert.ok(networkPreviews.every(shape => shape._points.length > 4), '线网面和椭圆内部都不应退化成圆心三角形')
assert.ok(networkNodes.every(shape => shape.Fill.Visible === 0), '生成节点内部区域时应关闭原椭圆填充并保留轮廓')

delete global.Application
console.log('WPS mock workflow tests passed')
