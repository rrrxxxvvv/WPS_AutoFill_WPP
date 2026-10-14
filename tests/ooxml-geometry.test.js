const assert = require('assert')
const fs = require('fs')
const path = require('path')
const os = require('os')

global.JSZip = require('jszip')
const geometry = require('../js/ooxml-geometry')
const fillEngine = require('../js/fill-engine')

const fixturePath = path.join(__dirname, 'fixtures', 'adjusted-shapes.pptx')
const definitionsPath = path.join(__dirname, '..', 'assets', 'presetShapeDefinitions.xml')

function mockShape(id, overrides) {
  return Object.assign({
    Id: id,
    Left: 10,
    Top: 20,
    Width: 200,
    Height: 100,
    Rotation: 0,
    HorizontalFlip: 0,
    VerticalFlip: 0
  }, overrides || {})
}

function assertFinitePaths(paths, label) {
  assert.ok(paths.length > 0, label + ' 应至少生成一条路径')
  paths.forEach((shapePath, pathIndex) => {
    assert.ok(shapePath.points.length >= (shapePath.closed ? 3 : 2), label + ' 路径点数不足')
    shapePath.points.forEach((point, pointIndex) => {
      assert.ok(Number.isFinite(point.x) && Number.isFinite(point.y),
        label + ' 第 ' + pathIndex + ' 条路径第 ' + pointIndex + ' 点无效')
    })
    assert.strictEqual(
      shapePath.edgeMeta.length,
      shapePath.closed ? shapePath.points.length : shapePath.points.length - 1,
      label + ' 的曲线元数据应与采样边一一对应'
    )
  })
}

async function main() {
  // Independent ray/ellipse equation: OOXML uses polar angles, whereas
  // Bezier construction needs ellipse parameter angles. A circle hides this bug.
  const arcCubics = geometry._internals.arcCubics
  const ray = (rx, ry, degrees) => {
    const a = degrees * Math.PI / 180
    const radius = 1 / Math.hypot(Math.cos(a) / rx, Math.sin(a) / ry)
    return { x: radius * Math.cos(a), y: radius * Math.sin(a) }
  }
  for (const [start, sweep] of [[35, 170], [200, -260], [-80, 360], [350, 40]]) {
    const origin = ray(180, 30, start)
    const arcs = arcCubics(origin, 180, 30, start * 60000, sweep * 60000)
    const end = ray(180, 30, start + sweep)
    const actual = arcs[arcs.length - 1].p3
    assert.ok(Math.hypot(actual.x - end.x, actual.y - end.y) < 1e-8,
      '非圆椭圆弧必须落在 OOXML 极角定义的端点')
    for (const arc of arcs) {
      for (let i = 0; i <= 20; i++) {
        const t = i / 20, u = 1 - t
        const x = u ** 3 * arc.p0.x + 3 * u * u * t * arc.p1.x + 3 * u * t * t * arc.p2.x + t ** 3 * arc.p3.x
        const y = u ** 3 * arc.p0.y + 3 * u * u * t * arc.p1.y + 3 * u * t * t * arc.p2.y + t ** 3 * arc.p3.y
        assert.ok(Math.abs((x / 180) ** 2 + (y / 30) ** 2 - 1) < 0.00002,
          '椭圆弧采样应满足原始椭圆方程')
      }
    }
  }
  const buffer = fs.readFileSync(fixturePath)
  const definitions = fs.readFileSync(definitionsPath, 'utf8')
  const original = await geometry.snapshotFromBuffer(buffer, 1, definitions)
  const adjusted = await geometry.snapshotFromBuffer(buffer, 2, definitions)

  assert.strictEqual(original.partPath, 'ppt/slides/slide1.xml')
  assert.strictEqual(adjusted.partPath, 'ppt/slides/slide2.xml')
  assert.strictEqual(Object.keys(original.shapes).length, 160)
  assert.strictEqual(Object.keys(adjusted.shapes).length, 160)
  const changedAdjustmentCount = Object.keys(original.shapes).filter(id => {
    return JSON.stringify(original.shapes[id].adjustments) !== JSON.stringify(adjusted.shapes[id].adjustments)
  }).length
  assert.strictEqual(changedAdjustmentCount, 113, '测试页应完整覆盖 113 个发生变化的黄点参数对象')

  assert.deepStrictEqual(original.shapes['20'].adjustments, {})
  assert.strictEqual(adjusted.shapes['20'].preset, 'roundRect')
  assert.deepStrictEqual(adjusted.shapes['20'].adjustments, { adj: 50000 })
  assert.strictEqual(adjusted.shapes['10'].adjustments.adj1, 75569)
  assert.deepStrictEqual(adjusted.shapes['44'].adjustments, { adj1: 54401, adj2: 83069 })

  const originalRoundRect = original.geometryForShape(mockShape(20), 160)
  const adjustedRoundRect = adjusted.geometryForShape(mockShape(20), 160)
  assertFinitePaths(originalRoundRect, '默认圆角矩形')
  assertFinitePaths(adjustedRoundRect, '调整后圆角矩形')
  assert.notDeepStrictEqual(
    originalRoundRect[0].points.slice(0, 6),
    adjustedRoundRect[0].points.slice(0, 6),
    '黄色调节点改变后应生成不同的圆角轮廓'
  )

  const freeCurve = adjusted.geometryForShape(mockShape(16), 160)
  const freePolyline = adjusted.geometryForShape(mockShape(17), 160)
  const freeSpline = adjusted.geometryForShape(mockShape(18), 160)
  assert.ok(freeCurve[0].curved, '三段自由贝塞尔曲线应保留曲线元数据')
  assert.strictEqual(freePolyline[0].curved, false, '直线任意多边形不应被误标为曲线')
  assert.strictEqual(freePolyline[0].points.length, 23, '22 条原始线段应生成 23 个端点')
  assert.ok(freeSpline[0].curved, '十九段自由贝塞尔曲线应保留曲线元数据')

  fillEngine.setGeometrySnapshot(adjusted)
  const engineCurvePaths = fillEngine._geometry.genericPathsFromShape(
    Object.assign(mockShape(16), { Name: '任意多边形 15', Type: 5, AutoShapeType: 138 }),
    160
  )
  assert.ok(engineCurvePaths.length > 0 && engineCurvePaths[0].curved,
    '填色引擎应优先采用 PPTX 中的真实自由曲线，不再读取空 Nodes')
  const engineHalfFrame = fillEngine._geometry.genericPathsFromShape(
    Object.assign(mockShape(44), { Name: '半闭框 43', Type: 1, AutoShapeType: 159 }),
    160
  )
  assert.ok(engineHalfFrame.some(shapePath => shapePath.closed),
    '填色引擎应使用 PPTX 黄点参数重建半闭框')

  const allIds = Object.keys(adjusted.shapes)
  allIds.forEach(id => {
    assertFinitePaths(adjusted.geometryForShape(mockShape(id), 160), adjusted.shapes[id].name)
  })

  const rotated = adjusted.geometryForShape(mockShape(20, { Rotation: 90, HorizontalFlip: -1 }), 96)
  assertFinitePaths(rotated, '旋转翻转后的圆角矩形')

  let saveCopyFormat = null
  let removedTemporaryCopy = false
  geometry.setPresetDefinitions(definitions)
  global.wps = {
    Env: { GetTempPath: () => os.tmpdir() },
    FileSystem: {
      ReadFileAsArrayBuffer(filePath) {
        const bytes = fs.readFileSync(filePath)
        return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
      },
      unlinkSync(filePath) {
        fs.unlinkSync(filePath)
        removedTemporaryCopy = true
      }
    }
  }
  const runtimeSnapshot = await geometry.refreshFromApplication({
    ActivePresentation: {
      SaveCopyAs(destination, format) {
        saveCopyFormat = format
        fs.copyFileSync(fixturePath, destination)
      }
    },
    ActiveWindow: { View: { Slide: { SlideIndex: 2 } } }
  })
  assert.strictEqual(saveCopyFormat, 24, '应使用 Open XML 演示文稿格式创建临时副本')
  assert.ok(removedTemporaryCopy, '读取完成后应清理临时 PPTX 副本')
  assert.strictEqual(runtimeSnapshot.shapes['20'].adjustments.adj, 50000)
  delete global.wps

  assert.strictEqual(geometry.evaluateFormula('*/ 10 20 5', {}), 40)
  assert.strictEqual(geometry.evaluateFormula('pin 0 150 100', {}), 100)
  assert.strictEqual(geometry.evaluateFormula('+- 10 7 3', {}), 14)

  console.log('OOXML geometry tests passed')
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})
