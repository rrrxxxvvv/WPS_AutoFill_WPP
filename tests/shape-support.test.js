const assert = require('assert')
const fs = require('fs')
const path = require('path')

global.JSZip = require('jszip')
const geometry = require('../js/ooxml-geometry')
const fillEngine = require('../js/fill-engine')

const fixturePath = path.join(__dirname, 'fixtures', 'adjusted-shapes.pptx')
const definitionsPath = path.join(__dirname, '..', 'assets', 'presetShapeDefinitions.xml')

async function main() {
  const snapshot = await geometry.snapshotFromBuffer(
    fs.readFileSync(fixturePath),
    2,
    fs.readFileSync(definitionsPath, 'utf8')
  )
  const matrix = []

  Object.entries(snapshot.shapes).forEach(([id, entry]) => {
    const paths = snapshot.geometryForShape({ Id: Number(id) }, 96)
      .filter(shapePath => shapePath.stroke !== false)
    const closed = paths.filter(shapePath => shapePath.closed)
    const open = paths.filter(shapePath => !shapePath.closed)
    const curved = paths.filter(shapePath => shapePath.curved)
    let standaloneRegions = 0
    let regionError = ''
    try {
      standaloneRegions = fillEngine._geometry.buildGenericRegions(paths).length
    } catch (error) {
      regionError = String(error && error.message ? error.message : error)
    }

    closed.filter(shapePath => shapePath.curved).forEach(shapePath => {
      shapePath.points._pathEdges = shapePath.edgeMeta
      const rows = fillEngine._geometry.addCurvePointRows(shapePath.points)
      assert.ok(rows && rows.length >= 4 && (rows.length - 1) % 3 === 0,
        entry.name + ' 的闭合曲边必须能转换成 AddCurve 点列')
      const start = rows[0]
      const end = rows[rows.length - 1]
      assert.ok(Math.hypot(start[0] - end[0], start[1] - end[1]) < 1e-5,
        entry.name + ' 的 AddCurve 点列必须精确闭合')
    })

    matrix.push({
      id: Number(id),
      name: entry.name,
      preset: entry.preset || 'custGeom',
      pathCount: paths.length,
      closedPathCount: closed.length,
      openPathCount: open.length,
      curvedPathCount: curved.length,
      standaloneRegions,
      role: standaloneRegions > 0 ? 'region-capable' : 'divider-only',
      regionError
    })
  })

  assert.strictEqual(matrix.length, 160)
  assert.strictEqual(matrix.filter(item => item.pathCount === 1 && item.closedPathCount === 1).length, 96)
  assert.strictEqual(matrix.filter(item => item.pathCount > 1).length, 38)
  assert.strictEqual(matrix.filter(item => item.curvedPathCount > 0).length, 64)
  assert.strictEqual(matrix.filter(item => item.role === 'region-capable').length, 135,
    '135 个对象单独使用时应形成至少一个候选区域')
  assert.strictEqual(matrix.filter(item => item.role === 'divider-only').length, 25,
    '其余 25 个开放线稿只能与其他边界组合使用')
  assert.ok(matrix.find(item => item.preset === 'noSmoking' && item.pathCount === 3 && item.standaloneRegions > 1),
    '禁止符必须保留三个可参与求交的曲线子路径')

  // 2026-08-14 WPS 实机反馈中的精确对象组合。这里按诊断数据重建形状，
  // 防止再次把“剪角矩形 + 禁止符”误判成单外框，也验证加入七边形后的求交结果。
  const feedbackShapes = [
    ['9004', '剪去同侧角的矩形 3', 'snip2SameRect', 377.95001220703125, 150.9499969482422, 221.14999389648438, 238.14999389648438],
    ['9005', '禁止符 4', 'noSmoking', 236.14999389648438, 31.850000381469727, 317.54998779296875, 272.1499938964844],
    ['9012', '七边形 11', 'heptagon', 542.3499755859375, 122.55000305175781, 283.5, 175.8000030517578]
  ]
  feedbackShapes.forEach(([id, name, preset, left, top, width, height]) => {
    snapshot.shapes[id] = {
      id,
      name,
      preset,
      adjustments: {},
      transform: {
        x: left * 12700,
        y: top * 12700,
        cx: width * 12700,
        cy: height * 12700,
        rotation: 0,
        horizontalFlip: false,
        verticalFlip: false
      },
      inGroup: false,
      customGeometryXml: ''
    }
  })
  const feedbackPaths = ids => ids.flatMap(id => (
    snapshot.geometryForShape({ Id: Number(id) }, 160).filter(shapePath => shapePath.stroke !== false)
  ))
  const twoShapeRegions = fillEngine._geometry.buildGenericRegions(feedbackPaths(['9004', '9005']))
  const threeShapeRegions = fillEngine._geometry.buildGenericRegions(feedbackPaths(['9004', '9005', '9012']))
  assert.strictEqual(twoShapeRegions.length, 7,
    '剪角矩形与禁止符组合应识别 7 个闭合候选区域，而不是 1 个外框')
  assert.strictEqual(threeShapeRegions.length, 9,
    '加入七边形后应识别 9 个闭合候选区域')
  assert.ok(threeShapeRegions.some(region => fillEngine._geometry.addCurvePointRows(region)),
    '实机反馈组合中的曲边候选区域必须能转换为 AddCurve 点列')

  console.log('shape support matrix passed:', {
    total: matrix.length,
    singleClosed: matrix.filter(item => item.pathCount === 1 && item.closedPathCount === 1).length,
    multiPath: matrix.filter(item => item.pathCount > 1).length,
    curved: matrix.filter(item => item.curvedPathCount > 0).length,
    regionCapable: matrix.filter(item => item.role === 'region-capable').length,
    dividerOnly: matrix.filter(item => item.role === 'divider-only').length
  })
}

main().catch(error => {
  console.error(error)
  process.exitCode = 1
})
